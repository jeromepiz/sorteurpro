const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const { Pool } = require('pg');

const app = express();
app.use(cors());
app.use(express.json({ limit: '20mb' }));
app.use(express.static('.'));

// ─── POSTGRESQL ───────────────────────────────────────────────────────────────
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

async function initDB() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS sessions (
      id TEXT PRIMARY KEY,
      agent TEXT,
      tournee_id TEXT,
      type_tournee TEXT,
      start_time TIMESTAMPTZ,
      end_time TIMESTAMPTZ,
      created_at TIMESTAMPTZ DEFAULT NOW()
    );
    CREATE TABLE IF NOT EXISTS validations (
      id SERIAL PRIMARY KEY,
      session_id TEXT REFERENCES sessions(id) ON DELETE CASCADE,
      agent TEXT,
      tournee_id TEXT,
      type_tournee TEXT,
      adresse TEXT,
      anomalie BOOLEAN DEFAULT FALSE,
      anomalie_type TEXT,
      commentaire TEXT,
      photo TEXT,
      timestamp TIMESTAMPTZ,
      created_at TIMESTAMPTZ DEFAULT NOW()
    );
    CREATE TABLE IF NOT EXISTS pdf_exports (
      id SERIAL PRIMARY KEY,
      session_id TEXT,
      tournee_id TEXT,
      type_tournee TEXT,
      date_tournee DATE,
      pdf_data TEXT,
      filename TEXT,
      created_at TIMESTAMPTZ DEFAULT NOW()
    );
    ALTER TABLE sessions ADD COLUMN IF NOT EXISTS loc_captured_at TIMESTAMPTZ;
    ALTER TABLE sessions ADD COLUMN IF NOT EXISTS loc_lat DOUBLE PRECISION;
    ALTER TABLE sessions ADD COLUMN IF NOT EXISTS loc_lng DOUBLE PRECISION;
    ALTER TABLE sessions ADD COLUMN IF NOT EXISTS loc_accuracy DOUBLE PRECISION;
    ALTER TABLE validations ADD COLUMN IF NOT EXISTS lat DOUBLE PRECISION;
    ALTER TABLE validations ADD COLUMN IF NOT EXISTS lng DOUBLE PRECISION;
    ALTER TABLE validations ADD COLUMN IF NOT EXISTS accuracy DOUBLE PRECISION;
    ALTER TABLE validations ADD COLUMN IF NOT EXISTS collecte_effectuee BOOLEAN;
  `);
  console.log('✅ Base PostgreSQL initialisée');
}

// ─── ONGLETS EXCEL ────────────────────────────────────────────────────────────
const ONGLETS_A_IGNORER = ['📊 Récapitulatif'];

function readTournees() {
  const wb = XLSX.readFile(path.join(__dirname, 'tournees_sorteurs.xlsx'));
  const tournees = {};
  wb.SheetNames.forEach(name => {
    if (ONGLETS_A_IGNORER.includes(name)) return;
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1 });

    // La commune est indiquée sur une ligne "Commune" / <valeur> (généralement
    // en A2/B2, mais on scanne les premières lignes pour rester robuste si
    // la mise en page d'un onglet varie légèrement).
    let commune = '';
    for (let i = 0; i < Math.min(rows.length, 5); i++) {
      const label = rows[i] && rows[i][0] ? String(rows[i][0]).trim().toLowerCase() : '';
      if (label === 'commune' && rows[i][1]) {
        commune = String(rows[i][1]).trim();
        break;
      }
    }

    const addresses = [];
    rows.forEach(row => {
      const adresse = row[1];
      if (!adresse || typeof adresse !== 'string' || adresse.trim() === '') return;
      if (adresse.toLowerCase().includes('adresse')) return;
      const complement = row[2] ? row[2].toString().trim() : '';
      const nbBacs = row[3] ? row[3] + ' bacs' : '';
      addresses.push({
        name: adresse.trim(),
        sub: [complement, nbBacs].filter(Boolean).join(' — ')
      });
    });
    tournees[name] = { label: name, commune, addresses };
  });
  return tournees;
}

// ─── GÉNÉRATION PDF ───────────────────────────────────────────────────────────
function generateAnomaliesPDF(session, anomalies, withPhotos) {
  const d = new Date(session.end_time || session.start_time);
  const dateStr = d.toLocaleDateString('fr-FR', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
  const heureDebut = new Date(session.start_time).toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });
  const heureFin = session.end_time ? new Date(session.end_time).toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }) : '—';
  const typeBadgeColor = session.type_tournee === 'Sortie' ? '#1a3a6b' : '#6b1a6b';
  const typeBadge = session.type_tournee || 'Non défini';

  const anomaliesHTML = anomalies.length === 0
    ? '<p style="color:#666;font-style:italic;text-align:center;padding:20px;">Aucune anomalie signalée lors de cette tournée.</p>'
    : anomalies.map((a, i) => {
      const heure = new Date(a.timestamp).toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });
      const photoHTML = (withPhotos && a.photo)
        ? `<div style="margin-top:10px;"><img src="${a.photo}" style="max-width:100%;max-height:300px;border-radius:6px;border:1px solid #ddd;" alt="Photo anomalie"/></div>`
        : '';
      const collecteHTML = a.collecte_effectuee === true
        ? `<span style="background:#e8f5e9;color:#1b5e20;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;margin-left:6px;">✅ Collecte effectuée</span>`
        : a.collecte_effectuee === false
        ? `<span style="background:#ffebee;color:#b71c1c;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;margin-left:6px;">❌ Collecte non effectuée</span>`
        : '';
      return `
        <div style="background:#fff8f8;border-left:4px solid #e53e3e;border-radius:6px;padding:14px;margin-bottom:14px;page-break-inside:avoid;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
            <span style="font-weight:700;font-size:15px;color:#1a1a2e;">📍 ${a.adresse}</span>
            <span style="background:#e53e3e;color:#fff;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600;">⚠️ ${a.anomalie_type || 'Anomalie'}</span>
          </div>
          <div style="color:#666;font-size:12px;margin-bottom:6px;">🕐 ${heure}${collecteHTML}</div>
          ${a.commentaire ? `<div style="background:#fff;border:1px solid #fecaca;border-radius:4px;padding:8px;font-size:13px;color:#444;">${a.commentaire}</div>` : ''}
          ${photoHTML}
        </div>`;
    }).join('');

  return `<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8"/>
<title>Anomalies — ${session.tournee_id} — ${dateStr}</title>
<style>
  @page { margin: 20mm 15mm; }
  body { font-family: 'Helvetica Neue', Arial, sans-serif; color: #1a1a2e; margin: 0; padding: 0; }
  .header { background: ${typeBadgeColor}; color: white; padding: 24px; border-radius: 0 0 12px 12px; margin-bottom: 24px; }
  .header h1 { margin: 0 0 4px; font-size: 22px; }
  .header .sub { opacity: 0.85; font-size: 14px; }
  .meta-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-bottom: 24px; }
  .meta-card { background: #f8f9fa; border-radius: 8px; padding: 12px; text-align: center; }
  .meta-card .label { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px; }
  .meta-card .value { font-size: 16px; font-weight: 700; color: #1a1a2e; }
  .section-title { font-size: 16px; font-weight: 700; margin-bottom: 14px; padding-bottom: 8px; border-bottom: 2px solid #e2e8f0; }
  .footer { text-align: center; color: #aaa; font-size: 10px; margin-top: 30px; padding-top: 10px; border-top: 1px solid #e2e8f0; }
</style>
</head>
<body>
  <div class="header">
    <h1>⚠️ Rapport d'anomalies — Tournée ${session.tournee_id}</h1>
    <div class="sub">${dateStr} &nbsp;|&nbsp; ${typeBadge}</div>
  </div>
  <div class="meta-grid">
    <div class="meta-card"><div class="label">Agent</div><div class="value">${session.agent}</div></div>
    <div class="meta-card"><div class="label">Début</div><div class="value">${heureDebut}</div></div>
    <div class="meta-card"><div class="label">Fin</div><div class="value">${heureFin}</div></div>
  </div>
  <div class="meta-grid" style="margin-bottom:24px;">
    <div class="meta-card"><div class="label">Anomalies</div><div class="value" style="color:#e53e3e;">${anomalies.length}</div></div>
    <div class="meta-card"><div class="label">Photos</div><div class="value">${anomalies.filter(a => a.photo).length}</div></div>
    <div class="meta-card"><div class="label">Type</div><div class="value" style="font-size:13px;">${typeBadge}</div></div>
  </div>
  <div class="section-title">🔍 Détail des anomalies</div>
  ${anomaliesHTML}
  <div class="footer">SorteurPro — Généré automatiquement le ${new Date().toLocaleString('fr-FR')} — ${withPhotos ? 'Avec photos' : 'Sans photos'}</div>
</body>
</html>`;
}

// Génère le nom de fichier et structure de dossier pour le PDF
function getPDFFilename(session) {
  const d = new Date(session.end_time || session.start_time);
  const annee = d.getFullYear();
  const mois = String(d.getMonth() + 1).padStart(2, '0');
  const jour = String(d.getDate()).padStart(2, '0');
  const type = (session.type_tournee || 'inconnu').replace(/[^a-zA-Z]/g, '');
  return {
    folder: `${annee}/${mois}/${jour}/${session.tournee_id}`,
    filename: `ANOMALIES_${session.tournee_id}_${type}_${annee}-${mois}-${jour}.pdf`
  };
}

// ─── API : TOURNÉES ──────────────────────────────────────────────────────────
app.get('/api/tournees', (req, res) => {
  try { res.json(readTournees()); }
  catch (err) { console.error(err); res.status(500).json({ error: 'Impossible de lire le fichier Excel.' }); }
});

// ─── API : VALIDATION ────────────────────────────────────────────────────────
app.post('/api/validation', async (req, res) => {
  try {
    const { sessionId, agent, tourneeId, typeTournee, adresse, anomalie, anomalieType, collecteEffectuee, commentaire, photo, timestamp, lat, lng, accuracy } = req.body;

    // Upsert session — met aussi à jour la dernière position connue si le téléphone en a transmis une
    await pool.query(`
      INSERT INTO sessions (id, agent, tournee_id, type_tournee, start_time, loc_lat, loc_lng, loc_accuracy, loc_captured_at)
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, CASE WHEN $6::double precision IS NOT NULL THEN NOW() END)
      ON CONFLICT (id) DO UPDATE SET
        agent = EXCLUDED.agent,
        type_tournee = COALESCE(sessions.type_tournee, EXCLUDED.type_tournee),
        loc_lat = COALESCE(EXCLUDED.loc_lat, sessions.loc_lat),
        loc_lng = COALESCE(EXCLUDED.loc_lng, sessions.loc_lng),
        loc_accuracy = COALESCE(EXCLUDED.loc_accuracy, sessions.loc_accuracy),
        loc_captured_at = CASE WHEN EXCLUDED.loc_lat IS NOT NULL THEN NOW() ELSE sessions.loc_captured_at END
    `, [sessionId, agent, tourneeId, typeTournee, timestamp, lat || null, lng || null, accuracy || null]);

    // Insérer validation (avec la position GPS capturée à ce moment, si disponible)
    await pool.query(`
      INSERT INTO validations (session_id, agent, tournee_id, type_tournee, adresse, anomalie, anomalie_type, commentaire, photo, timestamp, lat, lng, accuracy, collecte_effectuee)
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14)
    `, [sessionId, agent, tourneeId, typeTournee, adresse, anomalie || false, anomalieType || null, commentaire || null, photo || null, timestamp, lat || null, lng || null, accuracy || null, typeof collecteEffectuee === 'boolean' ? collecteEffectuee : null]);

    res.json({ ok: true });
  } catch (err) {
    console.error('Erreur validation:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─── API : FIN DE TOURNÉE (génère PDF auto) ──────────────────────────────────
app.post('/api/fin-tournee', async (req, res) => {
  try {
    const { sessionId, endTime, withPhotos = true } = req.body;

    // Marquer fin de session
    await pool.query(`UPDATE sessions SET end_time = $1 WHERE id = $2`, [endTime, sessionId]);

    // Récupérer session + anomalies
    const sessionRes = await pool.query(`SELECT * FROM sessions WHERE id = $1`, [sessionId]);
    const session = sessionRes.rows[0];
    if (!session) return res.status(404).json({ error: 'Session non trouvée' });

    const anomaliesRes = await pool.query(`
      SELECT * FROM validations
      WHERE session_id = $1 AND anomalie = true
      ORDER BY timestamp ASC
    `, [sessionId]);
    const anomalies = anomaliesRes.rows;

    // Générer PDF HTML
    const pdfHTML = generateAnomaliesPDF(session, anomalies, withPhotos);
    const { folder, filename } = getPDFFilename(session);
    const d = new Date(session.end_time || session.start_time);

    // Sauvegarder en base (HTML qu'on convertira côté client en PDF)
    await pool.query(`
      INSERT INTO pdf_exports (session_id, tournee_id, type_tournee, date_tournee, pdf_data, filename)
      VALUES ($1, $2, $3, $4, $5, $6)
    `, [sessionId, session.tournee_id, session.type_tournee, d.toISOString().split('T')[0], pdfHTML, filename]);

    res.json({
      ok: true,
      anomaliesCount: anomalies.length,
      pdfGenerated: true,
      folder,
      filename
    });
  } catch (err) {
    console.error('Erreur fin-tournee:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─── API : TÉLÉCHARGER PDF anomalies par session ────────────────────────────
app.get('/api/export/anomalies-pdf/:sessionId', async (req, res) => {
  try {
    const { sessionId } = req.params;
    const withPhotos = req.query.photos !== 'false';

    const sessionRes = await pool.query(`SELECT * FROM sessions WHERE id = $1`, [sessionId]);
    const session = sessionRes.rows[0];
    if (!session) return res.status(404).json({ error: 'Session non trouvée' });

    const anomaliesRes = await pool.query(`
      SELECT * FROM validations WHERE session_id = $1 AND anomalie = true ORDER BY timestamp ASC
    `, [sessionId]);

    const pdfHTML = generateAnomaliesPDF(session, anomaliesRes.rows, withPhotos);
    const { filename } = getPDFFilename(session);

    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename.replace('.pdf', '.html')}"`);
    res.send(pdfHTML);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── API : LISTE DES EXPORTS PDF (pour dashboard) ────────────────────────────
app.get('/api/pdf-exports', async (req, res) => {
  try {
    const { date } = req.query;
    let query = `SELECT id, session_id, tournee_id, type_tournee, date_tournee, filename, created_at FROM pdf_exports`;
    const params = [];
    if (date) {
      query += ` WHERE date_tournee = $1`;
      params.push(date);
    }
    query += ` ORDER BY created_at DESC LIMIT 100`;
    const result = await pool.query(query, params);
    res.json(result.rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── API : TÉLÉCHARGER UN PDF DEPUIS LA BASE ────────────────────────────────
app.get('/api/pdf-exports/:id/download', async (req, res) => {
  try {
    const result = await pool.query(`SELECT * FROM pdf_exports WHERE id = $1`, [req.params.id]);
    if (!result.rows[0]) return res.status(404).send('Non trouvé');
    const { pdf_data, filename } = result.rows[0];
    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename.replace('.pdf', '.html')}"`);
    res.send(pdf_data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── API : DASHBOARD ─────────────────────────────────────────────────────────
app.get('/api/dashboard', async (req, res) => {
  try {
    const { date } = req.query;
    const targetDate = date || new Date().toISOString().split('T')[0];

    const sessionsRes = await pool.query(`
      SELECT s.*, COUNT(v.id) as total_validations,
             COUNT(CASE WHEN v.anomalie THEN 1 END) as total_anomalies
      FROM sessions s
      LEFT JOIN validations v ON v.session_id = s.id
      WHERE DATE(s.start_time AT TIME ZONE 'Europe/Paris') = $1
      GROUP BY s.id ORDER BY s.start_time DESC
    `, [targetDate]);

    const anomaliesRes = await pool.query(`
      SELECT v.* FROM validations v
      JOIN sessions s ON s.id = v.session_id
      WHERE DATE(s.start_time AT TIME ZONE 'Europe/Paris') = $1 AND v.anomalie = true
      ORDER BY v.timestamp DESC
    `, [targetDate]);

    const statsRes = await pool.query(`
      SELECT
        COUNT(DISTINCT s.id) as total_sessions,
        COUNT(v.id) as total_validations,
        COUNT(CASE WHEN v.anomalie THEN 1 END) as total_anomalies
      FROM sessions s
      LEFT JOIN validations v ON v.session_id = s.id
      WHERE DATE(s.start_time AT TIME ZONE 'Europe/Paris') = $1
    `, [targetDate]);

    // Charger les tournées pour calculer la progression
    let tournees = {};
    try { tournees = readTournees(); } catch (e) {}

    const sessions = sessionsRes.rows.map(s => {
      const t = tournees[s.tournee_id];
      const totalAddresses = t ? t.addresses.length : 0;
      return { ...s, totalAddresses };
    });

    res.json({
      sessions,
      anomalies: anomaliesRes.rows,
      stats: statsRes.rows[0],
      date: targetDate
    });
  } catch (err) {
    console.error('Erreur dashboard:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─── API : VERSION DASHBOARD (léger, sans photos) ────────────────────────────
// Utilisé pour le polling automatique : renvoie juste de quoi détecter un
// changement (compteurs + dernier timestamp), sans transférer les données
// complètes (ni les photos). Le dashboard ne re-télécharge /api/dashboard en
// entier que si cette empreinte a changé depuis le dernier appel.
app.get('/api/dashboard/version', async (req, res) => {
  try {
    const { date } = req.query;
    const targetDate = date || new Date().toISOString().split('T')[0];

    const r = await pool.query(`
      SELECT
        COUNT(DISTINCT s.id) as session_count,
        COUNT(v.id) as validation_count,
        COUNT(CASE WHEN v.anomalie THEN 1 END) as anomaly_count,
        MAX(v.timestamp) as last_validation,
        MAX(s.end_time) as last_end
      FROM sessions s
      LEFT JOIN validations v ON v.session_id = s.id
      WHERE DATE(s.start_time AT TIME ZONE 'Europe/Paris') = $1
    `, [targetDate]);

    res.json(r.rows[0]);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── API : HISTORIQUE ────────────────────────────────────────────────────────
app.get('/api/history', async (req, res) => {
  try {
    const result = await pool.query(`
      SELECT s.*, COUNT(v.id) as total_validations,
             COUNT(CASE WHEN v.anomalie THEN 1 END) as total_anomalies
      FROM sessions s
      LEFT JOIN validations v ON v.session_id = s.id
      GROUP BY s.id ORDER BY s.start_time DESC LIMIT 200
    `);
    res.json(result.rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── API : DÉTAIL SESSION ────────────────────────────────────────────────────
app.get('/api/session/:sessionId', async (req, res) => {
  try {
    const sRes = await pool.query(`SELECT * FROM sessions WHERE id = $1`, [req.params.sessionId]);
    const vRes = await pool.query(`SELECT * FROM validations WHERE session_id = $1 ORDER BY timestamp ASC`, [req.params.sessionId]);
    if (!sRes.rows[0]) return res.status(404).json({ error: 'Non trouvée' });

    let tournees = {};
    try { tournees = readTournees(); } catch (e) {}
    const t = tournees[sRes.rows[0].tournee_id];
    const session = { ...sRes.rows[0], totalAddresses: t ? t.addresses.length : 0 };

    res.json({ session, validations: vRes.rows });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── UTILITAIRE : DÉCOUPAGE N° DE RUE / ADRESSE ──────────────────────────────
// Exemples : "9 Cour de la République" → { numero: "9", voie: "Cour de la République" }
//            "9-10 Cour de la République" → { numero: "9-10", voie: "Cour de la République" }
//            "Cour de la République" → { numero: "", voie: "Cour de la République" }
function parseAdresse(adresse) {
  if (!adresse) return { numero: '', voie: '' };
  const match = adresse.trim().match(/^(\d+(?:-\d+)?[a-zA-Z]?)\s+(.+)$/);
  if (match) return { numero: match[1], voie: match[2].trim() };
  return { numero: '', voie: adresse.trim() };
}

// Libellés propres (sans emoji) pour l'export Excel — inclut les anciens codes
// pour que les anomalies historiques restent lisibles dans l'export.
const ANOMALY_LABELS_EXPORT = {
  porte_fermee_non_sorti: 'Porte fermée : bac non sorti',
  porte_fermee_non_rentre: 'Porte fermée : bac non rentré',
  absence_bloc_porte: 'Absence de bloc-porte',
  absence_lumiere: 'Absence de lumière',
  local_encombre: 'Local encombré',
  bac_inaccessible: 'Bac inaccessible',
  absence_bacs: 'Absence de bacs',
  effectue_riverain: 'Effectué par riverain',
  local_insalubre: 'Local insalubre / nuisibles',
  squat: 'Squat',
  autre: 'Autre',
  // Anciens codes (avant la mise à jour de la liste)
  bac_absent: 'Bac non sorti', acces_bloque: 'Accès bloqué',
  bac_endommage: 'Bac endommagé', bac_plein: 'Bac trop plein',
  mauvais_tri: 'Mauvais tri', bac_non_rentre: 'Non rentré',
  adresse_absente: 'Adresse introuvable'
};
function anomalyLabelExport(type) {
  return ANOMALY_LABELS_EXPORT[type] || type || '';
}

// Prestataire assurant la collecte — à ajuster ici si besoin (pas encore
// une donnée saisie ailleurs dans l'appli).
const PRESTATAIRE_DEFAUT = 'Pizzorno';

// ─── API : EXPORT EXCEL ANOMALIES ────────────────────────────────────────────
app.get('/api/export/excel', async (req, res) => {
  try {
    const { date, sessions: sessionsParam } = req.query;

    // sessions peut être "id1,id2,id3" ou absent (= toutes)
    const sessionIds = sessionsParam ? sessionsParam.split(',').map(s => s.trim()).filter(Boolean) : null;

    let anomaliesQuery;
    let params = [];

    if (sessionIds && sessionIds.length > 0) {
      // Filtre sur les sessions sélectionnées
      const placeholders = sessionIds.map((_, i) => `$${i + 1}`).join(',');
      anomaliesQuery = `
        SELECT v.*, s.start_time as session_start, s.end_time as session_end
        FROM validations v
        JOIN sessions s ON s.id = v.session_id
        WHERE v.session_id IN (${placeholders}) AND v.anomalie = true
        ORDER BY v.timestamp ASC
      `;
      params = sessionIds;
    } else if (date) {
      anomaliesQuery = `
        SELECT v.*, s.start_time as session_start, s.end_time as session_end
        FROM validations v
        JOIN sessions s ON s.id = v.session_id
        WHERE DATE(s.start_time AT TIME ZONE 'Europe/Paris') = $1 AND v.anomalie = true
        ORDER BY v.timestamp ASC
      `;
      params = [date];
    } else {
      anomaliesQuery = `
        SELECT v.*, s.start_time as session_start, s.end_time as session_end
        FROM validations v
        JOIN sessions s ON s.id = v.session_id
        WHERE v.anomalie = true
        ORDER BY v.timestamp ASC
      `;
    }

    const anomalies = (await pool.query(anomaliesQuery, params)).rows;

    // Commune par tournée (lue depuis tournees_sorteurs.xlsx)
    let tournees = {};
    try { tournees = readTournees(); } catch (e) { console.error('Erreur lecture communes:', e); }

    // ── Construction de la feuille (format calé sur Exemple_extraction.xlsx) ──
    const wb = XLSX.utils.book_new();

    // En-têtes A→J identiques à l'exemple, + 2 colonnes bonus en fin de
    // tableau (K, L) pour ne pas perdre la traçabilité agent / sortie-rentrée.
    const headers = [
      'Date',              // A
      'Heure',             // B
      'N° circuit',        // C
      'Commune',           // D (vide pour l'instant, cf. remarque)
      'N° rue',            // E
      'Adresse',           // F
      'Type anomalie',     // G
      'Complément',        // H
      'Collecte',          // I  O / N
      'Prestataire',       // J
      'Sortie / Rentrée',  // K (bonus, hors format exemple)
      'Nom Prénom sorteur' // L (bonus, hors format exemple)
    ];

    const rows = anomalies.map(a => {
      const ts = a.timestamp ? new Date(a.timestamp) : null;
      const { numero, voie } = parseAdresse(a.adresse);
      // N° rue : nombre pur si possible (ex. "9"), sinon texte tel quel (ex. "9-10")
      const numeroCell = /^\d+$/.test(numero) ? Number(numero) : numero;
      const collecteCode = a.collecte_effectuee === true ? 'O' : a.collecte_effectuee === false ? 'N' : '';
      const commune = (tournees[a.tournee_id] && tournees[a.tournee_id].commune) || '';
      return [
        ts,                                   // A - Date (valeur date réelle)
        ts,                                   // B - Heure (valeur heure réelle)
        a.tournee_id || '',                   // C - N° circuit
        commune,                              // D - Commune
        numeroCell,                           // E - N° rue
        voie,                                 // F - Adresse
        anomalyLabelExport(a.anomalie_type),  // G - Type anomalie
        a.commentaire || '',                  // H - Complément
        collecteCode,                         // I - Collecte (O/N)
        PRESTATAIRE_DEFAUT,                   // J - Prestataire
        a.type_tournee || '',                 // K - Sortie / Rentrée (bonus)
        a.agent || ''                         // L - Nom Prénom sorteur (bonus)
      ];
    });

    // Construire la feuille manuellement (header + données)
    const wsData = [headers, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(wsData, { cellDates: true });

    // Format des colonnes Date / Heure en valeurs Excel réelles (triables, filtrables)
    for (let r = 0; r < rows.length; r++) {
      const rowNum = r + 2; // +1 pour l'en-tête, +1 car 1-indexé
      const dateCell = ws[`A${rowNum}`];
      const heureCell = ws[`B${rowNum}`];
      if (dateCell && dateCell.v) dateCell.z = 'dd/mm/yyyy';
      if (heureCell && heureCell.v) heureCell.z = 'hh:mm';
    }

    // Largeurs de colonnes
    ws['!cols'] = [
      { wch: 12 }, // A Date
      { wch: 8  }, // B Heure
      { wch: 12 }, // C N° circuit
      { wch: 16 }, // D Commune
      { wch: 8  }, // E N° rue
      { wch: 32 }, // F Adresse
      { wch: 26 }, // G Type anomalie
      { wch: 30 }, // H Complément
      { wch: 10 }, // I Collecte
      { wch: 16 }, // J Prestataire
      { wch: 14 }, // K Sortie/Rentrée
      { wch: 24 }, // L Nom Prénom
    ];

    XLSX.utils.book_append_sheet(wb, ws, 'Anomalies');

    // Nom du fichier
    const fileSuffix = date || (sessionIds ? `selection` : 'complet');
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Disposition', `attachment; filename="anomalies_${fileSuffix}.xlsx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch (err) {
    console.error('Erreur export excel:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─── API : LOCALISATION ───────────────────────────────────────────────────────
// La position est capturée passivement à chaque validation d'adresse (voir /api/validation).
// Cet endpoint renvoie simplement la dernière position connue, instantanément, sans attente.
app.get('/api/location/status/:sessionId', async (req, res) => {
  try {
    const r = await pool.query(`
      SELECT loc_captured_at, loc_lat, loc_lng, loc_accuracy
      FROM sessions WHERE id = $1
    `, [req.params.sessionId]);
    if (!r.rows[0]) return res.status(404).json({ error: 'Non trouvée' });
    res.json(r.rows[0]);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── API : DEBUG ─────────────────────────────────────────────────────────────
app.get('/api/debug', (req, res) => {
  try {
    const wb = XLSX.readFile(path.join(__dirname, 'tournees_sorteurs.xlsx'));
    const debug = {};
    wb.SheetNames.forEach(name => {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1 });
      debug[name] = { nb_lignes: rows.length, ligne_1: rows[0] || null, ligne_2: rows[1] || null, ligne_3: rows[2] || null };
    });
    res.json(debug);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ─── DÉMARRAGE ───────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
initDB().then(() => {
  app.listen(PORT, () => console.log(`🚀 SorteurPro démarré sur le port ${PORT}`));
}).catch(err => {
  console.error('❌ Erreur init DB:', err);
  process.exit(1);
});
