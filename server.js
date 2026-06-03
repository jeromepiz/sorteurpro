const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static('.'));

const ONGLETS_A_IGNORER = ['📊 Récapitulatif'];
const DATA_FILE = path.join(__dirname, 'data.json');

// ─── PERSISTANCE JSON ───
function loadData() {
  try {
    if (fs.existsSync(DATA_FILE)) return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
  } catch(e) {}
  return { validations: [], sessions: [] };
}
function saveData(data) {
  try { fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2)); } catch(e) {}
}
let db = loadData();

// ─── LECTURE EXCEL : retourne les tournées avec adresses ───
function readTournees() {
  const wb = XLSX.readFile(path.join(__dirname, 'tournees_sorteurs.xlsx'));
  const tournees = {};
  wb.SheetNames.forEach(name => {
    if (ONGLETS_A_IGNORER.includes(name)) return;
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1 });
    const addresses = [];
    rows.forEach(row => {
      const adresse = row[1];
      if (!adresse || typeof adresse !== 'string' || adresse.trim() === '') return;
      if (adresse.toLowerCase().includes('adresse')) return;
      addresses.push({
        name: adresse.trim(),
        sub: [row[2] ? row[2].toString().trim() : '', row[3] ? row[3] + ' bacs' : ''].filter(Boolean).join(' — ')
      });
    });
    tournees[name] = { label: name, addresses };
  });
  return tournees;
}

// ─── TOURNÉES ───
app.get('/api/tournees', (req, res) => {
  try {
    res.json(readTournees());
  } catch(err) {
    console.error('Erreur Excel:', err);
    res.status(500).json({ error: 'Impossible de lire le fichier Excel.' });
  }
});

// ─── VALIDATION ───
app.post('/api/validation', (req, res) => {
  const entry = { ...req.body, receivedAt: new Date().toISOString() };

  if (entry.type === 'fin_tournee') {
    const idx = db.sessions.findIndex(s =>
      s.agent === entry.agent &&
      s.tourneeId === entry.tourneeId &&
      s.startTime === entry.startTime
    );
    if (idx >= 0) db.sessions[idx] = { ...db.sessions[idx], ...entry, status: 'terminee' };
    else db.sessions.push({ ...entry, status: 'terminee' });

  } else {
    db.validations.push(entry);

    const idx = db.sessions.findIndex(s =>
      s.agent === entry.agent &&
      s.tourneeId === entry.tourneeId &&
      s.startTime === entry.startTime
    );

    if (idx < 0) {
      // Nouvelle session : on récupère le nombre total d'adresses depuis l'Excel
      let totalAdresses = null;
      try {
        const tournees = readTournees();
        if (tournees[entry.tourneeId]) {
          totalAdresses = tournees[entry.tourneeId].addresses.length;
        }
      } catch(e) {}

      db.sessions.push({
        agent:        entry.agent,
        tourneeId:    entry.tourneeId,
        tourneeName:  entry.tourneeName,
        startTime:    entry.startTime,
        status:       'en_cours',
        totalAdresses: totalAdresses,  // ← stocké dès la création
        validations:  [entry]
      });
    } else {
      if (!db.sessions[idx].validations) db.sessions[idx].validations = [];
      db.sessions[idx].validations.push(entry);

      // Si le total n'était pas encore connu, on le résout maintenant
      if (!db.sessions[idx].totalAdresses) {
        try {
          const tournees = readTournees();
          if (tournees[entry.tourneeId]) {
            db.sessions[idx].totalAdresses = tournees[entry.tourneeId].addresses.length;
          }
        } catch(e) {}
      }
    }
  }

  saveData(db);
  res.json({ ok: true });
});

// ─── DASHBOARD ───
app.get('/api/dashboard', (req, res) => {
  const { date } = req.query;
  let sessions    = db.sessions;
  let validations = db.validations;

  if (date) {
    sessions    = sessions.filter(s => s.startTime && s.startTime.startsWith(date));
    validations = validations.filter(v => v.timestamp && v.timestamp.startsWith(date));
  }

  // Résout le total depuis l'Excel pour les sessions qui ne l'ont pas encore
  let tournees = null;
  const sessionsEnrichies = sessions.map(s => {
    const vals      = validations.filter(v => v.agent === s.agent && v.tourneeId === s.tourneeId && v.startTime === s.startTime);
    const anomalies = vals.filter(v => v.anomalie);

    // Priorité : totalAdresses stocké → Excel → null
    let total = s.totalAdresses || s.total || null;
    if (!total) {
      if (!tournees) {
        try { tournees = readTournees(); } catch(e) { tournees = {}; }
      }
      if (tournees[s.tourneeId]) total = tournees[s.tourneeId].addresses.length;
    }

    return {
      ...s,
      total,
      nbValidations: vals.length,
      nbAnomalies:   anomalies.length,
      status:        s.endTime ? 'terminee' : 'en_cours',
      progression:   vals.length
    };
  });

  const stats = {
    tourneesEnCours:   sessionsEnrichies.filter(s => s.status === 'en_cours').length,
    tourneesTerminees: sessionsEnrichies.filter(s => s.status === 'terminee').length,
    totalAnomalies:    validations.filter(v => v.anomalie).length,
    totalAdresses:     validations.length,
  };

  const anomalies = validations.filter(v => v.anomalie).map(v => ({
    agent:        v.agent,
    tourneeId:    v.tourneeId,
    adresse:      v.adresse,
    anomalieType: v.anomalieType,
    commentaire:  v.commentaire,
    timestamp:    v.timestamp
  }));

  res.json({ stats, sessions: sessionsEnrichies, anomalies, validations });
});

// ─── EXPORT EXCEL ───
app.get('/api/export/excel', (req, res) => {
  const { date } = req.query;
  let sessions    = db.sessions;
  let validations = db.validations;
  if (date) {
    sessions    = sessions.filter(s => s.startTime && s.startTime.startsWith(date));
    validations = validations.filter(v => v.timestamp && v.timestamp.startsWith(date));
  }

  const wb = XLSX.utils.book_new();

  const sessData = sessions.map(s => {
    const vals = validations.filter(v => v.agent === s.agent && v.tourneeId === s.tourneeId && v.startTime === s.startTime);
    return {
      'Agent':             s.agent,
      'Tournée':           s.tourneeId,
      'Date':              s.startTime ? new Date(s.startTime).toLocaleDateString('fr-FR') : '',
      'Heure début':       s.startTime ? new Date(s.startTime).toLocaleTimeString('fr-FR') : '',
      'Heure fin':         s.endTime   ? new Date(s.endTime).toLocaleTimeString('fr-FR') : 'En cours',
      'Durée (min)':       s.endTime   ? Math.round((new Date(s.endTime) - new Date(s.startTime)) / 60000) : '',
      'Adresses traitées': vals.length,
      'Total adresses':    s.totalAdresses || s.total || '',
      'Anomalies':         vals.filter(v => v.anomalie).length,
      'Statut':            s.endTime ? 'Terminée' : 'En cours'
    };
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sessData), 'Tournées');

  const anomData = validations.filter(v => v.anomalie).map(v => ({
    'Date':          v.timestamp ? new Date(v.timestamp).toLocaleDateString('fr-FR') : '',
    'Heure':         v.timestamp ? new Date(v.timestamp).toLocaleTimeString('fr-FR') : '',
    'Agent':         v.agent,
    'Tournée':       v.tourneeId,
    'Adresse':       v.adresse,
    'Type anomalie': v.anomalieType || '',
    'Commentaire':   v.commentaire || ''
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(anomData.length ? anomData : [{ 'Info': 'Aucune anomalie' }]), 'Anomalies');

  const detailData = validations.map(v => ({
    'Date':          v.timestamp ? new Date(v.timestamp).toLocaleDateString('fr-FR') : '',
    'Heure':         v.timestamp ? new Date(v.timestamp).toLocaleTimeString('fr-FR') : '',
    'Agent':         v.agent,
    'Tournée':       v.tourneeId,
    'Adresse':       v.adresse,
    'Statut':        v.anomalie ? 'Anomalie' : 'OK',
    'Type anomalie': v.anomalieType || '',
    'Commentaire':   v.commentaire || ''
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detailData.length ? detailData : [{ 'Info': 'Aucune donnée' }]), 'Détail validations');

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const filename = `sorteurpro_${date || 'complet'}.xlsx`;
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// ─── DEBUG ───
app.get('/api/debug', (req, res) => {
  try {
    const wb = XLSX.readFile(path.join(__dirname, 'tournees_sorteurs.xlsx'));
    const debug = {};
    wb.SheetNames.forEach(name => {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1 });
      debug[name] = { nb_lignes: rows.length, ligne_1: rows[0]||null, ligne_2: rows[1]||null, ligne_3: rows[2]||null, ligne_4: rows[3]||null };
    });
    res.json(debug);
  } catch(err) { res.status(500).json({ error: err.message }); }
});

// ─── DÉMARRAGE ───
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('SorteurPro démarré sur le port ' + PORT));
