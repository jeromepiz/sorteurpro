const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const { Pool } = require('pg');
const app = express();
app.use(cors());
app.use(express.json({ limit: '20mb' }));
app.use(express.static('.'));
// ─── POSTGRESQL
────────────────────────────────────
───────────────────────────
const pool = new Pool({
connectionString:
process.env.DATABASE_URL,
ssl: process.env.DATABASE_URL ? {
rejectUnauthorized: false } : false
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
CREATE TABLE IF NOT EXISTS validations
(
id SERIAL PRIMARY KEY,
session_id TEXT REFERENCES
sessions(id) ON DELETE CASCADE,
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
CREATE TABLE IF NOT EXISTS pdf_exports
(
id SERIAL PRIMARY KEY,
session_id TEXT,
tournee_id TEXT,
type_tournee TEXT,
date_tournee DATE,
pdf_data TEXT,
filename TEXT,
created_at TIMESTAMPTZ DEFAULT NOW()
);
ALTER TABLE sessions ADD COLUMN IF NOT
EXISTS loc_captured_at TIMESTAMPTZ;
ALTER TABLE sessions ADD COLUMN IF NOT
EXISTS loc_lat DOUBLE PRECISION;
ALTER TABLE sessions ADD COLUMN IF NOT
EXISTS loc_lng DOUBLE PRECISION;
ALTER TABLE sessions ADD COLUMN IF NOT
EXISTS loc_accuracy DOUBLE PRECISION;
ALTER TABLE validations ADD COLUMN IF
NOT EXISTS lat DOUBLE PRECISION;
ALTER TABLE validations ADD COLUMN IF
NOT EXISTS lng DOUBLE PRECISION;
ALTER TABLE validations ADD COLUMN IF
NOT EXISTS accuracy DOUBLE PRECISION;
ALTER TABLE validations ADD COLUMN IF
NOT EXISTS collecte_effectuee BOOLEAN;
ALTER TABLE validations ADD COLUMN IF
NOT EXISTS anomalie_precision TEXT;
`);
console.log('✅ Base PostgreSQL
initialisée');
}
// ─── ONGLETS EXCEL
────────────────────────────────────
────────────────────────
const ONGLETS_A_IGNORER = ['📊
Récapitulatif'];
function readTournees() {
const wb =
XLSX.readFile(path.join(__dirname,
'tournees_sorteurs.xlsx'));
const tournees = {};
wb.SheetNames.forEach(name => {
if (ONGLETS_A_IGNORER.includes(name))
return;
const rows =
XLSX.utils.sheet_to_json(wb.Sheets[name], {
header: 1 });
// La commune est indiquée sur une
ligne "Commune" / <valeur> (généralement
// en A2/B2, mais on scanne les
premières lignes pour rester robuste si
// la mise en page d'un onglet varie
légèrement).
let commune = '';
for (let i = 0; i <
Math.min(rows.length, 5); i++) {
const label = rows[i] && rows[i][0] ?
String(rows[i][0]).trim().toLowerCase() :
'';
if (label === 'commune' && rows[i]
[1]) {
commune = String(rows[i]
[1]).trim();
break;
}
}
// Les adresses commencent à la ligne 4
(ligne 1 = titre, ligne 2 = commune,
// ligne 3 = en-têtes de colonnes). On
ignore donc les 3 premières lignes
// pour ne pas lire la valeur de
"Commune" (ex. "Villeurbanne") comme
adresse.
const addresses = [];
rows.slice(3).forEach(row => {
const adresse = row[1];
if (!adresse || typeof adresse !==
'string' || adresse.trim() === '') return;
 if
(adresse.toLowerCase().includes('adresse'))
return;
const complement = row[2] ?
row[2].toString().trim() : '';
const nbBacs = row[3] ? row[3] + '
bacs' : '';
addresses.push({
name: adresse.trim(),
sub: [complement,
nbBacs].filter(Boolean).join(' — ')
});
});
// Tableau N°SDA (colonnes H, I, J) :
"Jour de la semaine" → "N°SDA".
// On scanne toutes les lignes plutôt
que de figer H1:J6 en dur, pour
// rester robuste si le nombre de jours
renseignés varie d'un onglet à l'autre.
const sdaParJour = {};
rows.forEach(row => {
const jour = row[8] ?
String(row[8]).trim().toLowerCase() : '';
const sda = row[9] ?
String(row[9]).trim() : '';
if (jour && sda && jour !== 'jour de
la semaine') {
sdaParJour[jour] = sda;
}
});
tournees[name] = { label: name,
commune, addresses, sdaParJour };
});
 return tournees;
}
// ─── LIBELLÉS ANOMALIES (arbre à 2
niveaux + repli anciens formats)
──────────
const CATEGORY_LABELS = {
sacs_vrac: 'Sacs, vrac à côté du bac',
service_incomplet: 'Pb de service
complet',
tri_mal_trie: 'Bac de tri mal trié',
mauvais_contenu_gris: 'Mauvais contenu du
bac gris',
bac_om_casse: 'Bac OM cassé',
bac_tri_casse: 'Bac Tri cassé',
pb_bac_autre: 'Pb de bac autre',
autre: 'Autre'
};
const PRECISION_LABELS = {
sacs:'Sacs', vrac:'Vrac',
sacs_et_vrac:'Sacs et vrac',
encombrants:'Encombrants',
volume_important:'Volume important',
dechets_dangereux:'Déchets dangereux',
sapin:'Sapin',
porte_fermee_non_sorti:'Porte fermée :
bac non sorti',
porte_fermee_non_rentre:'Porte fermée : bac
non rentré',
absence_bloc_porte:'Absence de bloc-
porte', absence_lumiere:'Absence de
lumière',
local_encombre:'Local encombré',
bac_inaccessible:'Bac inaccessible',
absence_bacs:'Absence de bacs',
effectue_riverain:'Effectué par
riverain', local_insalubre:'Local insalubre
/ nuisibles', squat:'Squat',
deja_scotche:'Déjà scotché',
sacs_fermes:'Sacs fermés', verre:'Verre',
bois_vegetaux:'Bois, végétaux',
textiles:'Textiles',
dechets_alimentaires:'Déchets
alimentaires', gravats:'Gravats',
dechets_animaux:"Déchets d'animaux",
dechets_non_menagers:'Déchets non
ménagers', dechets_verts:'Déchets verts',
cuve:'Cuve', couvercle:'Couvercle',
roue:'Roue', roues:'Roues',
collerette:'Collerette',
debordant:'Débordant', a_rentrer:'A
rentrer', trop_lourd:'Trop lourd',
tres_sale:'Très sale',
trop_tasse:'Trop tassé',
non_identifie:'Non identifié',
trop_bacs_surlitrage:'Trop de bacs
(surlitrage)',
modele_hors_norme:'Modèle hors norme',
bac_750l:'Bac de 750L',
nuisibles:'Nuisibles',
panne:'Panne', accident:'Accident',
altercation:'Altercation',
agression:'Agression',
autre:'Autre'
};
// Très anciens codes (première version de
l'appli, avant la liste à 10 items puis
l'arbre à 2 niveaux)
const OLD_CODES = {
bac_absent:'Bac non sorti (ancien)',
acces_bloque:'Accès bloqué (ancien)',
bac_endommage:'Bac endommagé (ancien)',
bac_plein:'Bac trop plein (ancien)',
mauvais_tri:'Mauvais tri (ancien)',
bac_non_rentre:'Non rentré (ancien)',
adresse_absente:'Adresse introuvable
(ancien)'
};
function anomalyFullLabel(type, precision)
{
if (precision &&
PRECISION_LABELS[precision]) {
const catLabel = CATEGORY_LABELS[type]
|| type;
return `${catLabel} -
${PRECISION_LABELS[precision]}`;
}
if (CATEGORY_LABELS[type]) return
CATEGORY_LABELS[type];
if (PRECISION_LABELS[type]) return
PRECISION_LABELS[type];
if (OLD_CODES[type]) return
OLD_CODES[type];
return type || '';
}
// Pour l'export Excel : "Type anomalie"
(niveau 1 seul) et "Complément"
// (niveau 2 + commentaire éventuel à la
suite, ex. "Autre - Cafards").
function anomalyTypeLabel(type) {
if (CATEGORY_LABELS[type]) return
CATEGORY_LABELS[type];
// Anciennes anomalies enregistrées avant
l'arbre à 2 niveaux : le "type"
// stocké correspond alors à ce qui est
aujourd'hui une précision, ou à un
// tout ancien code — on l'affiche quand
même en "Type anomalie" à défaut.
if (PRECISION_LABELS[type]) return
PRECISION_LABELS[type];
if (OLD_CODES[type]) return
OLD_CODES[type];
return type || '';
}
function anomalyComplementLabel(precision,
commentaire) {
const precisionLabel = (precision &&
PRECISION_LABELS[precision]) ?
PRECISION_LABELS[precision] : '';
if (commentaire) return precisionLabel ?
`${precisionLabel} - ${commentaire}` :
commentaire;
return precisionLabel;
}
// ─── GÉNÉRATION PDF
────────────────────────────────────
───────────────────────
function generateAnomaliesPDF(session,
anomalies, withPhotos) {
const d = new Date(session.end_time ||
session.start_time);
const dateStr = d.toLocaleDateString('fr-
FR', { weekday: 'long', year: 'numeric',
month: 'long', day: 'numeric' });
 const heureDebut = new
Date(session.start_time).toLocaleTimeString
('fr-FR', { hour: '2-digit', minute: '2-
digit' });
const heureFin = session.end_time ? new
Date(session.end_time).toLocaleTimeString('
fr-FR', { hour: '2-digit', minute: '2-
digit' }) : '—';
const typeBadgeColor =
session.type_tournee === 'Sortie' ?
'#1a3a6b' : '#6b1a6b';
const typeBadge = session.type_tournee ||
'Non défini';
const anomaliesHTML = anomalies.length
=== 0
? '<p style="color:#666;font-
style:italic;text-
align:center;padding:20px;">Aucune anomalie
signalée lors de cette tournée.</p>'
: anomalies.map((a, i) => {
const heure = new
Date(a.timestamp).toLocaleTimeString('fr-
FR', { hour: '2-digit', minute: '2-digit'
});
const photoHTML = (withPhotos &&
a.photo)
? `<div style="margin-top:10px;">
<img src="${a.photo}" style="max-
width:100%;max-height:300px;border-
radius:6px;border:1px solid #ddd;"
alt="Photo anomalie"/></div>`
: '';
const collecteHTML =
a.collecte_effectuee === true
? `<span
style="background:#e8f5e9;color:#1b5e20;pad
ding:2px 8px;border-radius:10px;font-
size:11px;font-weight:600;margin-
left:6px;">✅ Collecte effectuée</span>`
: a.collecte_effectuee === false
? `<span
style="background:#ffebee;color:#b71c1c;pad
ding:2px 8px;border-radius:10px;font-
size:11px;font-weight:600;margin-
left:6px;">❌ Collecte non
effectuée</span>`
: '';
return `
<div
style="background:#fff8f8;border-left:4px
solid #e53e3e;border-
radius:6px;padding:14px;margin-
bottom:14px;page-break-inside:avoid;">
<div style="display:flex;justify-
content:space-between;align-
items:center;margin-bottom:6px;">
<span style="font-
weight:700;font-size:15px;color:#1a1a2e;">
📍 ${a.adresse}</span>
<span
style="background:#e53e3e;color:#fff;paddin
g:2px 10px;border-radius:12px;font-
size:12px;font-weight:600;">⚠
${anomalyFullLabel(a.anomalie_type,
a.anomalie_precision)}</span>
</div>
<div style="color:#666;font-
size:12px;margin-bottom:6px;">🕐
${heure}${collecteHTML}</div>
${a.commentaire ? `<div
style="background:#fff;border:1px solid
#fecaca;border-radius:4px;padding:8px;font-
size:13px;color:#444;">${a.commentaire}
</div>` : ''}
${photoHTML}
</div>`;
}).join('');
return `<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8"/>
<title>Anomalies — ${session.tournee_id} —
${dateStr}</title>
<style>
@page { margin: 20mm 15mm; }
body { font-family: 'Helvetica Neue',
Arial, sans-serif; color: #1a1a2e; margin:
0; padding: 0; }
.header { background: ${typeBadgeColor};
color: white; padding: 24px; border-radius:
0 0 12px 12px; margin-bottom: 24px; }
.header h1 { margin: 0 0 4px; font-size:
22px; }
.header .sub { opacity: 0.85; font-size:
14px; }
.meta-grid { display: grid; grid-
template-columns: repeat(3, 1fr); gap:
12px; margin-bottom: 24px; }
.meta-card { background: #f8f9fa; border-
radius: 8px; padding: 12px; text-align:
center; }
.meta-card .label { font-size: 11px;
color: #888; text-transform: uppercase;
letter-spacing: 0.5px; margin-bottom: 4px;
}
.meta-card .value { font-size: 16px;
font-weight: 700; color: #1a1a2e; }
.section-title { font-size: 16px; font-
weight: 700; margin-bottom: 14px; padding-
bottom: 8px; border-bottom: 2px solid
#e2e8f0; }
.footer { text-align: center; color:
#aaa; font-size: 10px; margin-top: 30px;
padding-top: 10px; border-top: 1px solid
#e2e8f0; }
</style>
</head>
<body>
<div class="header">
<h1>⚠ Rapport d'anomalies — Tournée
${session.tournee_id}</h1>
<div class="sub">${dateStr}
&nbsp;|&nbsp; ${typeBadge}</div>
</div>
<div class="meta-grid">
<div class="meta-card"><div
class="label">Agent</div><div
class="value">${session.agent}</div></div>
<div class="meta-card"><div
class="label">Début</div><div
class="value">${heureDebut}</div></div>
<div class="meta-card"><div
class="label">Fin</div><div
class="value">${heureFin}</div></div>
 </div>
<div class="meta-grid" style="margin-
bottom:24px;">
<div class="meta-card"><div
class="label">Anomalies</div><div
class="value"
style="color:#e53e3e;">${anomalies.length}
</div></div>
<div class="meta-card"><div
class="label">Photos</div><div
class="value">${anomalies.filter(a =>
a.photo).length}</div></div>
<div class="meta-card"><div
class="label">Type</div><div class="value"
style="font-size:13px;">${typeBadge}</div>
</div>
</div>
<div class="section-title">🔍 Détail des
anomalies</div>
${anomaliesHTML}
<div class="footer">SorteurPro — Généré
automatiquement le ${new
Date().toLocaleString('fr-FR')} —
${withPhotos ? 'Avec photos' : 'Sans
photos'}</div>
</body>
</html>`;
}
// Génère le nom de fichier et structure de
dossier pour le PDF
function getPDFFilename(session) {
const d = new Date(session.end_time ||
session.start_time);
 const annee = d.getFullYear();
const mois = String(d.getMonth() +
1).padStart(2, '0');
const jour =
String(d.getDate()).padStart(2, '0');
const type = (session.type_tournee ||
'inconnu').replace(/[^a-zA-Z]/g, '');
return {
folder:
`${annee}/${mois}/${jour}/${session.tournee
_id}`,
filename:
`ANOMALIES_${session.tournee_id}_${type}_${
annee}-${mois}-${jour}.pdf`
};
}
// ─── API : TOURNÉES
────────────────────────────────────
──────────────────────
app.get('/api/tournees', (req, res) => {
try { res.json(readTournees()); }
catch (err) { console.error(err);
res.status(500).json({ error: 'Impossible
de lire le fichier Excel.' }); }
});
// ─── API : VALIDATION
────────────────────────────────────
────────────────────
app.post('/api/validation', async (req,
res) => {
try {
const { sessionId, agent, tourneeId,
typeTournee, adresse, anomalie,
anomalieType, anomaliePrecision,
collecteEffectuee, commentaire, photo,
timestamp, lat, lng, accuracy } = req.body;
// Sécurité côté serveur : on retronque
à 60 caractères même si le champ
// libre du navigateur a déjà cette
limite (ex. ancienne version de l'app
// en cache, ou appel direct à l'API).
const agentSafe = agent ?
String(agent).slice(0, 60) : agent;
const commentaireSafe = commentaire ?
String(commentaire).slice(0, 60) :
commentaire;
// Upsert session — met aussi à jour la
dernière position connue si le téléphone en
a transmis une
await pool.query(`
INSERT INTO sessions (id, agent,
tournee_id, type_tournee, start_time,
loc_lat, loc_lng, loc_accuracy,
loc_captured_at)
VALUES ($1, $2, $3, $4, $5, $6, $7,
$8, CASE WHEN $6::double precision IS NOT
NULL THEN NOW() END)
ON CONFLICT (id) DO UPDATE SET
agent = EXCLUDED.agent,
type_tournee =
COALESCE(sessions.type_tournee,
EXCLUDED.type_tournee),
loc_lat =
COALESCE(EXCLUDED.loc_lat,
sessions.loc_lat),
loc_lng =
COALESCE(EXCLUDED.loc_lng,
sessions.loc_lng),
loc_accuracy =
COALESCE(EXCLUDED.loc_accuracy,
sessions.loc_accuracy),
loc_captured_at = CASE WHEN
EXCLUDED.loc_lat IS NOT NULL THEN NOW()
ELSE sessions.loc_captured_at END
`, [sessionId, agentSafe, tourneeId,
typeTournee, timestamp, lat || null, lng ||
null, accuracy || null]);
// Insérer validation (avec la position
GPS capturée à ce moment, si disponible)
await pool.query(`
INSERT INTO validations (session_id,
agent, tournee_id, type_tournee, adresse,
anomalie, anomalie_type, commentaire,
photo, timestamp, lat, lng, accuracy,
collecte_effectuee, anomalie_precision)
VALUES ($1, $2, $3, $4, $5, $6, $7,
$8, $9, $10, $11, $12, $13, $14, $15)
`, [sessionId, agentSafe, tourneeId,
typeTournee, adresse, anomalie || false,
anomalieType || null, commentaireSafe ||
null, photo || null, timestamp, lat ||
null, lng || null, accuracy || null, typeof
collecteEffectuee === 'boolean' ?
collecteEffectuee : null, anomaliePrecision
|| null]);
res.json({ ok: true });
 } catch (err) {
console.error('Erreur validation:',
err);
res.status(500).json({ error:
err.message });
}
});
// ─── API : FIN DE TOURNÉE (génère PDF
auto)
──────────────────────────────────
app.post('/api/fin-tournee', async (req,
res) => {
try {
const { sessionId, endTime, withPhotos
= true } = req.body;
// Marquer fin de session
await pool.query(`UPDATE sessions SET
end_time = $1 WHERE id = $2`, [endTime,
sessionId]);
// Récupérer session + anomalies
const sessionRes = await
pool.query(`SELECT * FROM sessions WHERE id
= $1`, [sessionId]);
const session = sessionRes.rows[0];
if (!session) return
res.status(404).json({ error: 'Session non
trouvée' });
const anomaliesRes = await pool.query(`
SELECT * FROM validations
WHERE session_id = $1 AND anomalie =
true
ORDER BY timestamp ASC
`, [sessionId]);
const anomalies = anomaliesRes.rows;
// Générer PDF HTML
const pdfHTML =
generateAnomaliesPDF(session, anomalies,
withPhotos);
const { folder, filename } =
getPDFFilename(session);
const d = new Date(session.end_time ||
session.start_time);
// Sauvegarder en base (HTML qu'on
convertira côté client en PDF)
await pool.query(`
INSERT INTO pdf_exports (session_id,
tournee_id, type_tournee, date_tournee,
pdf_data, filename)
VALUES ($1, $2, $3, $4, $5, $6)
`, [sessionId, session.tournee_id,
session.type_tournee,
d.toISOString().split('T')[0], pdfHTML,
filename]);
res.json({
ok: true,
anomaliesCount: anomalies.length,
pdfGenerated: true,
folder,
filename
});
} catch (err) {
 console.error('Erreur fin-tournee:',
err);
res.status(500).json({ error:
err.message });
}
});
// ─── API : TÉLÉCHARGER PDF anomalies par
session ────────────────────────────
app.get('/api/export/anomalies-
pdf/:sessionId', async (req, res) => {
try {
const { sessionId } = req.params;
const withPhotos = req.query.photos !==
'false';
const sessionRes = await
pool.query(`SELECT * FROM sessions WHERE id
= $1`, [sessionId]);
const session = sessionRes.rows[0];
if (!session) return
res.status(404).json({ error: 'Session non
trouvée' });
const anomaliesRes = await pool.query(`
SELECT * FROM validations WHERE
session_id = $1 AND anomalie = true ORDER
BY timestamp ASC
`, [sessionId]);
const pdfHTML =
generateAnomaliesPDF(session,
anomaliesRes.rows, withPhotos);
const { filename } =
getPDFFilename(session);
res.setHeader('Content-Type',
'text/html; charset=utf-8');
res.setHeader('Content-Disposition',
`attachment;
filename="${filename.replace('.pdf',
'.html')}"`);
res.send(pdfHTML);
} catch (err) {
res.status(500).json({ error:
err.message });
}
});
// ─── API : LISTE DES EXPORTS PDF (pour
dashboard)
────────────────────────────
app.get('/api/pdf-exports', async (req,
res) => {
try {
const { date } = req.query;
let query = `SELECT id, session_id,
tournee_id, type_tournee, date_tournee,
filename, created_at FROM pdf_exports`;
const params = [];
if (date) {
query += ` WHERE date_tournee = $1`;
params.push(date);
}
query += ` ORDER BY created_at DESC
LIMIT 100`;
const result = await pool.query(query,
params);
 res.json(result.rows);
} catch (err) {
res.status(500).json({ error:
err.message });
}
});
// ─── API : TÉLÉCHARGER UN PDF DEPUIS LA
BASE ────────────────────────────────
app.get('/api/pdf-exports/:id/download',
async (req, res) => {
try {
const result = await pool.query(`SELECT
* FROM pdf_exports WHERE id = $1`,
[req.params.id]);
if (!result.rows[0]) return
res.status(404).send('Non trouvé');
const { pdf_data, filename } =
result.rows[0];
res.setHeader('Content-Type',
'text/html; charset=utf-8');
res.setHeader('Content-Disposition',
`attachment;
filename="${filename.replace('.pdf',
'.html')}"`);
res.send(pdf_data);
} catch (err) {
res.status(500).json({ error:
err.message });
}
});
// ─── API : DASHBOARD
────────────────────────────────────
─────────────────────
app.get('/api/dashboard', async (req, res)
=> {
try {
const { date } = req.query;
const targetDate = date || new
Date().toISOString().split('T')[0];
const sessionsRes = await pool.query(`
SELECT s.*, COUNT(v.id) as
total_validations,
COUNT(CASE WHEN v.anomalie
THEN 1 END) as total_anomalies
FROM sessions s
LEFT JOIN validations v ON
v.session_id = s.id
WHERE DATE(s.start_time AT TIME ZONE
'Europe/Paris') = $1
GROUP BY s.id ORDER BY s.start_time
DESC
`, [targetDate]);
const anomaliesRes = await pool.query(`
SELECT v.* FROM validations v
JOIN sessions s ON s.id =
v.session_id
WHERE DATE(s.start_time AT TIME ZONE
'Europe/Paris') = $1 AND v.anomalie = true
ORDER BY v.timestamp DESC
`, [targetDate]);
const statsRes = await pool.query(`
SELECT
COUNT(DISTINCT s.id) as
total_sessions,
COUNT(v.id) as total_validations,
COUNT(CASE WHEN v.anomalie THEN 1
END) as total_anomalies
FROM sessions s
LEFT JOIN validations v ON
v.session_id = s.id
WHERE DATE(s.start_time AT TIME ZONE
'Europe/Paris') = $1
`, [targetDate]);
// Charger les tournées pour calculer
la progression
let tournees = {};
try { tournees = readTournees(); }
catch (e) {}
const sessions = sessionsRes.rows.map(s
=> {
const t = tournees[s.tournee_id];
const totalAddresses = t ?
t.addresses.length : 0;
return { ...s, totalAddresses };
});
res.json({
sessions,
anomalies: anomaliesRes.rows,
stats: statsRes.rows[0],
date: targetDate
});
} catch (err) {
console.error('Erreur dashboard:',
err);
 res.status(500).json({ error:
err.message });
}
});
// ─── API : VERSION DASHBOARD (léger,
sans photos)
────────────────────────────
app.get('/api/dashboard/version', async
(req, res) => {
try {
const { date } = req.query;
const targetDate = date || new
Date().toISOString().split('T')[0];
const r = await pool.query(`
SELECT
COUNT(DISTINCT s.id) as
session_count,
COUNT(v.id) as validation_count,
COUNT(CASE WHEN v.anomalie THEN 1
END) as anomaly_count,
MAX(v.timestamp) as
last_validation,
MAX(s.end_time) as last_end
FROM sessions s
LEFT JOIN validations v ON
v.session_id = s.id
WHERE DATE(s.start_time AT TIME ZONE
'Europe/Paris') = $1
`, [targetDate]);
res.json(r.rows[0]);
} catch (err) {
 res.status(500).json({ error:
err.message });
}
});
// ─── API : HISTORIQUE
────────────────────────────────────
────────────────────
app.get('/api/history', async (req, res) =>
{
try {
const result = await pool.query(`
SELECT s.*, COUNT(v.id) as
total_validations,
COUNT(CASE WHEN v.anomalie
THEN 1 END) as total_anomalies
FROM sessions s
LEFT JOIN validations v ON
v.session_id = s.id
GROUP BY s.id ORDER BY s.start_time
DESC LIMIT 200
`);
res.json(result.rows);
} catch (err) {
res.status(500).json({ error:
err.message });
}
});
// ─── API : DÉTAIL SESSION
────────────────────────────────────
────────────────
app.get('/api/session/:sessionId', async
(req, res) => {
 try {
const sRes = await pool.query(`SELECT *
FROM sessions WHERE id = $1`,
[req.params.sessionId]);
const vRes = await pool.query(`SELECT *
FROM validations WHERE session_id = $1
ORDER BY timestamp ASC`,
[req.params.sessionId]);
if (!sRes.rows[0]) return
res.status(404).json({ error: 'Non trouvée'
});
let tournees = {};
try { tournees = readTournees(); }
catch (e) {}
const t =
tournees[sRes.rows[0].tournee_id];
const session = { ...sRes.rows[0],
totalAddresses: t ? t.addresses.length : 0
};
res.json({ session, validations:
vRes.rows });
} catch (err) {
res.status(500).json({ error:
err.message });
}
});
// ─── UTILITAIRE : DÉCOUPAGE N° DE RUE /
ADRESSE
──────────────────────────────
// Exemples : "9 Cour de la République" →
{ numero: "9", voie: "Cour de la
République" }
// "9-10 Cour de la République"
→ { numero: "9-10", voie: "Cour de la
République" }
// "Cour de la République" → {
numero: "", voie: "Cour de la République" }
function parseAdresse(adresse) {
if (!adresse) return { numero: '', voie:
'' };
const match = adresse.trim().match(/^(\d+
(?:-\d+)?[a-zA-Z]?)\s+(.+)$/);
if (match) return { numero: match[1],
voie: match[2].trim() };
return { numero: '', voie: adresse.trim()
};
}
// ─── API : EXPORT EXCEL ANOMALIES
────────────────────────────────────
────────
app.get('/api/export/excel', async (req,
res) => {
try {
const { date, sessions: sessionsParam }
= req.query;
// sessions peut être "id1,id2,id3" ou
absent (= toutes)
const sessionIds = sessionsParam ?
sessionsParam.split(',').map(s =>
s.trim()).filter(Boolean) : null;
let anomaliesQuery;
let params = [];
 if (sessionIds && sessionIds.length >
0) {
// Filtre sur les sessions
sélectionnées
const placeholders =
sessionIds.map((_, i) => `$${i +
1}`).join(',');
anomaliesQuery = `
SELECT v.*, s.start_time as
session_start, s.end_time as session_end
FROM validations v
JOIN sessions s ON s.id =
v.session_id
WHERE v.session_id IN
(${placeholders}) AND v.anomalie = true
ORDER BY v.timestamp ASC
`;
params = sessionIds;
} else if (date) {
anomaliesQuery = `
SELECT v.*, s.start_time as
session_start, s.end_time as session_end
FROM validations v
JOIN sessions s ON s.id =
v.session_id
WHERE DATE(s.start_time AT TIME
ZONE 'Europe/Paris') = $1 AND v.anomalie =
true
ORDER BY v.timestamp ASC
`;
params = [date];
} else {
anomaliesQuery = `
 SELECT v.*, s.start_time as
session_start, s.end_time as session_end
FROM validations v
JOIN sessions s ON s.id =
v.session_id
WHERE v.anomalie = true
ORDER BY v.timestamp ASC
`;
}
const anomalies = (await
pool.query(anomaliesQuery, params)).rows;
// Commune par tournée (lue depuis
tournees_sorteurs.xlsx)
let tournees = {};
try { tournees = readTournees(); }
catch (e) { console.error('Erreur lecture
communes:', e); }
// Prestataire assurant la collecte
(fixe pour l'instant)
const PRESTATAIRE = 'Pizzorno';
// ── Construction de la feuille
(format calé sur Exemple_extraction.xlsx)
──
const wb = XLSX.utils.book_new();
const headers = [
'Date', // A
'Heure', // B
'N° circuit', // C
'Commune', // D
 'N° rue', // E
'Adresse', // F
'Type anomalie', // G
'Complément', // H
'Collecte', // I O / N
'Prestataire', // J
'Sortie / Rentrée', // K (bonus,
hors format exemple)
'Nom Prénom sorteur' // L (bonus,
hors format exemple)
];
const rows = anomalies.map(a => {
const ts = a.timestamp ? new
Date(a.timestamp) : null;
const { numero, voie } =
parseAdresse(a.adresse);
const numeroCell =
/^\d+$/.test(numero) ? Number(numero) :
numero;
const collecteCode =
a.collecte_effectuee === true ? 'O' :
a.collecte_effectuee === false ? 'N' : '';
const commune =
(tournees[a.tournee_id] &&
tournees[a.tournee_id].commune) || '';
// N° circuit : on part du N°SDA
correspondant au jour de la semaine de
// l'anomalie (table H:J de
l'onglet). Si aucune correspondance n'est
// trouvée (jour non renseigné dans
le tableau, tournée inconnue, etc.),
// on retombe simplement sur le
numéro de tournée brut.
let circuit = a.tournee_id || '';
const tInfo = tournees[a.tournee_id];
if (tInfo && tInfo.sdaParJour && ts)
{
const jour =
ts.toLocaleDateString('fr-FR', { weekday:
'long', timeZone: 'Europe/Paris'
}).toLowerCase();
if (tInfo.sdaParJour[jour]) circuit
= tInfo.sdaParJour[jour];
}
return [
ts,
// A - Date
ts,
// B - Heure
circuit,
// C - N° circuit (N°SDA du jour, ou
tournée brute à défaut)
commune,
// D - Commune
numeroCell,
// E - N° rue
voie,
// F - Adresse
anomalyTypeLabel(a.anomalie_type),
// G - Type anomalie
(niveau 1)
anomalyComplementLabel(a.anomalie_precision
, a.commentaire), // H - Complément (niveau
2 + commentaire)
collecteCode,
// I - Collecte (O/N)
PRESTATAIRE,
// J - Prestataire
a.type_tournee ||
'', // K -
Sortie / Rentrée (bonus)
a.agent ||
'' //
L - Nom Prénom sorteur (bonus)
];
});
const wsData = [headers, ...rows];
const ws =
XLSX.utils.aoa_to_sheet(wsData, {
cellDates: true });
// Format des colonnes Date / Heure en
valeurs Excel réelles (triables,
filtrables)
for (let r = 0; r < rows.length; r++) {
const rowNum = r + 2;
const dateCell = ws[`A${rowNum}`];
const heureCell = ws[`B${rowNum}`];
 if (dateCell && dateCell.v)
dateCell.z = 'dd/mm/yyyy';
if (heureCell && heureCell.v)
heureCell.z = 'hh:mm';
}
ws['!cols'] = [
{ wch: 12 }, // A Date
{ wch: 8 }, // B Heure
{ wch: 12 }, // C N° circuit
{ wch: 16 }, // D Commune
{ wch: 8 }, // E N° rue
{ wch: 32 }, // F Adresse
{ wch: 32 }, // G Type anomalie
{ wch: 30 }, // H Complément
{ wch: 10 }, // I Collecte
{ wch: 16 }, // J Prestataire
{ wch: 14 }, // K Sortie/Rentrée
{ wch: 24 }, // L Nom Prénom
];
XLSX.utils.book_append_sheet(wb, ws,
'Anomalies');
// Nom du fichier
const fileSuffix = date || (sessionIds
? `selection` : 'complet');
const buf = XLSX.write(wb, { type:
'buffer', bookType: 'xlsx' });
res.setHeader('Content-Disposition',
`attachment;
filename="anomalies_${fileSuffix}.xlsx"`);
res.setHeader('Content-Type',
'application/vnd.openxmlformats-
officedocument.spreadsheetml.sheet');
res.send(buf);
} catch (err) {
console.error('Erreur export excel:',
err);
res.status(500).json({ error:
err.message });
}
});
// ─── API : LOCALISATION
────────────────────────────────────
───────────────────
// La position est capturée passivement à
chaque validation d'adresse (voir
/api/validation).
// Cet endpoint renvoie simplement la
dernière position connue, instantanément,
sans attente.
app.get('/api/location/status/:sessionId',
async (req, res) => {
try {
const r = await pool.query(`
SELECT loc_captured_at, loc_lat,
loc_lng, loc_accuracy
FROM sessions WHERE id = $1
`, [req.params.sessionId]);
if (!r.rows[0]) return
res.status(404).json({ error: 'Non trouvée'
});
res.json(r.rows[0]);
} catch (err) {
res.status(500).json({ error:
err.message });
 }
});
// ─── API : DEBUG
────────────────────────────────────
─────────────────────────
app.get('/api/debug', (req, res) => {
try {
const wb =
XLSX.readFile(path.join(__dirname,
'tournees_sorteurs.xlsx'));
const debug = {};
wb.SheetNames.forEach(name => {
const rows =
XLSX.utils.sheet_to_json(wb.Sheets[name], {
header: 1 });
debug[name] = { nb_lignes:
rows.length, ligne_1: rows[0] || null,
ligne_2: rows[1] || null, ligne_3: rows[2]
|| null };
});
res.json(debug);
} catch (err) { res.status(500).json({
error: err.message }); }
});
// ─── DÉMARRAGE
────────────────────────────────────
───────────────────────────
const PORT = process.env.PORT || 3000;
initDB().then(() => {
app.listen(PORT, () => console.log(`🚀
SorteurPro démarré sur le port ${PORT}`));
}).catch(err => {
 console.error('❌ Erreur init DB:', err);
process.exit(1);
});
