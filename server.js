const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static('.'));

// Lit le fichier Excel et retourne les tournées
app.get('/api/tournees', (req, res) => {
  const filePath = path.join(__dirname, 'tournees_sorteurs.xlsx');
  const wb = XLSX.readFile(filePath);
  const tournees = {};

  wb.SheetNames.forEach(name => {
    if (name !== '📊 Récapitulatif') {
      const ws = wb.Sheets[name];
      const rows = XLSX.utils.sheet_to_json(ws, { range: 3 });
      tournees[name] = {
        label: name,
        addresses: rows
          .filter(r => r['Adresse complète'])
          .map(r => ({
            name: r['Adresse complète'],
            sub: (r['Complément / Nom résidence'] || '')
               + (r['Nb bacs'] ? ' — ' + r['Nb bacs'] + ' bacs' : '')
          }))
      };
    }
  });
  res.json(tournees);
});

// Reçoit les validations des salariés
const validations = [];
app.post('/api/validation', (req, res) => {
  validations.push({ ...req.body, receivedAt: new Date() });
  console.log('Validation reçue:', req.body);
  res.json({ ok: true });
});

app.get('/api/validations', (req, res) => {
  res.json(validations);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log('SorteurPro démarré sur le port ' + PORT);
});
app.get('/api/debug', (req, res) => {
  const filePath = path.join(__dirname, 'tournees_sorteurs.xlsx');
  const wb = XLSX.readFile(filePath);
  const debug = {};
  wb.SheetNames.forEach(name => {
    const ws = wb.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { range: 3 });
    debug[name] = {
      colonnes: rows.length > 0 ? Object.keys(rows[0]) : [],
      premiere_ligne: rows[0] || null,
      nb_lignes: rows.length
    };
  });
  res.json(debug);
});
