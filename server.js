const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static('.'));

const ONGLETS_A_IGNORER = ['📊 Récapitulatif'];

// ─── TOURNÉES : lit le fichier Excel par position de colonne ───
app.get('/api/tournees', (req, res) => {
  try {
    const filePath = path.join(__dirname, 'tournees_sorteurs.xlsx');
    const wb = XLSX.readFile(filePath);
    const tournees = {};

    wb.SheetNames.forEach(name => {
      if (ONGLETS_A_IGNORER.includes(name)) return;

      const ws = wb.Sheets[name];
      // Lit toutes les lignes sans en-tête — chaque ligne = tableau de valeurs
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

      const addresses = [];
      rows.forEach(row => {
        // Colonne B (index 1) = adresse complète
        // Colonne C (index 2) = complément / nom résidence
        // Colonne D (index 3) = nb bacs
        const adresse    = row[1];
        const complement = row[2] || '';
        const nbBacs     = row[3] || '';

        // Ignore les lignes vides
        if (!adresse || typeof adresse !== 'string' || adresse.trim() === '') return;
        // Ignore la ligne d'en-tête si elle contient le mot "adresse"
        if (adresse.toLowerCase().includes('adresse')) return;

        addresses.push({
          name: adresse.trim(),
          sub: [complement.toString().trim(), nbBacs ? nbBacs + ' bacs' : '']
            .filter(Boolean).join(' — ')
        });
      });

      tournees[name] = {
        label: name,
        addresses
      };
    });

    res.json(tournees);
  } catch (err) {
    console.error('Erreur lecture Excel:', err);
    res.status(500).json({ error: 'Impossible de lire le fichier Excel.' });
  }
});

// ─── VALIDATIONS : reçoit et stocke les validations des salariés ───
const validations = [];

app.post('/api/validation', (req, res) => {
  const entry = { ...req.body, receivedAt: new Date().toISOString() };
  validations.push(entry);
  console.log('Validation reçue:', JSON.stringify(entry));
  res.json({ ok: true });
});

app.get('/api/validations', (req, res) => {
  res.json(validations);
});

// ─── DEBUG : affiche la structure brute du fichier Excel ───
app.get('/api/debug', (req, res) => {
  try {
    const filePath = path.join(__dirname, 'tournees_sorteurs.xlsx');
    const wb = XLSX.readFile(filePath);
    const debug = {};

    wb.SheetNames.forEach(name => {
      const ws = wb.Sheets[name];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      debug[name] = {
        nb_lignes: rows.length,
        ligne_1: rows[0] || null,
        ligne_2: rows[1] || null,
        ligne_3: rows[2] || null,
        ligne_4: rows[3] || null,
      };
    });

    res.json(debug);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── DÉMARRAGE ───
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log('SorteurPro démarré sur le port ' + PORT);
});
