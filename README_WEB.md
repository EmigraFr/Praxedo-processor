# 📊 Praxedo Processor v3.0 — Version Web (Streamlit)

Application web locale qui fonctionne dans votre navigateur. Compatible avec **toutes les installations Python 3.8+**, sans aucune dépendance à tkinter.

## ✨ Nouveautés v3.0 (Phase 3)

- 📈 **Nouvel onglet Tableau de bord** avec 8 KPI en temps réel
- 📊 **6 graphiques interactifs** (timeseries, donut, barres, tops)
- ⚠️ **Système d'alertes** pour anomalies (POI manquants, emails suspects…)
- 📄 **Export PDF professionnel** (9 pages A4) pour reporting externe
- 🎯 **Filtres temporels** (7j / 30j / 90j / 6 mois / 12 mois / tout)
- 🏢 **Filtre par agence** sur tout le dashboard

## 🚀 Démarrage rapide

### Windows
Double-cliquez sur **`lancer_web.bat`**.

### macOS / Linux
```bash
chmod +x lancer_web.sh
./lancer_web.sh
```

L'app s'ouvre sur **http://localhost:8501** dans votre navigateur.

## 📑 Fonctionnalités complètes

### 📥 Onglet Traitement
- Glisser-déposer de fichiers Excel
- Traitement par lot avec progression temps réel
- Téléchargement direct des fichiers enrichis
- Rapport CSV automatique des lignes sans POI
- Alimentation automatique de la base SQLite

### 🔍 Onglet Recherche
- Barre de recherche globale (texte libre)
- 11 filtres combinables
- Tableau de résultats triable
- Export Excel et CSV
- Détail + historique des versions

### 📈 Onglet Tableau de bord (nouveau)
- **8 KPI** : total, taux STD, taux NSTD, ARC, POI manquants, doublons, tendances
- **Graphique timeseries** : évolution par semaine (Total / STD / NSTD)
- **Donut** : répartition STD / NSTD / Inconnu
- **Barres empilées** : volume par agence × tipologie
- **4 tops** : techniciens, CAFF, clients, ND récurrents
- **Bloc alertes** : POI manquants récents, doublons, emails suspects, CAFF sans téléphone
- **Export PDF** : rapport 9 pages pour diffusion externe

### 📊 Onglet État de la base
- 8 KPI persistants
- Journal des 50 derniers imports

## 🗂️ Arborescence

```
praxedo-app-web-v3/
├── app.py                    ← app Streamlit (4 onglets)
├── run_web.py                ← lanceur avec ouverture navigateur
├── praxedo_core.py           ← moteur d'extraction
├── praxedo_db.py             ← écriture SQLite (historisation)
├── praxedo_queries.py        ← lecture SQLite (recherche)
├── praxedo_dashboard.py      ← KPI + graphiques + export PDF (NOUVEAU)
├── requirements_web.txt      ← 6 dépendances
├── lancer_web.bat            ← lanceur Windows
├── lancer_web.sh             ← lanceur macOS/Linux
├── README_WEB.md             ← ce document
├── GUIDE_INSTALLATION.md     ← guide pas à pas
└── CHECKLIST_TESTS.md        ← tests de validation
```

## 📦 Dépendances

- `openpyxl` — lecture/écriture Excel
- `streamlit` — interface web
- `pandas` — manipulation de données
- `plotly` — graphiques interactifs (NOUVEAU)
- `reportlab` — génération PDF (NOUVEAU)
- `kaleido` — conversion graphiques Plotly en PNG pour le PDF (NOUVEAU)

## ⚙️ Paramètres

Dans la **barre latérale gauche** :
- Dossier de sortie (où sont placés les .xlsx/.csv/.pdf)
- Chemin de la base SQLite

## 📄 Contenu du rapport PDF

Le bouton « Générer le rapport PDF » produit un document A4 de 9 pages :

1. **Page de garde** — titre, période, agence, date, total
2. **Page KPI** — les 8 indicateurs en tableau coloré
3. **Page Évolution temporelle** — graphique ligne
4. **Page Répartition STD/NSTD** — donut
5. **Page Volume par agence** — barres empilées
6. **Page Top techniciens** — classement
7. **Page Top CAFF** — classement
8. **Page Synthèse agence** — tableau détaillé
9. **Page Alertes** — compteurs d'anomalies

Pied de page sur chaque page avec numérotation et horodatage.

## 🔒 Sécurité

- Toutes les données restent sur votre poste
- Le serveur Streamlit écoute uniquement sur `localhost`
- La base `praxedo.db` est un fichier local déplaçable/sauvegardable
- Le PDF est généré localement sans connexion internet

## ❓ Dépannage

**Le port 8501 est déjà utilisé**
→ Le lanceur essaie automatiquement 8502, 8503, 8504, 8505.

**Le navigateur ne s'ouvre pas**
→ Ouvrez-le manuellement : `http://localhost:8501`.

**Le PDF contient "Graphique indisponible"**
→ `kaleido` mal installé. Tapez `python -m pip install --upgrade kaleido`.

**Génération PDF lente (10-30 s)**
→ Normal, la conversion des 5 graphiques en PNG prend du temps.
