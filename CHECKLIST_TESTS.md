# ✅ Checklist de tests — Praxedo Processor Web v3.0

Suivez cette checklist dans l'ordre pour valider l'installation.

---

## 🔧 Tests d'installation

- [ ] **T01** Python détecté (`python --version` ≥ 3.8)
- [ ] **T02** Dépendances installées (streamlit, pandas, openpyxl, plotly, reportlab, kaleido)
- [ ] **T03** Lancer l'app (double-clic `lancer_web.bat`)
- [ ] **T04** Le navigateur s'ouvre sur `http://localhost:8501`
- [ ] **T05** Bannière bleue « 📊 Praxedo Processor » visible
- [ ] **T06** Barre latérale avec « ⚙️ Paramètres » et « 📈 État de la base »

## 📥 Tests — Onglet Traitement

- [ ] **T10** Glisser un fichier `.xlsm` → apparaît dans la liste
- [ ] **T11** Options cochées par défaut (CSV + base)
- [ ] **T12** Barre de progression fonctionne
- [ ] **T13** Statistiques affichées (POI, Tipo, ARC, CAFF, RAI)
- [ ] **T14** Bandeau bleu DB (nouvelles/maj/identiques)
- [ ] **T15** Téléchargement Excel OK
- [ ] **T16** Téléchargement CSV OK
- [ ] **T17** Vérifier Excel enrichi (colonnes 21-33)
- [ ] **T18** Vérifier CSV des lignes sans POI

## 💾 Tests — Base de données

- [ ] **T20** Fichier `praxedo.db` créé
- [ ] **T21** Sidebar actualisée (interventions, versions, imports)
- [ ] **T22** Re-traitement = N identiques (pas de faux doublons)
- [ ] **T23** Modification = historisation (mise à jour détectée)

## 🔍 Tests — Onglet Recherche

- [ ] **T30** Accès à l'onglet Recherche
- [ ] **T31** Recherche sans filtre (500 résultats max)
- [ ] **T32** Filtre Tipologie = NSTD
- [ ] **T33** Filtre ARC = Oui
- [ ] **T34** Filtre texte sur ND (ex. `0059`)
- [ ] **T35** Recherche globale (ex. `orange`)
- [ ] **T36** Filtres combinés (STD + ARC)
- [ ] **T37** Export Excel
- [ ] **T38** Export CSV

## 📄 Tests — Détail d'intervention

- [ ] **T40** Sélecteur « Voir le détail » présent
- [ ] **T41** Fiche détaillée s'affiche en bas
- [ ] **T42** Champs regroupés en 7 catégories
- [ ] **T43** Onglet Historique accessible
- [ ] **T44** Marqueur ⭐ sur version courante
- [ ] **T45** Champs modifiés listés entre versions

## 📈 Tests — Tableau de bord (NOUVEAU Phase 3)

- [ ] **T70** Accès à l'onglet « Tableau de bord »
- [ ] **T71** Sélecteur de période (6 options)
- [ ] **T72** Sélecteur d'agence (« toutes » + liste)
- [ ] **T73** Bouton « Rafraîchir » fonctionnel
- [ ] **T74** 8 KPI affichés (total, taux, std, nstd, arc, poi, doublons)
- [ ] **T75** Tendance (delta) affichée sur période courte (7j, 30j)
- [ ] **T76** Graphique timeseries lisible
- [ ] **T77** Donut STD/NSTD/Inconnu
- [ ] **T78** Barres empilées par agence
- [ ] **T79** Top 10 techniciens (horizontal)
- [ ] **T80** Top 10 CAFF
- [ ] **T81** Top 10 clients
- [ ] **T82** Top 10 ND récurrents
- [ ] **T83** Bloc alertes affiche 4 compteurs
- [ ] **T84** Expander « Voir le détail » déroule 4 tabs
- [ ] **T85** Filtre agence met à jour tous les graphiques
- [ ] **T86** Filtre période « 7 jours » fonctionne

## 📄 Tests — Export PDF (NOUVEAU Phase 3)

- [ ] **T90** Bouton « Générer le rapport PDF » visible
- [ ] **T91** Clic → spinner « Génération en cours… »
- [ ] **T92** Succès « Rapport généré »
- [ ] **T93** Bouton « Télécharger le PDF » apparaît
- [ ] **T94** PDF téléchargé (entre 8 Ko et 500 Ko)
- [ ] **T95** Page de garde lisible (titre, période, date)
- [ ] **T96** Page KPI avec 8 indicateurs
- [ ] **T97** Graphique timeseries présent (pas « indisponible »)
- [ ] **T98** Graphique donut présent
- [ ] **T99** Graphique barres agence présent
- [ ] **T100** Top techniciens présent
- [ ] **T101** Top CAFF présent
- [ ] **T102** Tableau synthèse par agence
- [ ] **T103** Tableau alertes en dernière page
- [ ] **T104** Pied de page avec numérotation et horodatage

## 📊 Tests — Onglet État de la base

- [ ] **T50** 8 KPI affichés
- [ ] **T51** Journal des imports (50 max)

## 🚪 Tests de fermeture

- [ ] **T60** Arrêt propre (fermer la fenêtre noire)
- [ ] **T61** Relance OK
- [ ] **T62** Persistance de la base

---

## 🎯 Validation finale

**Critère** : au moins **85 tests sur 95** OK.

En cas d'échec, reportez :
- Numéro du test
- Capture d'écran de l'erreur
- Derniers messages de la fenêtre noire
