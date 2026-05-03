# 🚀 Guide d'installation Windows — Praxedo Processor Web

Ce guide vous accompagne pas à pas pour lancer l'application sur votre poste Windows.

---

## 📋 Prérequis

- **Windows 10 ou 11**
- **Python 3.8 à 3.12** installé (peu importe la provenance, y compris Microsoft Store)
- Connexion internet uniquement pour le **premier lancement** (installation des dépendances)

---

## 🎯 Installation en 4 étapes

### Étape 1 — Vérifier Python

Ouvrez une invite de commandes :
- Appuyez sur <kbd>Windows</kbd> + <kbd>R</kbd>
- Tapez `cmd` et appuyez sur <kbd>Entrée</kbd>

Tapez :

```
python --version
```

Vous devez voir quelque chose comme `Python 3.11.5`. Si vous obtenez un message d'erreur, installez Python depuis **https://www.python.org/downloads/** (cochez bien « Add Python to PATH »).

### Étape 2 — Récupérer le dossier de l'app

Extrayez le fichier `praxedo-app-web.zip` à l'endroit de votre choix, par exemple :

```
C:\Users\VotreNom\Documents\praxedo-app-web\
```

### Étape 3 — Premier lancement

Dans le dossier `praxedo-app-web`, **double-cliquez sur `lancer_web.bat`**.

Une fenêtre noire s'ouvre et affiche :

```
========================================
  Praxedo Processor - Version Web
========================================

[1/2] Installation des dependances...
[2/2] Demarrage de l'application...

Le navigateur va s'ouvrir automatiquement.
```

Après 10 à 30 secondes (première fois uniquement), votre navigateur s'ouvre automatiquement sur **http://localhost:8501** avec l'application.

### Étape 4 — Utilisation

- **Laissez la fenêtre noire ouverte** tant que vous utilisez l'app
- Pour arrêter : fermez la fenêtre noire (ou <kbd>Ctrl</kbd>+<kbd>C</kbd>)
- Pour rouvrir plus tard : re-double-cliquez sur `lancer_web.bat`

---

## 🖼️ Utilisation pas à pas

### Traiter un fichier

1. Restez sur l'onglet **📥 Traitement**
2. **Glissez-déposez** votre fichier `.xlsm` dans la zone de dépôt
3. Cliquez sur **🚀 Lancer le traitement**
4. Téléchargez le `_traite.xlsx` via le bouton vert en bas

### Rechercher dans la base

1. Cliquez sur l'onglet **🔍 Recherche**
2. Saisissez un ou plusieurs filtres (par exemple : ND = `0059GEB2`)
3. Cliquez sur **🔍 Rechercher**
4. Double-cliquez sur une ligne pour voir le détail et l'historique

### Exporter un résultat filtré

1. Après une recherche, cliquez sur **📤 Exporter vers Excel**
2. Le fichier se télécharge immédiatement

---

## ❓ Problèmes fréquents

### Le navigateur ne s'ouvre pas automatiquement
Ouvrez-le manuellement et tapez dans la barre d'adresse : **http://localhost:8501**

### « Python n'est pas reconnu »
Python n'est pas dans le PATH. Réinstallez-le depuis python.org en cochant **« Add Python to PATH »**.

### « ModuleNotFoundError: No module named 'streamlit' »
L'installation auto a échoué. Tapez dans l'invite :
```
python -m pip install streamlit pandas openpyxl
```

### Le port 8501 est déjà utilisé
Le lanceur bascule automatiquement sur 8502, 8503, etc. Regardez l'adresse indiquée dans la fenêtre noire.
