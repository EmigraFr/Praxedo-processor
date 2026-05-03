#!/usr/bin/env python3
"""
Praxedo DB — Couche d'accès à la base de données SQLite
========================================================
Gère la persistance cumulative des interventions traitées.

Stratégie de doublons : HISTORISATION
- Chaque Refint unique = 1 ligne dans `interventions` (version courante)
- Toutes les versions successives sont conservées dans `interventions_historique`
- Détection automatique des changements via comparaison champ par champ

Utilisation basique :
    from praxedo_db import PraxedoDB
    db = PraxedoDB("praxedo.db")
    db.import_file_stats(fichier_source_id, import_id, stats)
    db.upsert_intervention(data, import_id)
    db.close()
"""

from __future__ import annotations

import hashlib
import json
import sqlite3
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


APP_VERSION = "3.0"
DB_SCHEMA_VERSION = 1

# Colonnes de données (1-33) — ordre = ordre SQL d'insertion
DATA_COLUMNS = [
    # Colonnes sources (1-20)
    "num_ot", "drapeaux", "statut", "agence", "type_inter", "ui",
    "date_creation", "date_cible_initiale", "a_faire_avant", "date_planifiee",
    "technicien_nom", "technicien_prenom", "equipiers", "client", "nd",
    "act_prod", "description", "code_intervention", "soldee", "validee_annulee",
    # Colonnes enrichies (21-33)
    "poi", "tipologie", "standard", "non_standard", "arc",
    "refint_extrait", "type_production",
    "caff_nom", "caff_tel", "rai_nom", "rai_tel",
    "email_rpe", "telephones",
]

# Mapping col Excel → clé data
EXCEL_COL_MAP = {
    1: "num_ot", 2: "drapeaux", 3: "statut", 4: "agence", 5: "type_inter",
    6: "ui", 7: "date_creation", 8: "date_cible_initiale",
    9: "a_faire_avant", 10: "date_planifiee",
    11: "technicien_nom", 12: "technicien_prenom", 13: "equipiers",
    14: "client", 15: "nd", 16: "act_prod", 17: "description",
    18: "code_intervention", 19: "soldee", 20: "validee_annulee",
    21: "poi", 22: "tipologie", 23: "standard", 24: "non_standard",
    25: "arc", 26: "refint_extrait", 27: "type_production",
    28: "caff_nom", 29: "caff_tel", 30: "rai_nom", 31: "rai_tel",
    32: "email_rpe", 33: "telephones",
}


# =============================================================================
# SCHÉMA SQL
# =============================================================================

SCHEMA_SQL = """
-- Métadonnées de la base
CREATE TABLE IF NOT EXISTS db_meta (
    key   TEXT PRIMARY KEY,
    value TEXT
);

-- ===== Table 1 : fichiers_sources =====
CREATE TABLE IF NOT EXISTS fichiers_sources (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    nom_fichier         TEXT    NOT NULL,
    chemin_complet      TEXT,
    hash_md5            TEXT    UNIQUE,
    taille_octets       INTEGER,
    date_fichier        TEXT,
    premier_import_at   TEXT    NOT NULL,
    dernier_import_at   TEXT    NOT NULL,
    nb_imports          INTEGER DEFAULT 1
);

-- ===== Table 2 : imports =====
CREATE TABLE IF NOT EXISTS imports (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    fichier_source_id   INTEGER REFERENCES fichiers_sources(id),
    date_import         TEXT    NOT NULL,
    version_app         TEXT,
    nb_lignes_lues      INTEGER,
    nb_nouvelles        INTEGER,
    nb_mises_a_jour     INTEGER,
    nb_identiques       INTEGER,
    nb_erreurs          INTEGER DEFAULT 0,
    duree_ms            INTEGER,
    commentaire         TEXT
);

-- ===== Table 3 : interventions (version courante) =====
CREATE TABLE IF NOT EXISTS interventions (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    refint                  TEXT    UNIQUE NOT NULL,
    version_courante_id     INTEGER,
    -- Colonnes sources
    num_ot                  TEXT,
    drapeaux                TEXT,
    statut                  TEXT,
    agence                  TEXT,
    type_inter              TEXT,
    ui                      TEXT,
    date_creation           TEXT,
    date_cible_initiale     TEXT,
    a_faire_avant           TEXT,
    date_planifiee          TEXT,
    technicien_nom          TEXT,
    technicien_prenom       TEXT,
    equipiers               TEXT,
    client                  TEXT,
    nd                      TEXT,
    act_prod                TEXT,
    description             TEXT,
    code_intervention       TEXT,
    soldee                  TEXT,
    validee_annulee         TEXT,
    -- Colonnes enrichies
    poi                     TEXT,
    tipologie               TEXT,
    standard                INTEGER DEFAULT 0,
    non_standard            INTEGER DEFAULT 0,
    arc                     INTEGER DEFAULT 0,
    refint_extrait          TEXT,
    type_production         TEXT,
    caff_nom                TEXT,
    caff_tel                TEXT,
    rai_nom                 TEXT,
    rai_tel                 TEXT,
    email_rpe               TEXT,
    telephones              TEXT,
    -- Métadonnées
    premier_import_id       INTEGER REFERENCES imports(id),
    dernier_import_id       INTEGER REFERENCES imports(id),
    nb_versions             INTEGER DEFAULT 1,
    tipologie_par_repli     INTEGER DEFAULT 0,
    created_at              TEXT    NOT NULL,
    updated_at              TEXT    NOT NULL
);

-- ===== Table 4 : interventions_historique =====
CREATE TABLE IF NOT EXISTS interventions_historique (
    id                       INTEGER PRIMARY KEY AUTOINCREMENT,
    intervention_id          INTEGER NOT NULL REFERENCES interventions(id),
    refint                   TEXT    NOT NULL,
    import_id                INTEGER REFERENCES imports(id),
    numero_version           INTEGER NOT NULL,
    est_version_courante     INTEGER DEFAULT 0,
    -- Snapshot
    num_ot                   TEXT,
    drapeaux                 TEXT,
    statut                   TEXT,
    agence                   TEXT,
    type_inter               TEXT,
    ui                       TEXT,
    date_creation            TEXT,
    date_cible_initiale      TEXT,
    a_faire_avant            TEXT,
    date_planifiee           TEXT,
    technicien_nom           TEXT,
    technicien_prenom        TEXT,
    equipiers                TEXT,
    client                   TEXT,
    nd                       TEXT,
    act_prod                 TEXT,
    description              TEXT,
    code_intervention        TEXT,
    soldee                   TEXT,
    validee_annulee          TEXT,
    poi                      TEXT,
    tipologie                TEXT,
    standard                 INTEGER,
    non_standard             INTEGER,
    arc                      INTEGER,
    refint_extrait           TEXT,
    type_production          TEXT,
    caff_nom                 TEXT,
    caff_tel                 TEXT,
    rai_nom                  TEXT,
    rai_tel                  TEXT,
    email_rpe                TEXT,
    telephones               TEXT,
    -- Diff
    champs_modifies          TEXT,
    version_precedente_id    INTEGER REFERENCES interventions_historique(id),
    created_at               TEXT NOT NULL
);

-- ===== Index =====
CREATE INDEX IF NOT EXISTS idx_inter_nd          ON interventions(nd);
CREATE INDEX IF NOT EXISTS idx_inter_poi         ON interventions(poi);
CREATE INDEX IF NOT EXISTS idx_inter_tipologie   ON interventions(tipologie);
CREATE INDEX IF NOT EXISTS idx_inter_caff        ON interventions(caff_nom);
CREATE INDEX IF NOT EXISTS idx_inter_rai         ON interventions(rai_nom);
CREATE INDEX IF NOT EXISTS idx_inter_email       ON interventions(email_rpe);
CREATE INDEX IF NOT EXISTS idx_inter_agence      ON interventions(agence);
CREATE INDEX IF NOT EXISTS idx_inter_technicien  ON interventions(technicien_nom);

CREATE INDEX IF NOT EXISTS idx_hist_refint       ON interventions_historique(refint);
CREATE INDEX IF NOT EXISTS idx_hist_courante     ON interventions_historique(est_version_courante);
CREATE INDEX IF NOT EXISTS idx_hist_created      ON interventions_historique(created_at);
CREATE INDEX IF NOT EXISTS idx_hist_intervention ON interventions_historique(intervention_id);

-- ===== Vues =====
CREATE VIEW IF NOT EXISTS v_interventions_avec_stats AS
SELECT
    i.*,
    (SELECT COUNT(*) FROM interventions_historique h WHERE h.intervention_id = i.id) AS total_versions,
    (SELECT MIN(created_at) FROM interventions_historique h WHERE h.intervention_id = i.id) AS premiere_apparition,
    (SELECT MAX(created_at) FROM interventions_historique h WHERE h.intervention_id = i.id) AS derniere_modification
FROM interventions i;

CREATE VIEW IF NOT EXISTS v_stats_mensuelles AS
SELECT
    substr(created_at, 1, 7) AS mois,
    COUNT(*)           AS nb_interventions,
    SUM(standard)      AS nb_std,
    SUM(non_standard)  AS nb_nstd,
    SUM(arc)           AS nb_arc
FROM interventions
GROUP BY substr(created_at, 1, 7)
ORDER BY mois;

CREATE VIEW IF NOT EXISTS v_stats_agence AS
SELECT
    COALESCE(agence, 'Non renseignée') AS agence,
    COUNT(*)          AS nb_interventions,
    SUM(standard)     AS nb_std,
    SUM(non_standard) AS nb_nstd,
    SUM(arc)          AS nb_arc
FROM interventions
GROUP BY COALESCE(agence, 'Non renseignée')
ORDER BY nb_interventions DESC;

CREATE VIEW IF NOT EXISTS v_top_techniciens AS
SELECT
    COALESCE(technicien_nom, '') || ' ' || COALESCE(technicien_prenom, '') AS technicien,
    COUNT(*) AS nb_interventions
FROM interventions
WHERE technicien_nom IS NOT NULL AND technicien_nom != ''
GROUP BY technicien
ORDER BY nb_interventions DESC;
"""


# =============================================================================
# CLASSE PRINCIPALE
# =============================================================================

class PraxedoDB:
    """Gestionnaire de base de données SQLite pour Praxedo Processor."""

    def __init__(self, db_path: Union[str, Path]):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(self.db_path))
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self._init_schema()

    def _init_schema(self):
        self.conn.executescript(SCHEMA_SQL)
        # Enregistrer version schéma + version app
        self.conn.execute(
            "INSERT OR REPLACE INTO db_meta(key, value) VALUES (?, ?)",
            ("schema_version", str(DB_SCHEMA_VERSION)),
        )
        self.conn.execute(
            "INSERT OR REPLACE INTO db_meta(key, value) VALUES (?, ?)",
            ("app_version", APP_VERSION),
        )
        if not self._get_meta("created_at"):
            self.conn.execute(
                "INSERT INTO db_meta(key, value) VALUES (?, ?)",
                ("created_at", _now()),
            )
        self.conn.commit()

    def _get_meta(self, key: str) -> Optional[str]:
        row = self.conn.execute(
            "SELECT value FROM db_meta WHERE key = ?", (key,)
        ).fetchone()
        return row["value"] if row else None

    # ------------------------------------------------------------------
    # FICHIERS SOURCES
    # ------------------------------------------------------------------
    def register_file(self, file_path: Path) -> Tuple[int, bool]:
        """Enregistre un fichier. Retourne (id, is_new)."""
        md5 = _md5_of(file_path)
        size = file_path.stat().st_size
        mtime = datetime.fromtimestamp(file_path.stat().st_mtime).isoformat(timespec="seconds")
        now = _now()

        row = self.conn.execute(
            "SELECT id FROM fichiers_sources WHERE hash_md5 = ?", (md5,)
        ).fetchone()

        if row:
            fid = row["id"]
            self.conn.execute(
                "UPDATE fichiers_sources "
                "SET dernier_import_at = ?, nb_imports = nb_imports + 1, "
                "    chemin_complet = ?, nom_fichier = ? "
                "WHERE id = ?",
                (now, str(file_path), file_path.name, fid),
            )
            self.conn.commit()
            return fid, False

        cur = self.conn.execute(
            "INSERT INTO fichiers_sources "
            "(nom_fichier, chemin_complet, hash_md5, taille_octets, date_fichier, "
            " premier_import_at, dernier_import_at, nb_imports) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, 1)",
            (file_path.name, str(file_path), md5, size, mtime, now, now),
        )
        self.conn.commit()
        return cur.lastrowid, True

    # ------------------------------------------------------------------
    # IMPORTS
    # ------------------------------------------------------------------
    def start_import(self, fichier_source_id: int) -> int:
        """Crée une ligne d'import (statistiques remplies à la fin)."""
        cur = self.conn.execute(
            "INSERT INTO imports (fichier_source_id, date_import, version_app) "
            "VALUES (?, ?, ?)",
            (fichier_source_id, _now(), APP_VERSION),
        )
        self.conn.commit()
        return cur.lastrowid

    def finalize_import(self, import_id: int, stats: Dict[str, Any]):
        self.conn.execute(
            "UPDATE imports SET nb_lignes_lues = ?, nb_nouvelles = ?, "
            "nb_mises_a_jour = ?, nb_identiques = ?, nb_erreurs = ?, "
            "duree_ms = ?, commentaire = ? WHERE id = ?",
            (
                stats.get("nb_lignes_lues", 0),
                stats.get("nb_nouvelles", 0),
                stats.get("nb_mises_a_jour", 0),
                stats.get("nb_identiques", 0),
                stats.get("nb_erreurs", 0),
                stats.get("duree_ms", 0),
                stats.get("commentaire", ""),
                import_id,
            ),
        )
        self.conn.commit()

    # ------------------------------------------------------------------
    # INTERVENTIONS — upsert avec historisation
    # ------------------------------------------------------------------
    def upsert_intervention(self, data: dict, import_id: int) -> str:
        """
        Insère ou met à jour une intervention. Retourne :
        - 'new'       si nouvelle intervention
        - 'updated'   si nouvelle version créée
        - 'identical' si aucun changement
        - 'skipped'   si refint manquant
        """
        refint = (data.get("refint_extrait") or "").strip()
        if not refint:
            return "skipped"

        # Normalisation des valeurs avant insertion/comparaison
        for col in DATA_COLUMNS:
            v = data.get(col)
            if v is not None and hasattr(v, "isoformat"):
                data[col] = (v.isoformat(timespec="seconds")
                             if hasattr(v, "hour") else v.isoformat())

        now = _now()
        existing = self.conn.execute(
            "SELECT * FROM interventions WHERE refint = ?", (refint,)
        ).fetchone()

        if existing is None:
            # ================ NOUVELLE INTERVENTION ================
            cols = ["refint"] + DATA_COLUMNS + [
                "premier_import_id", "dernier_import_id",
                "nb_versions", "tipologie_par_repli",
                "created_at", "updated_at",
            ]
            values = [refint] + [data.get(c) for c in DATA_COLUMNS] + [
                import_id, import_id, 1,
                1 if data.get("tipologie_par_repli") else 0,
                now, now,
            ]
            placeholders = ",".join(["?"] * len(values))
            cur = self.conn.execute(
                f"INSERT INTO interventions ({','.join(cols)}) VALUES ({placeholders})",
                values,
            )
            inter_id = cur.lastrowid

            # Créer la 1re version dans l'historique
            hist_id = self._insert_historique(
                inter_id, refint, import_id, 1, True, data,
                champs_modifies=None, version_precedente_id=None,
            )
            self.conn.execute(
                "UPDATE interventions SET version_courante_id = ? WHERE id = ?",
                (hist_id, inter_id),
            )
            self.conn.commit()
            return "new"

        # ================ INTERVENTION EXISTANTE ================
        diffs = _compute_diff(existing, data)
        if not diffs:
            self.conn.commit()
            return "identical"

        # Nouvelle version historisée
        inter_id = existing["id"]
        nouvelle_version = (existing["nb_versions"] or 1) + 1
        prev_hist_id = existing["version_courante_id"]

        # Marquer l'ancienne version comme non-courante
        self.conn.execute(
            "UPDATE interventions_historique SET est_version_courante = 0 "
            "WHERE id = ?",
            (prev_hist_id,),
        )
        hist_id = self._insert_historique(
            inter_id, refint, import_id, nouvelle_version, True, data,
            champs_modifies=diffs, version_precedente_id=prev_hist_id,
        )

        # Mettre à jour la vue courante
        set_clause = ",".join([f"{c} = ?" for c in DATA_COLUMNS])
        update_values = [data.get(c) for c in DATA_COLUMNS] + [
            import_id, nouvelle_version,
            1 if data.get("tipologie_par_repli") else 0,
            hist_id, now, inter_id,
        ]
        self.conn.execute(
            f"UPDATE interventions SET {set_clause}, "
            f"dernier_import_id = ?, nb_versions = ?, tipologie_par_repli = ?, "
            f"version_courante_id = ?, updated_at = ? WHERE id = ?",
            update_values,
        )
        self.conn.commit()
        return "updated"

    def _insert_historique(self, intervention_id: int, refint: str,
                           import_id: int, numero_version: int,
                           courante: bool, data: dict,
                           champs_modifies: Optional[List],
                           version_precedente_id: Optional[int]) -> int:
        cols = [
            "intervention_id", "refint", "import_id",
            "numero_version", "est_version_courante",
        ] + DATA_COLUMNS + [
            "champs_modifies", "version_precedente_id", "created_at",
        ]
        values = [
            intervention_id, refint, import_id,
            numero_version, 1 if courante else 0,
        ] + [data.get(c) for c in DATA_COLUMNS] + [
            json.dumps(champs_modifies) if champs_modifies else None,
            version_precedente_id,
            _now(),
        ]
        placeholders = ",".join(["?"] * len(values))
        cur = self.conn.execute(
            f"INSERT INTO interventions_historique "
            f"({','.join(cols)}) VALUES ({placeholders})",
            values,
        )
        return cur.lastrowid

    # ------------------------------------------------------------------
    # STATISTIQUES GLOBALES
    # ------------------------------------------------------------------
    def summary(self) -> dict:
        c = self.conn
        r = lambda q, *a: c.execute(q, a).fetchone()
        stats = {
            "nb_interventions":     r("SELECT COUNT(*) AS n FROM interventions")["n"],
            "nb_historique":        r("SELECT COUNT(*) AS n FROM interventions_historique")["n"],
            "nb_fichiers":          r("SELECT COUNT(*) AS n FROM fichiers_sources")["n"],
            "nb_imports":           r("SELECT COUNT(*) AS n FROM imports")["n"],
            "dernier_import":       r("SELECT MAX(date_import) AS d FROM imports")["d"],
            "taille_octets":        self.db_path.stat().st_size if self.db_path.exists() else 0,
            "nb_std":               r("SELECT SUM(standard) AS n FROM interventions")["n"] or 0,
            "nb_nstd":              r("SELECT SUM(non_standard) AS n FROM interventions")["n"] or 0,
            "nb_arc":               r("SELECT SUM(arc) AS n FROM interventions")["n"] or 0,
        }
        return stats

    # ------------------------------------------------------------------
    # SAUVEGARDE
    # ------------------------------------------------------------------
    def backup(self, backup_dir: Path, keep: int = 10) -> Path:
        """Crée une copie datée de la base. Rotation : garde les `keep` plus récentes."""
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = backup_dir / f"praxedo_backup_{stamp}.db"
        self.conn.commit()
        shutil.copy2(self.db_path, target)

        backups = sorted(backup_dir.glob("praxedo_backup_*.db"), reverse=True)
        for old in backups[keep:]:
            try:
                old.unlink()
            except OSError:
                pass
        return target

    def close(self):
        self.conn.commit()
        self.conn.close()

    def __enter__(self): return self
    def __exit__(self, *a): self.close()


# =============================================================================
# UTILITAIRES
# =============================================================================

def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _md5_of(path: Path, chunk: int = 65536) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def _normalize(v) -> str:
    """Convertit n'importe quelle valeur en chaîne normalisée pour comparaison.

    Gère : None, str, int, float, datetime, date, bool.
    """
    if v is None or v == "":
        return ""
    # datetime / date : utiliser isoformat pour cohérence
    if hasattr(v, "isoformat"):
        return v.isoformat(timespec="seconds") if hasattr(v, "hour") else v.isoformat()
    # bool : True/False -> 1/0
    if isinstance(v, bool):
        return "1" if v else "0"
    # Nombres : normaliser float(1.0) == int(1)
    if isinstance(v, (int, float)):
        try:
            f = float(v)
            if f.is_integer():
                return str(int(f))
            return str(f)
        except (TypeError, ValueError):
            pass
    return str(v).strip()


def _compute_diff(existing: sqlite3.Row, data: dict) -> List[str]:
    """Renvoie la liste des champs DATA_COLUMNS dont la valeur a changé."""
    diffs = []
    for col in DATA_COLUMNS:
        old = existing[col] if col in existing.keys() else None
        new = data.get(col)
        if _normalize(old) != _normalize(new):
            diffs.append(col)
    return diffs


# =============================================================================
# HELPER : conversion ligne Excel -> dict
# =============================================================================

def row_to_data(ws_row_values: Dict[int, Any], tipologie_par_repli: bool = False) -> dict:
    """
    Convertit un dict {numero_colonne: valeur} (col 1-33) en dict data
    utilisable par upsert_intervention.
    """
    d = {EXCEL_COL_MAP[c]: v for c, v in ws_row_values.items() if c in EXCEL_COL_MAP}
    d["tipologie_par_repli"] = 1 if tipologie_par_repli else 0
    return d
