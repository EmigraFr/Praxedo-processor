#!/usr/bin/env python3
"""
Praxedo Search — Moteur de recherche avancée
=============================================
Couche de requêtes dédiée à la page Recherche avancée :
  - advanced_search(filters) : recherche multi-critères avec opérateurs
  - get_history_diff(refint)  : historique avec comparaison entre versions
  - detect_anomalies()        : tableau récapitulatif des anomalies
  - export_to_excel/csv()     : exports enrichis
"""

from __future__ import annotations
from typing import Any, Dict, List, Optional, Tuple

import io
import json
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path


# =============================================================================
# DÉFINITION DES FILTRES AVANCÉS
# =============================================================================

# Opérateurs disponibles par type de champ
TEXT_OPERATORS = {
    "contient":        lambda col, v: (f"LOWER({col}) LIKE ?", f"%{v.lower()}%"),
    "commence par":    lambda col, v: (f"LOWER({col}) LIKE ?", f"{v.lower()}%"),
    "finit par":       lambda col, v: (f"LOWER({col}) LIKE ?", f"%{v.lower()}"),
    "égal à":          lambda col, v: (f"LOWER({col}) = ?", v.lower()),
    "différent de":    lambda col, v: (f"LOWER({col}) != ?", v.lower()),
    "vide":            lambda col, v: (f"({col} IS NULL OR {col} = '')", None),
    "non vide":        lambda col, v: (f"({col} IS NOT NULL AND {col} != '')", None),
}

BOOL_OPERATORS = {
    "oui":  lambda col, v: (f"{col} = 1", None),
    "non":  lambda col, v: (f"{col} = 0 OR {col} IS NULL", None),
    "tous": lambda col, v: (None, None),
}

DATE_OPERATORS = {
    "après":       lambda col, v: (f"{col} >= ?", v),
    "avant":       lambda col, v: (f"{col} <= ?", v),
    "exactement":  lambda col, v: (f"DATE({col}) = DATE(?)", v),
    "7 derniers jours":  lambda col, v: (f"{col} >= ?",
                                         (datetime.now() - timedelta(days=7)).isoformat()),
    "30 derniers jours": lambda col, v: (f"{col} >= ?",
                                         (datetime.now() - timedelta(days=30)).isoformat()),
    "90 derniers jours": lambda col, v: (f"{col} >= ?",
                                         (datetime.now() - timedelta(days=90)).isoformat()),
}

# Champs disponibles pour les filtres avancés
ADVANCED_FIELDS = [
    # (clé, libellé, type, colonne SQL)
    ("refint",         "Refint",              "text", "refint"),
    ("nd",             "ND",                  "text", "nd"),
    ("poi",            "POI",                 "text", "poi"),
    ("num_ot",         "N° OT",               "text", "num_ot"),
    ("agence",         "Agence",              "text", "agence"),
    ("client",         "Client",              "text", "client"),
    ("technicien_nom", "Technicien",          "text", "technicien_nom"),
    ("tipologie",      "Tipologie",           "text", "tipologie"),
    ("caff_nom",       "CAFF",                "text", "caff_nom"),
    ("caff_tel",       "CAFF Tel",            "text", "caff_tel"),
    ("rai_nom",        "RAI",                 "text", "rai_nom"),
    ("rai_tel",        "RAI Tel",             "text", "rai_tel"),
    ("email_rpe",      "Email RPE",           "text", "email_rpe"),
    ("type_production","Type Production",     "text", "type_production"),
    ("description",    "Description",         "text", "description"),
    ("standard",       "Standard",            "bool", "standard"),
    ("non_standard",   "Non Standard",        "bool", "non_standard"),
    ("arc",            "ARC",                 "bool", "arc"),
    ("date_creation",  "Date création",       "date", "date_creation"),
    ("date_planifiee", "Date planifiée",      "date", "date_planifiee"),
    ("updated_at",     "Dernière mise à jour","date", "updated_at"),
    ("created_at",     "Créée en base",       "date", "created_at"),
]

FIELD_BY_KEY = {k: (lbl, typ, col) for k, lbl, typ, col in ADVANCED_FIELDS}


# =============================================================================
# RECHERCHE AVANCÉE AVEC OPÉRATEURS
# =============================================================================

def advanced_search(conn: sqlite3.Connection,
                    criteria: List[Dict[str, Any]],
                    combinator: str = "AND",
                    order_by: str = "updated_at",
                    order_dir: str = "DESC",
                    limit: int = 1000,
                    offset: int = 0) -> Tuple[List[sqlite3.Row], int]:
    """
    Recherche avancée avec liste de critères et combinateur.

    criteria = [
        {"field": "poi",       "operator": "commence par", "value": "SBR"},
        {"field": "tipologie", "operator": "égal à",       "value": "NSTD"},
        {"field": "arc",       "operator": "oui"},
    ]
    combinator = "AND" ou "OR"
    """
    where_parts = []
    params = []

    for c in criteria:
        field_key = c.get("field")
        operator = c.get("operator", "").strip()
        value = c.get("value", "")

        if field_key not in FIELD_BY_KEY:
            continue

        _, ftype, sql_col = FIELD_BY_KEY[field_key]

        if ftype == "text":
            if operator not in TEXT_OPERATORS:
                continue
            if operator in ("vide", "non vide"):
                clause, param = TEXT_OPERATORS[operator](sql_col, None)
                where_parts.append(clause)
            elif not str(value).strip():
                continue
            else:
                clause, param = TEXT_OPERATORS[operator](sql_col, str(value).strip())
                where_parts.append(clause)
                if param is not None:
                    params.append(param)

        elif ftype == "bool":
            if operator not in BOOL_OPERATORS:
                continue
            clause, _ = BOOL_OPERATORS[operator](sql_col, None)
            if clause:
                where_parts.append(clause)

        elif ftype == "date":
            if operator not in DATE_OPERATORS:
                continue
            if operator in ("après", "avant", "exactement") and not str(value).strip():
                continue
            clause, param = DATE_OPERATORS[operator](sql_col, str(value).strip())
            where_parts.append(clause)
            if param is not None:
                params.append(param)

    # Construire la clause WHERE
    if where_parts:
        joiner = " AND " if combinator == "AND" else " OR "
        where_sql = " WHERE " + joiner.join(f"({p})" for p in where_parts)
    else:
        where_sql = ""

    # Sécurisation du tri
    allowed_order = {c for c, _, _, _ in [(k, l, t, col)
                                           for k, l, t, col in ADVANCED_FIELDS]}
    allowed_order.update({"updated_at", "created_at", "nb_versions"})
    order_col = order_by if order_by in allowed_order else "updated_at"
    order_sql = f"ORDER BY {order_col} {'ASC' if order_dir == 'ASC' else 'DESC'}"

    # Compteur
    count_q = f"SELECT COUNT(*) FROM interventions {where_sql}"
    total = conn.execute(count_q, params).fetchone()[0]

    # Résultats
    conn.row_factory = sqlite3.Row
    data_q = (f"SELECT * FROM interventions {where_sql} "
              f"{order_sql} LIMIT ? OFFSET ?")
    rows = conn.execute(data_q, params + [limit, offset]).fetchall()
    return list(rows), total


# =============================================================================
# HISTORIQUE DÉTAILLÉ AVEC COMPARAISON
# =============================================================================

def get_history_detailed(conn: sqlite3.Connection, refint: str) -> List[Dict]:
    """Retourne l'historique enrichi avec les valeurs avant/après pour
    chaque champ modifié."""
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        "SELECT h.*, i.date_import AS import_date, "
        "       f.nom_fichier AS source_file "
        "FROM interventions_historique h "
        "LEFT JOIN imports i ON i.id = h.import_id "
        "LEFT JOIN fichiers_sources f ON f.id = i.fichier_source_id "
        "WHERE h.refint = ? "
        "ORDER BY h.numero_version ASC",
        (refint,),
    ).fetchall()

    result = []
    for i, r in enumerate(rows):
        entry = dict(r)
        modif = r["champs_modifies"]
        diffs = []
        if modif:
            try:
                champs = json.loads(modif) if isinstance(modif, str) else modif
            except (json.JSONDecodeError, TypeError):
                champs = []
            if i > 0 and champs:
                prev = rows[i - 1]
                for field in champs:
                    old_val = prev[field] if field in prev.keys() else None
                    new_val = r[field] if field in r.keys() else None
                    diffs.append({
                        "champ": field,
                        "avant": _short(old_val),
                        "apres": _short(new_val),
                    })
        entry["diffs_detailles"] = diffs
        result.append(entry)
    return result


def _short(v, max_len=80):
    if v is None:
        return "(vide)"
    s = str(v)
    if len(s) > max_len:
        return s[:max_len] + "…"
    return s


# =============================================================================
# TABLEAU RÉCAPITULATIF DES ANOMALIES
# =============================================================================

def detect_anomalies(conn: sqlite3.Connection) -> Dict[str, List[Dict]]:
    """
    Détecte toutes les anomalies et retourne un dict structuré.
    Chaque catégorie a : title, severity, count, items
    """
    conn.row_factory = sqlite3.Row
    anomalies = {}

    # 1. POI manquants
    rows = conn.execute(
        "SELECT id, refint, nd, client, updated_at "
        "FROM interventions "
        "WHERE (poi IS NULL OR poi = '') "
        "ORDER BY updated_at DESC"
    ).fetchall()
    anomalies["poi_manquants"] = {
        "title": "POI manquants",
        "severity": "warning",
        "description": "Interventions sans POI extrait — contrôle manuel nécessaire",
        "items": [dict(r) for r in rows],
    }

    # 2. Tipologie absente
    rows = conn.execute(
        "SELECT id, refint, nd, client, updated_at "
        "FROM interventions "
        "WHERE (tipologie IS NULL OR tipologie = '') "
        "ORDER BY updated_at DESC"
    ).fetchall()
    anomalies["tipologie_absente"] = {
        "title": "Tipologie absente",
        "severity": "info",
        "description": "Aucun STD/NSTD détecté (même après repli par ND)",
        "items": [dict(r) for r in rows],
    }

    # 3. Doublons (interventions historisées)
    rows = conn.execute(
        "SELECT id, refint, nd, nb_versions, updated_at "
        "FROM interventions "
        "WHERE nb_versions > 1 "
        "ORDER BY nb_versions DESC, updated_at DESC"
    ).fetchall()
    anomalies["doublons"] = {
        "title": "Interventions versionnées",
        "severity": "info",
        "description": "Interventions qui ont évolué dans le temps (plusieurs versions)",
        "items": [dict(r) for r in rows],
    }

    # 4. ND récurrents (même ND plusieurs interventions distinctes)
    rows = conn.execute(
        "SELECT nd, COUNT(*) AS n "
        "FROM interventions "
        "WHERE nd IS NOT NULL AND nd != '' "
        "GROUP BY nd HAVING COUNT(*) > 1 "
        "ORDER BY n DESC"
    ).fetchall()
    anomalies["nd_recurrents"] = {
        "title": "ND récurrents",
        "severity": "warning",
        "description": "Un même ND apparaît dans plusieurs interventions distinctes",
        "items": [dict(r) for r in rows],
    }

    # 5. Email RPE non-Orange
    rows = conn.execute(
        "SELECT id, refint, nd, email_rpe, updated_at "
        "FROM interventions "
        "WHERE email_rpe IS NOT NULL AND email_rpe != '' "
        "AND LOWER(email_rpe) NOT LIKE '%orange%' "
        "ORDER BY updated_at DESC"
    ).fetchall()
    anomalies["emails_suspects"] = {
        "title": "Emails RPE non-Orange",
        "severity": "warning",
        "description": "Email RPE au format inattendu (pas de domaine orange)",
        "items": [dict(r) for r in rows],
    }

    # 6. CAFF sans téléphone
    rows = conn.execute(
        "SELECT id, refint, nd, caff_nom, updated_at "
        "FROM interventions "
        "WHERE caff_nom IS NOT NULL AND caff_nom != '' "
        "AND (caff_tel IS NULL OR caff_tel = '') "
        "ORDER BY updated_at DESC"
    ).fetchall()
    anomalies["caff_sans_tel"] = {
        "title": "CAFF sans téléphone",
        "severity": "info",
        "description": "Nom CAFF présent mais numéro absent",
        "items": [dict(r) for r in rows],
    }

    # 7. RAI sans téléphone
    rows = conn.execute(
        "SELECT id, refint, nd, rai_nom, updated_at "
        "FROM interventions "
        "WHERE rai_nom IS NOT NULL AND rai_nom != '' "
        "AND (rai_tel IS NULL OR rai_tel = '') "
        "ORDER BY updated_at DESC"
    ).fetchall()
    anomalies["rai_sans_tel"] = {
        "title": "RAI sans téléphone",
        "severity": "info",
        "description": "Nom RAI présent mais numéro absent",
        "items": [dict(r) for r in rows],
    }

    # 8. Incohérence STD/NSTD
    rows = conn.execute(
        "SELECT id, refint, nd, standard, non_standard, tipologie, updated_at "
        "FROM interventions "
        "WHERE (standard = 1 AND non_standard = 1) "
        "OR (standard = 1 AND tipologie = 'NSTD') "
        "OR (non_standard = 1 AND tipologie = 'STD')"
    ).fetchall()
    anomalies["incoherences_tipo"] = {
        "title": "Incohérences Tipologie",
        "severity": "error",
        "description": "Valeurs Standard/Non Standard contradictoires",
        "items": [dict(r) for r in rows],
    }

    # 9. Description très courte (< 50 caractères)
    rows = conn.execute(
        "SELECT id, refint, nd, LENGTH(description) AS lg, updated_at "
        "FROM interventions "
        "WHERE description IS NOT NULL AND LENGTH(description) < 50 "
        "ORDER BY lg ASC"
    ).fetchall()
    anomalies["desc_courtes"] = {
        "title": "Descriptions trop courtes",
        "severity": "info",
        "description": "Description de moins de 50 caractères (potentiellement incomplète)",
        "items": [dict(r) for r in rows],
    }

    return anomalies


def anomalies_summary(anomalies: Dict[str, Dict]) -> List[Dict]:
    """Transforme le dict d'anomalies en tableau récapitulatif plat."""
    severity_icons = {"error": "🔴", "warning": "🟠", "info": "🔵"}
    summary = []
    for key, data in anomalies.items():
        summary.append({
            "Gravité":      severity_icons.get(data["severity"], "⚪"),
            "Catégorie":    data["title"],
            "Description":  data["description"],
            "Nombre":       len(data["items"]),
            "_key":         key,
        })
    summary.sort(key=lambda x: -x["Nombre"])
    return summary


# =============================================================================
# EXPORTS ENRICHIS
# =============================================================================

def export_search_to_excel(rows: List[sqlite3.Row],
                           output_path: Path,
                           criteria_used: Optional[List[Dict]] = None) -> Path:
    """Exporte les résultats vers Excel avec feuille « Critères » séparée."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()

    # Feuille 1 : Résultats
    ws = wb.active
    ws.title = "Resultats"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="305496")
    center = Alignment(horizontal="center", vertical="center")

    if not rows:
        ws["A1"] = "Aucun résultat"
    else:
        keys = list(rows[0].keys())
        for c, k in enumerate(keys, start=1):
            cell = ws.cell(row=1, column=c, value=k)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center

        for r_idx, row in enumerate(rows, start=2):
            for c_idx, k in enumerate(keys, start=1):
                ws.cell(row=r_idx, column=c_idx, value=row[k])

        for c, k in enumerate(keys, start=1):
            letter = ws.cell(row=1, column=c).column_letter
            ws.column_dimensions[letter].width = max(10, min(40, len(k) + 4))

        ws.freeze_panes = "A2"

    # Feuille 2 : Critères de recherche
    if criteria_used:
        ws2 = wb.create_sheet("Criteres")
        ws2["A1"] = "Critère"
        ws2["B1"] = "Opérateur"
        ws2["C1"] = "Valeur"
        for c in ws2["A1:C1"][0]:
            c.font = header_font
            c.fill = header_fill
            c.alignment = center
        for i, crit in enumerate(criteria_used, start=2):
            ws2.cell(row=i, column=1, value=crit.get("field", ""))
            ws2.cell(row=i, column=2, value=crit.get("operator", ""))
            ws2.cell(row=i, column=3, value=crit.get("value", ""))
        ws2.column_dimensions["A"].width = 20
        ws2.column_dimensions["B"].width = 18
        ws2.column_dimensions["C"].width = 30

        ws2.cell(row=len(criteria_used) + 3, column=1,
                 value="Date export :").font = Font(bold=True)
        ws2.cell(row=len(criteria_used) + 3, column=2,
                 value=datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        ws2.cell(row=len(criteria_used) + 4, column=1,
                 value="Nb résultats :").font = Font(bold=True)
        ws2.cell(row=len(criteria_used) + 4, column=2, value=len(rows))

    wb.save(output_path)
    return output_path


def export_rows_to_csv_bytes(rows: List[sqlite3.Row]) -> bytes:
    """Exporte les résultats en CSV UTF-8 BOM avec séparateur ';'."""
    if not rows:
        return "aucun resultat\n".encode("utf-8-sig")

    import csv
    buf = io.StringIO()
    keys = list(rows[0].keys())
    writer = csv.DictWriter(buf, fieldnames=keys, delimiter=";")
    writer.writeheader()
    for r in rows:
        writer.writerow({k: r[k] if r[k] is not None else "" for k in keys})
    return buf.getvalue().encode("utf-8-sig")


def export_history_to_excel(history: List[Dict],
                            refint: str,
                            output_path: Path) -> Path:
    """Exporte l'historique détaillé d'un Refint vers Excel."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()

    # Feuille 1 : Versions
    ws = wb.active
    ws.title = "Versions"
    ws["A1"] = f"Historique de l'intervention : {refint}"
    ws["A1"].font = Font(bold=True, size=14, color="305496")
    ws.merge_cells("A1:E1")

    headers = ["Version", "Date", "Fichier source",
               "Champs modifiés", "Statut"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="305496")

    for i, h in enumerate(history, start=4):
        modif = h.get("champs_modifies") or "(version initiale)"
        ws.cell(row=i, column=1,
                value=f"v{h['numero_version']}" +
                      (" ⭐" if h.get("est_version_courante") else ""))
        ws.cell(row=i, column=2, value=h.get("created_at") or "")
        ws.cell(row=i, column=3, value=h.get("source_file") or "")
        ws.cell(row=i, column=4, value=modif)
        ws.cell(row=i, column=5,
                value="Courante" if h.get("est_version_courante") else "")

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 35
    ws.column_dimensions["D"].width = 50
    ws.column_dimensions["E"].width = 14

    # Feuille 2 : Diffs détaillés
    ws2 = wb.create_sheet("Changements")
    headers = ["Version", "Champ", "Avant", "Après"]
    for c, h in enumerate(headers, start=1):
        cell = ws2.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="305496")

    row = 2
    for h in history:
        diffs = h.get("diffs_detailles", [])
        for d in diffs:
            ws2.cell(row=row, column=1, value=f"v{h['numero_version']}")
            ws2.cell(row=row, column=2, value=d["champ"])
            ws2.cell(row=row, column=3, value=str(d["avant"]))
            ws2.cell(row=row, column=4, value=str(d["apres"]))
            row += 1

    for col, w in [("A", 10), ("B", 20), ("C", 50), ("D", 50)]:
        ws2.column_dimensions[col].width = w

    wb.save(output_path)
    return output_path


def export_anomalies_to_excel(anomalies: Dict[str, Dict],
                              output_path: Path) -> Path:
    """Exporte toutes les anomalies dans un classeur multi-feuilles."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()
    ws = wb.active
    ws.title = "Synthese"

    # Page de synthèse
    ws["A1"] = "Rapport d'anomalies — Praxedo"
    ws["A1"].font = Font(bold=True, size=16, color="305496")
    ws.merge_cells("A1:D1")
    ws["A2"] = f"Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}"
    ws["A2"].font = Font(italic=True, color="666666")

    headers = ["Gravité", "Catégorie", "Description", "Nombre"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="305496")

    severity_map = {"error": "🔴 ERREUR", "warning": "🟠 ATTENTION", "info": "🔵 INFO"}
    for i, (key, data) in enumerate(anomalies.items(), start=5):
        ws.cell(row=i, column=1, value=severity_map.get(data["severity"], ""))
        ws.cell(row=i, column=2, value=data["title"])
        ws.cell(row=i, column=3, value=data["description"])
        ws.cell(row=i, column=4, value=len(data["items"]))

    for col, w in [("A", 16), ("B", 30), ("C", 55), ("D", 10)]:
        ws.column_dimensions[col].width = w

    # Une feuille par type d'anomalie
    for key, data in anomalies.items():
        if not data["items"]:
            continue
        # Nom de feuille limité à 31 caractères
        sheet_name = data["title"][:31]
        ws2 = wb.create_sheet(sheet_name)
        items = data["items"]
        if items:
            keys = list(items[0].keys())
            for c, k in enumerate(keys, start=1):
                cell = ws2.cell(row=1, column=c, value=k)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="305496")
            for r_idx, item in enumerate(items, start=2):
                for c_idx, k in enumerate(keys, start=1):
                    ws2.cell(row=r_idx, column=c_idx, value=item.get(k))
            for c in range(1, len(keys) + 1):
                letter = ws2.cell(row=1, column=c).column_letter
                ws2.column_dimensions[letter].width = 20
            ws2.freeze_panes = "A2"

    wb.save(output_path)
    return output_path
