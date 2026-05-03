#!/usr/bin/env python3
"""
Praxedo Core - Moteur de traitement
====================================
Contient toute la logique d'extraction et d'enrichissement d'un export
Praxedo. Utilisé par :
  - praxedo_gui.py   (interface Tkinter glisser-déposer + batch)
  - praxedo_cli.py   (ligne de commande)

Colonnes produites :
  21 POI              - SBR/STM/BRS + 6 chiffres
  22 Tipologie        - STD / NSTD (avec repli par ND col 15)
  23 Standard         - 1 / 0
  24 Non Standard     - 1 / 0
  25 ARC              - 1 / 0
  26 Refint           - référence interne
  27 Type Production  - ex. TIRAGE STD +300m
  28 CAFF Nom         - nom du CAFF
  29 CAFF Tel         - téléphone du CAFF
  30 RAI Nom          - nom du RAI
  31 RAI Tel          - téléphone du RAI
  32 Email RPE        - email *.orange.com
  33 Téléphones       - liste des tél. FR détectés dans la description
"""

from typing import Any, Dict, List, Optional, Tuple, Union

import re
import csv
import time
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

try:
    from praxedo_db import PraxedoDB, row_to_data
    DB_AVAILABLE = True
except ImportError:
    DB_AVAILABLE = False


# =============================================================================
# REGEX
# =============================================================================

POI_RE   = re.compile(r"\b(SBR|STM|BRS)\s*(\d{6})\b", re.IGNORECASE)
NSTD_RE  = re.compile(r"\b(NSTD|NON\s*STD|NON\s*STANDARD|NONSTANDARD)\b", re.IGNORECASE)
STD_RE   = re.compile(r"(?<!NON)(?<!NON\s)\b(STD|STANDARD)\b", re.IGNORECASE)
ARC_RE   = re.compile(r"\bARC\b", re.IGNORECASE)

# Refint : précédé de "Refint:" ou apparaissant dans une séquence type
# "PRVFTO0059GEB2-2672193" ou "CMEFTO0059CMG1-2648079=1"
REFINT_LABEL_RE = re.compile(
    r"Refint\s*[:\-]?\s*([A-Z0-9\-=]+)",
    re.IGNORECASE,
)

# Type de Production : "Type de Production TIRAGE STD +300m" ou
# "TIRAGE NSTD -300 / C2E" - on capture jusqu'à la fin de ligne ou / ou RDV
TYPE_PROD_RE = re.compile(
    r"Type\s*de\s*Production\s*[:\-]?\s*([^/\n\r]+?)(?=\s*/|\s*RDV|\s*Prestation|\s*$)",
    re.IGNORECASE,
)

# CAFF : "CAFF : Christelle MBAJON 0608188638" ou "CAF David Bleas 0785247129"
# ou "CAFF : MICHEL Dimitri  +33 677894849"
# \b en début pour éviter de matcher CAFF dans un autre mot
CAFF_RE = re.compile(
    r"\bCAFF?\s*[:\-]?\s*([A-Za-zÀ-ÿ'\-\.\s]+?)\s*"
    r"(\+?33\s?\d(?:[\s\.\-]?\d){8}|0\d(?:[\s\.\-]?\d){8})",
    re.IGNORECASE,
)

# RAI Nom : "Nom RAI : GUESMI Souhail"
RAI_NOM_RE = re.compile(
    r"Nom\s*RAI\s*[:\-]?\s*([A-Za-zÀ-ÿ'\-\.\s]+?)(?=\s*/|\s*Tel|\s*\n|\s*-{2,}|$)",
    re.IGNORECASE,
)
# RAI Tel : "Tel. RAI : 03 67 22 04 07" ou "Tel.RAI:0426837314"
RAI_TEL_RE = re.compile(
    r"Tel\.?\s*RAI\s*[:\-]?\s*(\+?(?:33)?[\s\.\-]?\d(?:[\s\.\-]?\d){8,9})",
    re.IGNORECASE,
)

# Email RPE : "Email RPE : xxx@orange.com"
# On limite le domaine aux TLD standards (com|fr|org|net) pour éviter de
# coller le mot suivant (ex: "orange.comRealisation").
EMAIL_RPE_RE = re.compile(
    r"Email\s*RPE\s*[:\-]?\s*"
    r"([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+?\.(?:com|fr|org|net))"
    r"(?![A-Za-z])",
    re.IGNORECASE,
)
EMAIL_ORANGE_RE = re.compile(
    r"([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]*orange\.(?:com|fr))(?![A-Za-z])",
    re.IGNORECASE,
)

# Téléphones FR : fixe ou mobile, tolère espaces, points, tirets, +33
# On veut exactement 10 chiffres (hors préfixe) ou 9 chiffres après +33
PHONE_RE = re.compile(
    r"(?<!\d)"                                   # pas de chiffre juste avant
    r"(?:\+?33[\s\.\-]?)?"                       # préfixe optionnel +33
    r"(?:0|(?<=\+33[\s\.\-])|(?<=\+33))"         # soit un 0, soit juste après +33
    r"[1-9]"                                     # 1er chiffre non nul
    r"(?:[\s\.\-]?\d){8}"                        # 8 chiffres restants
    r"(?!\d)"                                    # pas de chiffre juste après
)


# =============================================================================
# HELPERS D'EXTRACTION
# =============================================================================

def _s(v) -> str:
    return "" if v is None else str(v)


def extract_poi(text: str) -> str:
    if not text:
        return ""
    seen, unique = set(), []
    for pref, num in POI_RE.findall(str(text)):
        token = f"{pref.upper()}{num}"
        if token not in seen:
            seen.add(token)
            unique.append(token)
    return " / ".join(unique)


def extract_tipologie(text: str) -> str:
    if not text:
        return ""
    s = str(text)
    if NSTD_RE.search(s):
        return "NSTD"
    if STD_RE.search(s):
        return "STD"
    return ""


def has_arc(text: str) -> int:
    return 1 if text and ARC_RE.search(str(text)) else 0


def extract_refint(text: str) -> str:
    if not text:
        return ""
    m = REFINT_LABEL_RE.search(str(text))
    if m:
        # On nettoie : on enlève d'éventuels caractères non-ref en fin
        ref = m.group(1).strip().rstrip(".,;")
        return ref
    return ""


def extract_type_production(text: str) -> str:
    if not text:
        return ""
    m = TYPE_PROD_RE.search(str(text))
    if m:
        return re.sub(r"\s+", " ", m.group(1)).strip(" -:")
    return ""


def _clean_phone(raw: str) -> str:
    """Normalise un numéro FR en '+33 X XX XX XX XX' ou garde l'original compact."""
    if not raw:
        return ""
    digits = re.sub(r"\D", "", raw)
    # +33 X... -> on enlève le préfixe 33
    if digits.startswith("33") and len(digits) == 11:
        digits = "0" + digits[2:]
    if len(digits) == 10 and digits.startswith("0"):
        return f"{digits[0:2]} {digits[2:4]} {digits[4:6]} {digits[6:8]} {digits[8:10]}"
    return raw.strip()


def extract_caff(text: str) -> Tuple[str, str]:
    if not text:
        return "", ""
    m = CAFF_RE.search(str(text))
    if m:
        nom = re.sub(r"\s+", " ", m.group(1)).strip(" -:")
        # Nettoie le nom (enlève "CAFF", ":" éventuels résiduels)
        nom = re.sub(r"^(CAFF?|:|\-)\s*", "", nom, flags=re.IGNORECASE).strip()
        tel = _clean_phone(m.group(2))
        return nom, tel
    return "", ""


def extract_rai(text: str) -> Tuple[str, str]:
    if not text:
        return "", ""
    nom, tel = "", ""
    mn = RAI_NOM_RE.search(str(text))
    if mn:
        nom = re.sub(r"\s+", " ", mn.group(1)).strip(" -:/")
    mt = RAI_TEL_RE.search(str(text))
    if mt:
        tel = _clean_phone(mt.group(1))
    return nom, tel


def extract_email_rpe(text: str) -> str:
    if not text:
        return ""
    s = str(text)
    m = EMAIL_RPE_RE.search(s)
    if m:
        return m.group(1).lower()
    # Fallback : n'importe quel email @orange.com ou @orange.fr
    m = EMAIL_ORANGE_RE.search(s)
    if m:
        return m.group(1).lower()
    return ""


def extract_phones(text: str) -> str:
    """Retourne tous les numéros FR uniques joints par ' / '."""
    if not text:
        return ""
    found = PHONE_RE.findall(str(text))
    # findall renvoie les groupes ; on refait une recherche pour obtenir la chaîne
    phones = [m.group(0) for m in PHONE_RE.finditer(str(text))]
    seen, unique = set(), []
    for raw in phones:
        cleaned = _clean_phone(raw)
        digits = re.sub(r"\D", "", cleaned)
        if digits and digits not in seen and len(digits) == 10:
            seen.add(digits)
            unique.append(cleaned)
    return " / ".join(unique)


# =============================================================================
# STRUCTURE DES COLONNES DE SORTIE
# =============================================================================

OUTPUT_COLUMNS = [
    (21, "POI",              28),
    (22, "Tipologie",        14),
    (23, "Standard",         12),
    (24, "Non Standard",     16),
    (25, "ARC",               8),
    (26, "Refint",           28),
    (27, "Type Production",  28),
    (28, "CAFF Nom",         24),
    (29, "CAFF Tel",         18),
    (30, "RAI Nom",          22),
    (31, "RAI Tel",          18),
    (32, "Email RPE",        32),
    (33, "Telephones",       32),
]


# =============================================================================
# TRAITEMENT PRINCIPAL
# =============================================================================

def process_workbook(input_path: Path,
                     output_path: Path,
                     missing_poi_csv: Optional[Path] = None,
                     db: "Optional[PraxedoDB]" = None) -> dict:
    """Traite un fichier et retourne des statistiques.

    Si `db` est fourni, les interventions sont aussi inserées en base.
    """
    t0 = time.time()
    wb = load_workbook(input_path, keep_vba=False, data_only=False)
    ws = wb.active
    max_row = ws.max_row

    # --- En-têtes -----------------------------------------------------------
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="305496")
    header_align = Alignment(horizontal="center", vertical="center")

    for col, title, width in OUTPUT_COLUMNS:
        cell = ws.cell(row=1, column=col, value=title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        ws.column_dimensions[cell.column_letter].width = width

    stats = dict(rows=0, poi=0, tipo_direct=0, tipo_fallback=0, arc=0,
                 refint=0, type_prod=0, caff=0, rai=0, email=0, phones=0)
    missing_poi_rows = []

    tipo_by_row = {}
    rows_by_nd = {}

    # --- Premier passage ---------------------------------------------------
    for row in range(2, max_row + 1):
        desc = ws.cell(row=row, column=17).value
        nd   = ws.cell(row=row, column=15).value
        stats["rows"] += 1

        # POI
        poi = extract_poi(desc)
        ws.cell(row=row, column=21, value=poi)
        if poi:
            stats["poi"] += 1
        else:
            missing_poi_rows.append({
                "ligne_excel": row,
                "N°":          _s(ws.cell(row=row, column=1).value),
                "ND (col 15)": _s(nd),
                "Client":      _s(ws.cell(row=row, column=14).value),
                "Description (aperçu)": _s(desc)[:200].replace("\n", " "),
            })

        # Tipologie directe
        tipo = extract_tipologie(desc)
        tipo_by_row[row] = tipo
        if tipo:
            stats["tipo_direct"] += 1

        # ARC
        arc_val = has_arc(desc)
        ws.cell(row=row, column=25, value=arc_val)
        if arc_val:
            stats["arc"] += 1

        # Refint
        ref = extract_refint(desc)
        ws.cell(row=row, column=26, value=ref)
        if ref:
            stats["refint"] += 1

        # Type Production
        tprod = extract_type_production(desc)
        ws.cell(row=row, column=27, value=tprod)
        if tprod:
            stats["type_prod"] += 1

        # CAFF
        caff_nom, caff_tel = extract_caff(desc)
        ws.cell(row=row, column=28, value=caff_nom)
        ws.cell(row=row, column=29, value=caff_tel)
        if caff_nom or caff_tel:
            stats["caff"] += 1

        # RAI
        rai_nom, rai_tel = extract_rai(desc)
        ws.cell(row=row, column=30, value=rai_nom)
        ws.cell(row=row, column=31, value=rai_tel)
        if rai_nom or rai_tel:
            stats["rai"] += 1

        # Email RPE
        email = extract_email_rpe(desc)
        ws.cell(row=row, column=32, value=email)
        if email:
            stats["email"] += 1

        # Téléphones
        phones = extract_phones(desc)
        ws.cell(row=row, column=33, value=phones)
        if phones:
            stats["phones"] += 1

        # Index par ND pour le repli tipologie
        if nd is not None and str(nd).strip():
            rows_by_nd.setdefault(str(nd).strip(), []).append(row)

    # --- Deuxième passage : repli tipologie ---------------------------------
    for row in range(2, max_row + 1):
        if tipo_by_row[row]:
            continue
        nd = ws.cell(row=row, column=15).value
        if nd is None or not str(nd).strip():
            continue
        for other in rows_by_nd.get(str(nd).strip(), []):
            if other != row and tipo_by_row.get(other):
                tipo_by_row[row] = tipo_by_row[other]
                stats["tipo_fallback"] += 1
                break

    # --- Écriture Tipologie + Standard / Non Standard -----------------------
    center = Alignment(horizontal="center", vertical="center")
    for row in range(2, max_row + 1):
        tipo = tipo_by_row.get(row, "")
        ws.cell(row=row, column=22, value=tipo)
        ws.cell(row=row, column=23, value=1 if tipo == "STD" else 0)
        ws.cell(row=row, column=24, value=1 if tipo == "NSTD" else 0)
        for col in (22, 23, 24, 25):
            ws.cell(row=row, column=col).alignment = center

    ws.freeze_panes = "A2"
    wb.save(output_path)

    # --- Rapport CSV des lignes sans POI ------------------------------------
    if missing_poi_csv is not None and missing_poi_rows:
        with open(missing_poi_csv, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=list(missing_poi_rows[0].keys()),
                                    delimiter=";")
            writer.writeheader()
            writer.writerows(missing_poi_rows)

    stats["missing_poi"] = len(missing_poi_rows)
    stats["missing_poi_csv"] = str(missing_poi_csv) if missing_poi_csv and missing_poi_rows else ""

    # --- Alimentation de la base de données ---------------------------------
    db_stats = {"db_new": 0, "db_updated": 0, "db_identical": 0, "db_skipped": 0}
    if db is not None:
        fid, _is_new = db.register_file(input_path)
        import_id = db.start_import(fid)

        for row in range(2, max_row + 1):
            row_values = {c: ws.cell(row=row, column=c).value for c in range(1, 34)}
            tipologie_repli = bool(
                tipo_by_row.get(row) and
                not extract_tipologie(ws.cell(row=row, column=17).value)
            )
            data = row_to_data(row_values, tipologie_par_repli=tipologie_repli)
            result = db.upsert_intervention(data, import_id)
            if result == "new":
                db_stats["db_new"] += 1
            elif result == "updated":
                db_stats["db_updated"] += 1
            elif result == "identical":
                db_stats["db_identical"] += 1
            else:
                db_stats["db_skipped"] += 1

        duree_ms = int((time.time() - t0) * 1000)
        db.finalize_import(import_id, {
            "nb_lignes_lues":  stats["rows"],
            "nb_nouvelles":    db_stats["db_new"],
            "nb_mises_a_jour": db_stats["db_updated"],
            "nb_identiques":   db_stats["db_identical"],
            "nb_erreurs":      db_stats["db_skipped"],
            "duree_ms":        duree_ms,
            "commentaire":     f"POI={stats['poi']} ARC={stats['arc']}",
        })

    stats.update(db_stats)
    return stats


def process_batch(input_files: List[Path], output_dir: Path,
                  generate_missing_csv: bool = True,
                  progress_callback=None,
                  db_path: Optional[Path] = None) -> List[dict]:
    """Traite plusieurs fichiers. Retourne une liste de dicts (stats par fichier).

    Si `db_path` est fourni, alimente la base SQLite.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    results = []
    total = len(input_files)

    db = None
    if db_path is not None and DB_AVAILABLE:
        db = PraxedoDB(db_path)
        try:
            backup_dir = db_path.parent / "backups"
            db.backup(backup_dir, keep=10)
        except Exception:  # noqa: BLE001
            pass

    for idx, in_path in enumerate(input_files, start=1):
        out_path = output_dir / f"{in_path.stem}_traite.xlsx"
        miss_path = (output_dir / f"{in_path.stem}_sans_POI.csv") if generate_missing_csv else None
        try:
            stats = process_workbook(in_path, out_path, miss_path, db=db)
            stats["status"] = "OK"
            stats["input"] = str(in_path)
            stats["output"] = str(out_path)
        except Exception as exc:   # noqa: BLE001
            stats = {"status": "ERREUR", "input": str(in_path),
                     "output": "", "error": str(exc)}
        results.append(stats)
        if progress_callback:
            progress_callback(idx, total, in_path.name, stats)

    if db is not None:
        db.close()
    return results
