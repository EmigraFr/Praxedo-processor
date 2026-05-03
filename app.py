#!/usr/bin/env python3
"""
Praxedo Processor — Version Web (Streamlit)
============================================
App web locale 100% fonctionnelle sans tkinter.

Lancement :
    streamlit run app.py
    (le navigateur s'ouvre automatiquement sur http://localhost:8501)

Fonctionnalités (toutes) :
  • Traitement par lot de fichiers Excel
  • Alimentation base SQLite avec historisation
  • Recherche multi-critères + tri
  • Détail d'intervention + historique des versions
  • Export Excel des résultats filtrés
  • Rapport CSV des lignes sans POI
"""

from __future__ import annotations
from typing import Any, Dict, List, Optional, Tuple

import io
import os
import sys
import time
import tempfile
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# Modules locaux
from praxedo_core import process_workbook
from praxedo_db import PraxedoDB
from praxedo_queries import (PraxedoQueries, DISPLAY_COLUMNS,
                             export_rows_to_xlsx)
from praxedo_dashboard import (
    PERIODS, compute_kpi, timeseries, repartition_tipo, stats_par_agence,
    top_techniciens, top_caff, top_clients, nd_recurrents, alertes,
    get_distinct_agences, fig_timeseries, fig_donut_tipo,
    fig_bar_agence, fig_top_bar, generate_pdf_report,
)
from praxedo_search import (
    advanced_search, get_history_detailed, detect_anomalies,
    anomalies_summary, export_search_to_excel, export_rows_to_csv_bytes,
    export_history_to_excel, export_anomalies_to_excel,
    ADVANCED_FIELDS, FIELD_BY_KEY,
    TEXT_OPERATORS, BOOL_OPERATORS, DATE_OPERATORS,
)


# =============================================================================
# CONFIGURATION GLOBALE DE LA PAGE
# =============================================================================

st.set_page_config(
    page_title="Praxedo Processor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# STYLE CSS
# =============================================================================

st.markdown("""
<style>
    /* En-tête principal */
    .main-header {
        background: linear-gradient(135deg, #1f4e79 0%, #305496 100%);
        color: white;
        padding: 24px 32px;
        border-radius: 10px;
        margin-bottom: 18px;
        box-shadow: 0 4px 14px rgba(31,78,121,.25);
    }
    .main-header h1 { margin: 0; font-size: 1.9em; }
    .main-header .subtitle { opacity: 0.92; margin-top: 6px; }

    /* Cartes KPI */
    .kpi-card {
        background: white;
        border-left: 5px solid #305496;
        border-radius: 6px;
        padding: 14px 18px;
        box-shadow: 0 1px 3px rgba(0,0,0,.08);
    }
    .kpi-value {
        font-size: 2em;
        font-weight: 700;
        color: #1f4e79;
        line-height: 1;
    }
    .kpi-label {
        color: #666;
        font-size: 0.9em;
        margin-top: 4px;
    }

    /* Sous-titres */
    h3 { color: #1f4e79; }

    /* Tableau de données plus lisible */
    .dataframe th {
        background: #305496 !important;
        color: white !important;
        padding: 8px !important;
    }

    /* Petits ajustements */
    .stButton>button {
        font-weight: 600;
    }
    .stDownloadButton>button {
        background: #305496;
        color: white;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)


# =============================================================================
# ÉTAT DE SESSION
# =============================================================================

DEFAULT_DB_DIR = Path.home() / "Praxedo_Resultats"
DEFAULT_DB_DIR.mkdir(parents=True, exist_ok=True)
DEFAULT_DB_PATH = DEFAULT_DB_DIR / "praxedo.db"

if "db_path" not in st.session_state:
    st.session_state.db_path = str(DEFAULT_DB_PATH)
if "output_dir" not in st.session_state:
    st.session_state.output_dir = str(DEFAULT_DB_DIR)
if "processed_files" not in st.session_state:
    st.session_state.processed_files = []
if "search_results" not in st.session_state:
    st.session_state.search_results = None
if "detail_id" not in st.session_state:
    st.session_state.detail_id = None
if "adv_criteria" not in st.session_state:
    st.session_state.adv_criteria = [
        {"field": "tipologie", "operator": "égal à", "value": ""},
    ]
if "adv_results" not in st.session_state:
    st.session_state.adv_results = None
if "adv_history_refint" not in st.session_state:
    st.session_state.adv_history_refint = None


# =============================================================================
# FONCTIONS UTILITAIRES
# =============================================================================

def fmt_date(s: Optional[str]) -> str:
    if not s:
        return "—"
    try:
        return datetime.fromisoformat(s).strftime("%d/%m/%Y %H:%M")
    except (ValueError, TypeError):
        return str(s)


def fmt_size(n: int) -> str:
    for unit in ("o", "Ko", "Mo", "Go"):
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} To"


def get_db_summary() -> Optional[Dict]:
    """Renvoie un dict de stats DB, ou None si la base n'existe pas."""
    dbp = Path(st.session_state.db_path)
    if not dbp.exists():
        return None
    try:
        db = PraxedoDB(dbp)
        s = db.summary()
        db.close()
        return s
    except Exception:
        return None


def render_header():
    st.markdown("""
    <div class="main-header">
        <h1>📊 Praxedo Processor</h1>
        <div class="subtitle">Version Web v3.0 — Traitement + Base de données + Recherche</div>
    </div>
    """, unsafe_allow_html=True)


def render_sidebar():
    """Barre latérale : paramètres + stats DB."""
    st.sidebar.markdown("### ⚙️ Paramètres")

    st.sidebar.text_input(
        "📁 Dossier de sortie",
        value=st.session_state.output_dir,
        key="output_dir",
        help="Les fichiers traités (.xlsx, .csv) seront placés ici.",
    )

    st.sidebar.text_input(
        "💾 Chemin de la base",
        value=st.session_state.db_path,
        key="db_path",
        help="Fichier .db où sont stockées les interventions cumulées.",
    )

    st.sidebar.divider()
    st.sidebar.markdown("### 📈 État de la base")

    summary = get_db_summary()
    if summary is None:
        st.sidebar.info("💾 Aucune base pour l'instant.\n\n"
                        "Lancez un traitement pour la créer automatiquement.")
    else:
        st.sidebar.markdown(f"""
        **Interventions :** {summary['nb_interventions']:,}
        **Versions historiques :** {summary['nb_historique']:,}
        **Fichiers traités :** {summary['nb_fichiers']}
        **Imports :** {summary['nb_imports']}
        **Dernier import :** {fmt_date(summary['dernier_import'])}
        **Taille :** {fmt_size(summary['taille_octets'])}
        """.replace(",", " "))

    st.sidebar.divider()
    st.sidebar.caption("Praxedo Processor v3.0 Web Edition")
    st.sidebar.caption(f"Python {sys.version.split()[0]} • Streamlit")


# =============================================================================
# PAGE 1 : TRAITEMENT
# =============================================================================

def page_traitement():
    st.markdown("## 📥 Traitement de fichiers Excel")
    st.caption("Déposez un ou plusieurs exports Praxedo (.xlsx / .xlsm). "
               "Chaque fichier est enrichi avec 13 colonnes et ajouté à la base.")

    # Uploader multi-fichiers
    uploaded_files = st.file_uploader(
        "Glissez-déposez vos fichiers ici (ou cliquez pour parcourir)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        help="Les fichiers sont traités localement, rien n'est envoyé sur internet.",
    )

    # Options
    col1, col2 = st.columns(2)
    with col1:
        gen_csv = st.checkbox(
            "📋 Générer un rapport CSV des lignes sans POI",
            value=True,
            help="Pour contrôle manuel des interventions sans code POI détecté.",
        )
    with col2:
        use_db = st.checkbox(
            "💾 Alimenter la base de données (historisation)",
            value=True,
            help="Chaque fichier traité enrichit la base cumulative.",
        )

    if not uploaded_files:
        st.info("⬆️ Déposez au moins un fichier pour commencer.")
        return

    st.success(f"**{len(uploaded_files)} fichier(s) prêt(s) à traiter**")
    for f in uploaded_files:
        st.markdown(f"- 📄 `{f.name}` ({fmt_size(f.size)})")

    # Bouton de traitement
    if st.button("🚀 Lancer le traitement", type="primary", use_container_width=True):
        out_dir = Path(st.session_state.output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        db = None
        if use_db:
            try:
                db = PraxedoDB(st.session_state.db_path)
            except Exception as e:
                st.error(f"❌ Impossible d'ouvrir la base : {e}")
                return

        progress = st.progress(0.0)
        status = st.empty()
        log_container = st.container()
        all_stats = []

        for idx, upl in enumerate(uploaded_files):
            status.info(f"⏳ Traitement de **{upl.name}** "
                        f"({idx + 1}/{len(uploaded_files)})…")

            # Écrire le fichier temporairement (openpyxl a besoin d'un chemin)
            tmp_path = out_dir / f"_tmp_{upl.name}"
            tmp_path.write_bytes(upl.getvalue())

            try:
                out_xlsx = out_dir / f"{Path(upl.name).stem}_traite.xlsx"
                out_csv = (out_dir / f"{Path(upl.name).stem}_sans_POI.csv"
                           if gen_csv else None)

                stats = process_workbook(tmp_path, out_xlsx,
                                         missing_poi_csv=out_csv, db=db)
                stats["status"] = "OK"
                stats["name"] = upl.name
                stats["output_xlsx"] = str(out_xlsx)
                stats["output_csv"] = str(out_csv) if out_csv else ""

                with log_container:
                    st.success(
                        f"✅ **{upl.name}** — "
                        f"{stats['rows']} lignes • POI : {stats['poi']} "
                        f"(manquants : {stats['missing_poi']}) • "
                        f"Tipo : {stats['tipo_direct']}+{stats['tipo_fallback']} • "
                        f"ARC : {stats['arc']} • "
                        f"CAFF : {stats['caff']} • RAI : {stats['rai']} • "
                        f"Email : {stats['email']} • Tél : {stats['phones']}"
                    )
                    if use_db and "db_new" in stats:
                        st.info(
                            f"    💾 Base : **{stats['db_new']}** nouvelles • "
                            f"**{stats['db_updated']}** mises à jour • "
                            f"**{stats['db_identical']}** identiques • "
                            f"**{stats['db_skipped']}** ignorées"
                        )
            except Exception as e:
                stats = {"status": "ERREUR", "name": upl.name, "error": str(e)}
                with log_container:
                    st.error(f"❌ **{upl.name}** — {e}")
            finally:
                if tmp_path.exists():
                    tmp_path.unlink()

            all_stats.append(stats)
            progress.progress((idx + 1) / len(uploaded_files))

        if db is not None:
            db.close()

        status.success(
            f"🎉 Terminé : "
            f"{sum(1 for s in all_stats if s.get('status') == 'OK')} OK / "
            f"{sum(1 for s in all_stats if s.get('status') != 'OK')} erreur(s)"
        )

        st.session_state.processed_files = all_stats

        # Liens de téléchargement directs
        st.markdown("---")
        st.markdown("### 📥 Téléchargement des résultats")
        for stats in all_stats:
            if stats.get("status") != "OK":
                continue
            col1, col2 = st.columns([3, 1])
            with col1:
                st.markdown(f"**{stats['name']}**")
            with col2:
                xlsx_path = Path(stats["output_xlsx"])
                if xlsx_path.exists():
                    st.download_button(
                        "📊 Excel enrichi",
                        data=xlsx_path.read_bytes(),
                        file_name=xlsx_path.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_xlsx_{stats['name']}",
                    )
            if stats.get("output_csv"):
                csv_path = Path(stats["output_csv"])
                if csv_path.exists():
                    col3, col4 = st.columns([3, 1])
                    with col3:
                        st.caption(f"    Rapport des lignes sans POI")
                    with col4:
                        st.download_button(
                            "📋 CSV sans POI",
                            data=csv_path.read_bytes(),
                            file_name=csv_path.name,
                            mime="text/csv",
                            key=f"dl_csv_{stats['name']}",
                        )


# =============================================================================
# PAGE 2 : RECHERCHE
# =============================================================================

def page_recherche():
    st.markdown("## 🔍 Recherche dans la base")
    st.caption("Filtrez les interventions, triez les résultats, exportez vers Excel.")

    db_path = Path(st.session_state.db_path)
    if not db_path.exists():
        st.warning("⚠️ Aucune base de données pour l'instant. "
                   "Lancez d'abord un traitement.")
        return

    # ===== Formulaire de filtres =====
    with st.expander("🎯 Filtres", expanded=True):
        col1, col2 = st.columns([3, 1])
        with col1:
            global_search = st.text_input(
                "🔎 Recherche globale (texte libre)",
                placeholder="Tapez n'importe quel mot : CAFF, email, POI, ND…",
                help="Cherche dans : Refint, ND, POI, description, CAFF, "
                     "RAI, email, téléphones, client, technicien.",
            )
        with col2:
            st.write("")
            st.write("")
            clear = st.button("🧹 Effacer", use_container_width=True)
            if clear:
                st.rerun()

        col1, col2, col3 = st.columns(3)
        with col1:
            f_nd = st.text_input("ND", key="f_nd")
            f_poi = st.text_input("POI", key="f_poi")
            f_refint = st.text_input("Refint", key="f_refint")
        with col2:
            f_tipo = st.selectbox("Tipologie", ["(tous)", "STD", "NSTD", ""],
                                  key="f_tipo")
            f_arc = st.selectbox("ARC", ["(tous)", "Oui", "Non"], key="f_arc")
            f_agence = st.text_input("Agence", key="f_agence")
        with col3:
            f_tech = st.text_input("Technicien", key="f_tech")
            f_caff = st.text_input("CAFF", key="f_caff")
            f_rai = st.text_input("RAI", key="f_rai")

        col1, col2 = st.columns(2)
        with col1:
            f_client = st.text_input("Client", key="f_client")
        with col2:
            f_email = st.text_input("Email RPE", key="f_email")

    # Construction des filtres
    filters: Dict[str, str] = {}
    if f_nd: filters["nd"] = f_nd
    if f_poi: filters["poi"] = f_poi
    if f_refint: filters["refint"] = f_refint
    if f_tipo and f_tipo != "(tous)":
        filters["tipologie"] = f_tipo
    if f_arc != "(tous)":
        filters["arc"] = "1" if f_arc == "Oui" else "0"
    if f_agence: filters["agence"] = f_agence
    if f_tech: filters["technicien_nom"] = f_tech
    if f_caff: filters["caff_nom"] = f_caff
    if f_rai: filters["rai_nom"] = f_rai
    if f_client: filters["client"] = f_client
    if f_email: filters["email_rpe"] = f_email

    # Bouton de recherche
    col1, col2, col3 = st.columns([1, 1, 3])
    with col1:
        do_search = st.button("🔍 Rechercher", type="primary",
                              use_container_width=True)
    with col2:
        limit = st.number_input("Limite", min_value=10, max_value=5000,
                                value=500, step=100,
                                label_visibility="collapsed")

    if do_search or st.session_state.search_results is not None:
        if do_search:
            try:
                with PraxedoQueries(db_path) as q:
                    rows, total = q.search_interventions(
                        filters=filters,
                        global_search=global_search or None,
                        limit=int(limit),
                    )
                # Convertir en DataFrame pour Streamlit
                data = []
                for r in rows:
                    data.append({k: r[k] for k in r.keys()})
                st.session_state.search_results = {
                    "data": data, "total": total, "rows": rows,
                }
            except Exception as e:
                st.error(f"❌ Erreur de recherche : {e}")
                return

        res = st.session_state.search_results
        if res is None:
            return

        total = res["total"]
        rows = res["rows"]
        data = res["data"]

        if total == 0:
            st.warning("🔎 Aucun résultat pour ces critères.")
            return

        # Compteur
        if total > len(rows):
            st.info(f"📊 **{total}** résultats trouvés — "
                    f"**{len(rows)} affichés** (limite atteinte, "
                    f"augmentez la limite ou affinez les filtres).")
        else:
            st.success(f"📊 **{total}** résultat(s) trouvé(s)")

        # Préparer DataFrame d'affichage
        df = pd.DataFrame(data)
        if not df.empty:
            # Colonnes à afficher en priorité
            preferred = [k for k, _ in DISPLAY_COLUMNS if k in df.columns]
            other_cols = [c for c in df.columns if c not in preferred and c != "id"]
            display_df = df[preferred + other_cols].copy()

            # Renommer les colonnes pour affichage
            col_labels = {k: v for k, v in DISPLAY_COLUMNS}
            display_df.rename(columns=col_labels, inplace=True)

            # Formater ARC
            if "ARC" in display_df.columns:
                display_df["ARC"] = display_df["ARC"].apply(
                    lambda x: "Oui" if x == 1 else ("" if x == 0 else x))

            # Formater dates
            for col in ("Maj", "Planifiée"):
                if col in display_df.columns:
                    display_df[col] = display_df[col].apply(fmt_date)

            # Affichage
            st.dataframe(
                display_df,
                use_container_width=True,
                hide_index=True,
                height=400,
            )

            # Boutons actions
            st.markdown("---")
            col1, col2, col3 = st.columns(3)

            with col1:
                # Export Excel
                buffer = io.BytesIO()
                tmp_xlsx = Path(tempfile.mktemp(suffix=".xlsx"))
                export_rows_to_xlsx(rows, tmp_xlsx, "Recherche Praxedo")
                buffer.write(tmp_xlsx.read_bytes())
                tmp_xlsx.unlink()
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    "📤 Exporter vers Excel",
                    data=buffer.getvalue(),
                    file_name=f"recherche_praxedo_{stamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            with col2:
                # Export CSV
                csv_buf = df.to_csv(index=False, sep=";", encoding="utf-8-sig")
                st.download_button(
                    "📋 Exporter vers CSV",
                    data=csv_buf,
                    file_name=f"recherche_praxedo_{stamp}.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

            with col3:
                # Sélecteur détail
                if len(rows) > 0:
                    options = {
                        f"{r['refint']} ({r['nd'] or '—'})": r["id"]
                        for r in rows
                    }
                    sel = st.selectbox("📄 Voir le détail d'une intervention",
                                       options=[""] + list(options.keys()),
                                       label_visibility="collapsed")
                    if sel:
                        st.session_state.detail_id = options[sel]

        # Affichage du détail
        if st.session_state.detail_id:
            st.markdown("---")
            render_detail(db_path, st.session_state.detail_id)


def render_detail(db_path: Path, intervention_id: int):
    """Affiche le détail d'une intervention + son historique."""
    try:
        with PraxedoQueries(db_path) as q:
            inter = q.get_intervention(intervention_id)
            if not inter:
                st.warning("Intervention introuvable.")
                return
            history = q.get_history(inter["refint"])
    except Exception as e:
        st.error(f"Erreur : {e}")
        return

    nb_ver = inter["nb_versions"] or 1
    st.markdown(f"### 📄 {inter['refint']}")
    st.caption(f"Version {nb_ver} — {len(history)} version(s) dans l'historique")

    tab_champs, tab_hist = st.tabs([
        f"📋 Champs ({sum(1 for k in inter.keys() if inter[k] is not None)})",
        f"📜 Historique ({len(history)})",
    ])

    with tab_champs:
        # Regrouper les champs par catégorie
        groups = {
            "📋 Identifiants": ["refint", "num_ot", "nd", "code_intervention"],
            "🏷️ Enrichi": ["poi", "tipologie", "standard", "non_standard",
                          "arc", "type_production"],
            "👥 Contacts": ["caff_nom", "caff_tel", "rai_nom", "rai_tel",
                           "email_rpe", "telephones"],
            "🏢 Métier": ["agence", "type_inter", "ui", "client",
                         "technicien_nom", "technicien_prenom", "equipiers",
                         "statut", "act_prod", "soldee", "validee_annulee"],
            "📅 Dates": ["date_creation", "date_cible_initiale",
                        "a_faire_avant", "date_planifiee"],
            "📝 Description": ["description"],
            "⚙️ Méta": ["nb_versions", "created_at", "updated_at",
                       "tipologie_par_repli"],
        }

        for group_name, fields in groups.items():
            vals = []
            for f in fields:
                if f in inter.keys() and inter[f] not in (None, ""):
                    v = inter[f]
                    if isinstance(v, int) and f in ("standard", "non_standard",
                                                     "arc", "tipologie_par_repli"):
                        v = "Oui" if v else "Non"
                    if f in ("created_at", "updated_at"):
                        v = fmt_date(v)
                    vals.append((f, v))
            if not vals:
                continue

            with st.expander(group_name, expanded=(group_name != "📝 Description"
                                                    and group_name != "⚙️ Méta")):
                for field, val in vals:
                    label = field.replace("_", " ").title()
                    col1, col2 = st.columns([1, 3])
                    col1.markdown(f"**{label}**")
                    if len(str(val)) > 100:
                        col2.text_area("", value=str(val), height=100,
                                       label_visibility="collapsed",
                                       key=f"det_{intervention_id}_{field}",
                                       disabled=True)
                    else:
                        col2.markdown(f"`{val}`")

    with tab_hist:
        if not history:
            st.info("Aucune version historique.")
        else:
            hist_data = []
            for h in history:
                hist_data.append({
                    "Version": f"v{h['numero_version']}" + (
                        " ⭐" if h["est_version_courante"] else ""),
                    "Date": fmt_date(h["created_at"]),
                    "Fichier source": h["source_file"] or "—",
                    "Champs modifiés": h["champs_modifies"] or "(version initiale)",
                    "Tipologie": h["tipologie"] or "—",
                    "POI": h["poi"] or "—",
                    "Statut": h["statut"] or "—",
                })
            st.dataframe(pd.DataFrame(hist_data),
                         use_container_width=True, hide_index=True)
            st.caption("⭐ = version courante active")


# =============================================================================
# PAGE 3 : INFOS BASE
# =============================================================================

def page_infos():
    st.markdown("## 📊 État détaillé de la base")

    summary = get_db_summary()
    if summary is None:
        st.warning("Aucune base de données pour l'instant.")
        return

    # KPI en 4 colonnes
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🏷️ Interventions", f"{summary['nb_interventions']:,}".replace(",", " "))
    c2.metric("📜 Versions historiques", f"{summary['nb_historique']:,}".replace(",", " "))
    c3.metric("📁 Fichiers traités", summary["nb_fichiers"])
    c4.metric("💾 Taille", fmt_size(summary["taille_octets"]))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("✅ Standards", summary["nb_std"])
    c2.metric("⚠️ Non Standards", summary["nb_nstd"])
    c3.metric("🔶 ARC détectés", summary["nb_arc"])
    c4.metric("🔄 Imports effectués", summary["nb_imports"])

    st.markdown("---")

    # Journal des imports
    db_path = Path(st.session_state.db_path)
    try:
        conn = sqlite3.connect(str(db_path))
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT i.date_import, f.nom_fichier, i.nb_lignes_lues, "
            "i.nb_nouvelles, i.nb_mises_a_jour, i.nb_identiques, i.duree_ms "
            "FROM imports i LEFT JOIN fichiers_sources f "
            "ON f.id = i.fichier_source_id "
            "ORDER BY i.date_import DESC LIMIT 50"
        ).fetchall()
        conn.close()

        if rows:
            st.markdown("### 📋 Derniers imports")
            data = []
            for r in rows:
                data.append({
                    "Date": fmt_date(r["date_import"]),
                    "Fichier": r["nom_fichier"] or "?",
                    "Lignes": r["nb_lignes_lues"] or 0,
                    "Nouvelles": r["nb_nouvelles"] or 0,
                    "Mises à jour": r["nb_mises_a_jour"] or 0,
                    "Identiques": r["nb_identiques"] or 0,
                    "Durée (ms)": r["duree_ms"] or 0,
                })
            st.dataframe(pd.DataFrame(data),
                         use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Impossible de lire le journal : {e}")


# =============================================================================
# APPLICATION PRINCIPALE
# =============================================================================

# =============================================================================
# PAGE 4 : TABLEAU DE BORD
# =============================================================================

def page_dashboard():
    st.markdown("## 📈 Tableau de bord")
    st.caption("Indicateurs clés, graphiques et export PDF pour reporting externe.")

    db_path = Path(st.session_state.db_path)
    if not db_path.exists():
        st.warning("⚠️ Aucune base de données. Lancez d'abord un traitement.")
        return

    # ===== Filtres du dashboard =====
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        period_label = st.selectbox(
            "📅 Période",
            options=list(PERIODS.keys()),
            index=list(PERIODS.keys()).index("Toute la base"),
        )
        period_days = PERIODS[period_label]

    conn = sqlite3.connect(str(db_path))
    try:
        agences_list = get_distinct_agences(conn)
    except Exception:
        agences_list = []

    with col2:
        agence_filter = st.selectbox(
            "🏢 Agence",
            options=["(toutes)"] + agences_list,
        )
        agence_filter = None if agence_filter == "(toutes)" else agence_filter

    with col3:
        st.write("")
        st.write("")
        refresh = st.button("🔄 Rafraîchir", use_container_width=True)

    # Calculer KPI
    try:
        kpi = compute_kpi(conn, period_days=period_days, agence=agence_filter)
    except Exception as e:
        st.error(f"Erreur calcul KPI : {e}")
        conn.close()
        return

    # ===== BLOC KPI =====
    st.markdown("### 📊 Indicateurs clés")

    def delta_str(trend):
        if trend is None:
            return None
        sign = "+" if trend > 0 else ""
        return f"{sign}{trend}"

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total interventions", f"{kpi['total']:,}".replace(",", " "),
              delta=delta_str(kpi["trend_total"]))
    k2.metric("Taux STD", f"{kpi['taux_std']} %")
    k3.metric("Standards", kpi["std"], delta=delta_str(kpi["trend_std"]))
    k4.metric("Non Standards", kpi["nstd"])

    k5, k6, k7, k8 = st.columns(4)
    k5.metric("ARC détectés", kpi["arc"], delta=delta_str(kpi["trend_arc"]))
    k6.metric("POI manquants", kpi["no_poi"],
              delta="⚠️" if kpi["no_poi"] > 0 else None,
              delta_color="inverse")
    k7.metric("Doublons (vers. >1)", kpi["doublons"])
    k8.metric("Taux NSTD", f"{kpi['taux_nstd']} %")

    st.markdown("---")

    # ===== BLOC GRAPHIQUES =====
    st.markdown("### 📈 Graphiques")

    # Ligne 1 : timeseries + donut
    c1, c2 = st.columns([2, 1])
    with c1:
        ts_data = timeseries(conn, period_days=period_days or 90,
                             granularity="week")
        st.plotly_chart(fig_timeseries(ts_data,
                                        title="Évolution par semaine"),
                         use_container_width=True)
    with c2:
        tipo_data = repartition_tipo(conn, period_days=period_days)
        st.plotly_chart(fig_donut_tipo(tipo_data), use_container_width=True)

    # Ligne 2 : barres par agence
    agences_data = stats_par_agence(conn, period_days=period_days)
    st.plotly_chart(fig_bar_agence(agences_data), use_container_width=True)

    # Ligne 3 : tops
    c1, c2 = st.columns(2)
    with c1:
        techs = top_techniciens(conn, n=10, period_days=period_days)
        st.plotly_chart(
            fig_top_bar(techs, title="🏆 Top 10 techniciens", color="#305496"),
            use_container_width=True,
        )
    with c2:
        caffs = top_caff(conn, n=10, period_days=period_days)
        st.plotly_chart(
            fig_top_bar(caffs, title="🏆 Top 10 CAFF", color="#2e7d32"),
            use_container_width=True,
        )

    c1, c2 = st.columns(2)
    with c1:
        clients = top_clients(conn, n=10, period_days=period_days)
        st.plotly_chart(
            fig_top_bar(clients, title="🏆 Top 10 clients", color="#6a1b9a"),
            use_container_width=True,
        )
    with c2:
        nds = nd_recurrents(conn, n=10)
        st.plotly_chart(
            fig_top_bar(nds, key_name="nd", title="🔴 Top 10 ND récurrents",
                         color="#c62828"),
            use_container_width=True,
        )

    st.markdown("---")

    # ===== BLOC ALERTES =====
    st.markdown("### ⚠️ Alertes et anomalies")
    try:
        al = alertes(conn)
    except Exception as e:
        st.error(f"Erreur calcul alertes : {e}")
        al = {"no_poi_recent": [], "doublons": [],
              "emails_suspects": [], "caff_sans_tel": []}

    a1, a2, a3, a4 = st.columns(4)
    a1.metric("POI manquants (7j)", len(al["no_poi_recent"]))
    a2.metric("Interv. avec versions", len(al["doublons"]))
    a3.metric("Emails non-Orange", len(al["emails_suspects"]))
    a4.metric("CAFF sans tél.", len(al["caff_sans_tel"]))

    with st.expander("🔍 Voir le détail des alertes"):
        tabs_a = st.tabs(["POI manquants", "Doublons", "Emails", "CAFF"])
        with tabs_a[0]:
            if al["no_poi_recent"]:
                st.dataframe(pd.DataFrame(al["no_poi_recent"]),
                              use_container_width=True, hide_index=True)
            else:
                st.info("Aucune alerte.")
        with tabs_a[1]:
            if al["doublons"]:
                st.dataframe(pd.DataFrame(al["doublons"]),
                              use_container_width=True, hide_index=True)
            else:
                st.info("Aucun doublon.")
        with tabs_a[2]:
            if al["emails_suspects"]:
                st.dataframe(pd.DataFrame(al["emails_suspects"]),
                              use_container_width=True, hide_index=True)
            else:
                st.info("Aucun email suspect.")
        with tabs_a[3]:
            if al["caff_sans_tel"]:
                st.dataframe(pd.DataFrame(al["caff_sans_tel"]),
                              use_container_width=True, hide_index=True)
            else:
                st.info("Aucun CAFF sans téléphone.")

    st.markdown("---")

    # ===== EXPORT PDF =====
    st.markdown("### 📄 Export rapport PDF")
    st.caption("Génère un document PDF professionnel pour diffusion externe.")

    if st.button("📄 Générer le rapport PDF", type="primary"):
        with st.spinner("Génération en cours (graphiques + mise en page)..."):
            try:
                out_dir = Path(st.session_state.output_dir)
                out_dir.mkdir(parents=True, exist_ok=True)
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                pdf_path = out_dir / f"rapport_praxedo_{stamp}.pdf"
                generate_pdf_report(
                    conn, pdf_path,
                    period_label=period_label,
                    period_days=period_days,
                    agence=agence_filter,
                )
                st.success(f"✅ Rapport généré : {pdf_path.name}")
                st.download_button(
                    "📄 Télécharger le PDF",
                    data=pdf_path.read_bytes(),
                    file_name=pdf_path.name,
                    mime="application/pdf",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"❌ Erreur de génération PDF : {e}")
                st.exception(e)

    conn.close()


# =============================================================================
# PAGE 5 : RECHERCHE AVANCÉE
# =============================================================================

def page_recherche_avancee():
    st.markdown("## 🔎 Recherche avancée")
    st.caption("Filtres combinés • Historique détaillé • Exports Excel/CSV • "
               "Tableau des anomalies")

    db_path = Path(st.session_state.db_path)
    if not db_path.exists():
        st.warning("⚠️ Aucune base de données. Lancez d'abord un traitement.")
        return

    # 3 sous-onglets : Recherche / Historique / Anomalies
    sub_tab1, sub_tab2, sub_tab3 = st.tabs([
        "🔍 Recherche multi-critères",
        "📜 Historique détaillé",
        "⚠️ Anomalies (récapitulatif)",
    ])

    with sub_tab1:
        _adv_search_tab(db_path)
    with sub_tab2:
        _adv_history_tab(db_path)
    with sub_tab3:
        _adv_anomalies_tab(db_path)


# -----------------------------------------------------------------------------
# SOUS-ONGLET 1 : RECHERCHE MULTI-CRITÈRES
# -----------------------------------------------------------------------------

def _adv_search_tab(db_path: Path):
    st.markdown("### Constructeur de requête")
    st.caption(
        "Ajoutez autant de critères que nécessaire. Chaque critère a un "
        "**champ**, un **opérateur** et une **valeur**."
    )

    # Combinateur
    col1, col2 = st.columns([1, 4])
    with col1:
        combinator = st.radio(
            "Combinaison",
            ["AND (tous)", "OR (au moins un)"],
            horizontal=True, key="adv_combinator",
        )
        combinator = "AND" if combinator.startswith("AND") else "OR"

    # Constructeur de critères dynamique
    st.markdown("#### Critères")

    criteria = st.session_state.adv_criteria
    remove_index = None

    for i, crit in enumerate(criteria):
        c1, c2, c3, c4 = st.columns([3, 2, 4, 1])

        with c1:
            field_keys = [k for k, _, _, _ in ADVANCED_FIELDS]
            field_labels = {k: lbl for k, lbl, _, _ in ADVANCED_FIELDS}
            default_idx = (field_keys.index(crit["field"])
                           if crit["field"] in field_keys else 0)
            new_field = st.selectbox(
                "Champ",
                field_keys,
                format_func=lambda k: field_labels[k],
                index=default_idx,
                key=f"adv_field_{i}",
                label_visibility="collapsed",
            )
            crit["field"] = new_field

        # Déterminer le type du champ
        ftype = FIELD_BY_KEY[new_field][1]

        with c2:
            if ftype == "text":
                ops = list(TEXT_OPERATORS.keys())
            elif ftype == "bool":
                ops = list(BOOL_OPERATORS.keys())
            else:
                ops = list(DATE_OPERATORS.keys())
            default_op = crit.get("operator", ops[0])
            if default_op not in ops:
                default_op = ops[0]
            new_op = st.selectbox(
                "Opérateur",
                ops,
                index=ops.index(default_op),
                key=f"adv_op_{i}",
                label_visibility="collapsed",
            )
            crit["operator"] = new_op

        with c3:
            needs_value = True
            if ftype == "bool":
                needs_value = False
            if ftype == "text" and new_op in ("vide", "non vide"):
                needs_value = False
            if ftype == "date" and new_op in ("7 derniers jours",
                                                "30 derniers jours",
                                                "90 derniers jours"):
                needs_value = False

            if needs_value:
                if ftype == "date" and new_op in ("après", "avant", "exactement"):
                    val = st.date_input(
                        "Valeur",
                        value=None,
                        key=f"adv_val_{i}",
                        label_visibility="collapsed",
                    )
                    crit["value"] = val.isoformat() if val else ""
                else:
                    crit["value"] = st.text_input(
                        "Valeur",
                        value=crit.get("value", ""),
                        key=f"adv_val_{i}",
                        label_visibility="collapsed",
                        placeholder="Saisissez la valeur...",
                    )
            else:
                st.caption("(aucune valeur requise)")
                crit["value"] = ""

        with c4:
            st.write("")
            if st.button("🗑", key=f"adv_del_{i}", help="Supprimer ce critère"):
                remove_index = i

    if remove_index is not None:
        criteria.pop(remove_index)
        st.rerun()

    # Boutons Ajouter / Réinitialiser
    col1, col2, col3 = st.columns([1, 1, 3])
    with col1:
        if st.button("➕ Ajouter un critère", use_container_width=True):
            criteria.append({"field": "refint",
                             "operator": "contient",
                             "value": ""})
            st.rerun()
    with col2:
        if st.button("🔄 Réinitialiser", use_container_width=True):
            st.session_state.adv_criteria = [
                {"field": "tipologie", "operator": "égal à", "value": ""},
            ]
            st.session_state.adv_results = None
            st.rerun()

    # Tri + limite
    st.markdown("#### Options")
    c1, c2, c3 = st.columns(3)
    with c1:
        order_options = {k: lbl for k, lbl, _, _ in ADVANCED_FIELDS}
        order_options["updated_at"] = "Dernière mise à jour"
        order_options["created_at"] = "Créée en base"
        order_by = st.selectbox(
            "Trier par",
            options=list(order_options.keys()),
            format_func=lambda k: order_options[k],
            index=list(order_options.keys()).index("updated_at"),
        )
    with c2:
        order_dir = st.selectbox("Ordre", ["DESC (récent)", "ASC (ancien)"])
        order_dir = "DESC" if order_dir.startswith("DESC") else "ASC"
    with c3:
        limit = st.number_input("Limite max", min_value=10,
                                max_value=10000, value=1000, step=100)

    # Lancer la recherche
    if st.button("🔍 Lancer la recherche", type="primary",
                 use_container_width=True):
        conn = sqlite3.connect(str(db_path))
        try:
            rows, total = advanced_search(
                conn, criteria=criteria, combinator=combinator,
                order_by=order_by, order_dir=order_dir,
                limit=int(limit),
            )
            st.session_state.adv_results = {
                "data": [dict(r) for r in rows],
                "rows": rows,
                "total": total,
                "criteria": list(criteria),
                "combinator": combinator,
            }
        except Exception as e:
            st.error(f"❌ Erreur de recherche : {e}")
            return
        finally:
            conn.close()

    # Affichage des résultats
    if st.session_state.adv_results:
        st.markdown("---")
        res = st.session_state.adv_results
        total = res["total"]
        rows = res["rows"]
        data = res["data"]

        if total == 0:
            st.warning("🔎 Aucun résultat pour ces critères.")
            return

        if total > len(rows):
            st.info(f"📊 **{total}** résultats — "
                    f"**{len(rows)} affichés** (augmentez la limite).")
        else:
            st.success(f"📊 **{total}** résultat(s) trouvé(s)")

        # DataFrame
        df = pd.DataFrame(data)
        if not df.empty:
            display_keys = ["refint", "nd", "poi", "tipologie", "arc",
                            "agence", "technicien_nom", "caff_nom",
                            "email_rpe", "nb_versions", "updated_at"]
            display_keys = [k for k in display_keys if k in df.columns]
            df_display = df[display_keys].copy()
            if "arc" in df_display.columns:
                df_display["arc"] = df_display["arc"].apply(
                    lambda x: "✅" if x == 1 else "")
            if "updated_at" in df_display.columns:
                df_display["updated_at"] = df_display["updated_at"].apply(
                    fmt_date)
            st.dataframe(df_display, use_container_width=True,
                         hide_index=True, height=420)

        # Boutons export
        st.markdown("#### 📤 Exporter les résultats")
        col1, col2, col3 = st.columns(3)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        with col1:
            # Export Excel avec critères
            tmp_xlsx = Path(tempfile.mktemp(suffix=".xlsx"))
            export_search_to_excel(rows, tmp_xlsx,
                                    criteria_used=res["criteria"])
            xlsx_bytes = tmp_xlsx.read_bytes()
            tmp_xlsx.unlink()
            st.download_button(
                "📊 Excel (avec critères)",
                data=xlsx_bytes,
                file_name=f"recherche_avancee_{stamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with col2:
            csv_bytes = export_rows_to_csv_bytes(rows)
            st.download_button(
                "📋 CSV (Excel FR)",
                data=csv_bytes,
                file_name=f"recherche_avancee_{stamp}.csv",
                mime="text/csv",
                use_container_width=True,
            )

        with col3:
            # Lien vers l'historique
            if len(rows) > 0:
                refint_opts = ["(sélectionner)"] + [
                    f"{r['refint']} (v{r['nb_versions']})" for r in rows
                ]
                sel = st.selectbox("Voir l'historique d'une intervention",
                                    options=refint_opts,
                                    label_visibility="collapsed")
                if sel != "(sélectionner)":
                    refint = sel.split(" (")[0]
                    st.session_state.adv_history_refint = refint
                    st.info("→ Allez dans l'onglet « 📜 Historique détaillé »")


# -----------------------------------------------------------------------------
# SOUS-ONGLET 2 : HISTORIQUE DÉTAILLÉ
# -----------------------------------------------------------------------------

def _adv_history_tab(db_path: Path):
    st.markdown("### Historique détaillé d'une intervention")
    st.caption("Visualisez toutes les versions successives d'un Refint "
               "avec comparaison avant/après de chaque champ modifié.")

    # Saisie du Refint
    col1, col2 = st.columns([3, 1])
    with col1:
        refint = st.text_input(
            "🔍 Refint à analyser",
            value=st.session_state.adv_history_refint or "",
            placeholder="ex: PRVFTO0059GEB2-2672193",
        )
    with col2:
        st.write("")
        st.write("")
        do_load = st.button("📜 Charger l'historique",
                              use_container_width=True)

    if not refint.strip():
        st.info("ℹ️ Saisissez un Refint ou sélectionnez-en un "
                 "depuis l'onglet « Recherche ».")
        return

    if do_load or st.session_state.adv_history_refint:
        st.session_state.adv_history_refint = refint.strip()

    conn = sqlite3.connect(str(db_path))
    try:
        history = get_history_detailed(conn, refint.strip())
    except Exception as e:
        st.error(f"Erreur : {e}")
        return
    finally:
        conn.close()

    if not history:
        st.warning("Aucun historique trouvé pour ce Refint.")
        return

    # Bandeau récapitulatif
    st.markdown("---")
    current = next((h for h in history if h.get("est_version_courante")),
                    history[-1])
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Versions totales", len(history))
    c2.metric("Version courante", f"v{current['numero_version']}")
    c3.metric("1ère apparition", fmt_date(history[0].get("created_at")))
    c4.metric("Dernière modif", fmt_date(current.get("created_at")))

    # Export Excel de l'historique
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    tmp_xlsx = Path(tempfile.mktemp(suffix=".xlsx"))
    export_history_to_excel(history, refint.strip(), tmp_xlsx)
    xlsx_bytes = tmp_xlsx.read_bytes()
    tmp_xlsx.unlink()
    st.download_button(
        "📄 Télécharger l'historique (Excel)",
        data=xlsx_bytes,
        file_name=f"historique_{refint.strip()}_{stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Timeline des versions
    st.markdown("### 🕰️ Chronologie des versions")

    for h in history:
        is_current = h.get("est_version_courante")
        icon = "⭐" if is_current else "🕒"
        ver_label = f"{icon} Version {h['numero_version']}"
        if is_current:
            ver_label += " — COURANTE"

        with st.expander(
                f"{ver_label}  •  {fmt_date(h.get('created_at'))}  •  "
                f"{h.get('source_file') or 'source inconnue'}",
                expanded=is_current,
        ):
            col1, col2, col3 = st.columns(3)
            col1.markdown(f"**Tipologie :** `{h.get('tipologie') or '(vide)'}`")
            col2.markdown(f"**POI :** `{h.get('poi') or '(vide)'}`")
            col3.markdown(f"**ARC :** {'✅' if h.get('arc') else '❌'}")

            col1, col2, col3 = st.columns(3)
            col1.markdown(f"**Statut :** `{h.get('statut') or '(vide)'}`")
            col2.markdown(f"**Agence :** `{h.get('agence') or '(vide)'}`")
            col3.markdown(
                f"**Technicien :** `{h.get('technicien_nom') or ''} "
                f"{h.get('technicien_prenom') or ''}`"
            )

            # Diffs détaillés
            diffs = h.get("diffs_detailles", [])
            if diffs:
                st.markdown("**🔄 Changements par rapport à la version précédente :**")
                diff_df = pd.DataFrame([
                    {
                        "Champ": FIELD_BY_KEY.get(d["champ"],
                                                  (d["champ"],))[0],
                        "Avant": str(d["avant"]),
                        "Après": str(d["apres"]),
                    } for d in diffs
                ])
                st.dataframe(diff_df, use_container_width=True,
                              hide_index=True)
            elif h["numero_version"] == 1:
                st.caption("🆕 Version initiale — aucune comparaison possible.")
            else:
                st.caption("✅ Aucun changement détecté (réimport identique).")


# -----------------------------------------------------------------------------
# SOUS-ONGLET 3 : ANOMALIES RÉCAPITULATIVES
# -----------------------------------------------------------------------------

def _adv_anomalies_tab(db_path: Path):
    st.markdown("### Tableau récapitulatif des anomalies")
    st.caption("Vue d'ensemble de toutes les incohérences détectées dans la base. "
               "Chaque catégorie peut être explorée en détail.")

    conn = sqlite3.connect(str(db_path))
    try:
        anomalies = detect_anomalies(conn)
    except Exception as e:
        st.error(f"Erreur : {e}")
        return
    finally:
        conn.close()

    # Tableau récapitulatif principal
    summary = anomalies_summary(anomalies)
    total_anomalies = sum(s["Nombre"] for s in summary)

    # KPI global
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🔴 Erreurs",
              sum(s["Nombre"] for s in summary if s["Gravité"] == "🔴"))
    c2.metric("🟠 Attentions",
              sum(s["Nombre"] for s in summary if s["Gravité"] == "🟠"))
    c3.metric("🔵 Informations",
              sum(s["Nombre"] for s in summary if s["Gravité"] == "🔵"))
    c4.metric("📊 Total anomalies", total_anomalies)

    st.markdown("---")

    # Tableau
    st.markdown("#### Synthèse par catégorie")
    df_summary = pd.DataFrame([{k: v for k, v in s.items() if not k.startswith("_")}
                                for s in summary])
    st.dataframe(df_summary, use_container_width=True, hide_index=True)

    # Export global
    col1, col2 = st.columns([1, 3])
    with col1:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        tmp = Path(tempfile.mktemp(suffix=".xlsx"))
        export_anomalies_to_excel(anomalies, tmp)
        xlsx_bytes = tmp.read_bytes()
        tmp.unlink()
        st.download_button(
            "📄 Export complet (Excel)",
            data=xlsx_bytes,
            file_name=f"anomalies_{stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="Un classeur avec une feuille par catégorie d'anomalies",
        )

    st.markdown("---")

    # Détails par catégorie
    st.markdown("#### 🔍 Explorer chaque catégorie")

    for key, data in anomalies.items():
        items = data["items"]
        icon = {"error": "🔴", "warning": "🟠", "info": "🔵"}.get(
            data["severity"], "⚪")
        label = f"{icon} **{data['title']}** — {len(items)} élément(s)"

        with st.expander(label, expanded=False):
            st.caption(data["description"])

            if not items:
                st.success("✅ Aucune anomalie dans cette catégorie.")
                continue

            # Afficher en DataFrame
            df = pd.DataFrame(items)
            if "updated_at" in df.columns:
                df["updated_at"] = df["updated_at"].apply(fmt_date)

            # Limiter à 200 lignes pour l'affichage
            if len(df) > 200:
                st.info(f"Affichage des 200 premières lignes sur {len(df)}")
                df = df.head(200)

            st.dataframe(df, use_container_width=True, hide_index=True,
                          height=300)

            # Export CSV de cette catégorie
            csv_data = df.to_csv(index=False, sep=";",
                                  encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(
                f"📋 Exporter CSV ({data['title']})",
                data=csv_data,
                file_name=f"anomalies_{key}_{datetime.now():%Y%m%d_%H%M%S}.csv",
                mime="text/csv",
                key=f"dl_anom_{key}",
            )


# =============================================================================
# APPLICATION PRINCIPALE
# =============================================================================

def main():
    render_header()
    render_sidebar()

    # Navigation par tabs (5 onglets maintenant)
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📥 Traitement",
        "🔍 Recherche",
        "🔎 Recherche avancée",
        "📈 Tableau de bord",
        "📊 État de la base",
    ])

    with tab1:
        page_traitement()
    with tab2:
        page_recherche()
    with tab3:
        page_recherche_avancee()
    with tab4:
        page_dashboard()
    with tab5:
        page_infos()


if __name__ == "__main__":
    main()
