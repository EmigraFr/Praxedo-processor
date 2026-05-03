#!/usr/bin/env python3
"""
Praxedo Processor — Version Web v3.1 (Corrigée)
============================================
Corrections : Gestion de la session et édition manuelle des POI.
"""

from __future__ import annotations
from typing import Any, Dict, List, Optional, Tuple

import io
import os
import sys
import tempfile
import sqlite3
import re
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
# CONFIGURATION ET STYLE
# =============================================================================

st.set_page_config(page_title="Praxedo Processor", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1f4e79 0%, #305496 100%);
        color: white; padding: 24px; border-radius: 10px; margin-bottom: 18px;
    }
    .stDownloadButton>button { background: #305496; color: white; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# =============================================================================
# ÉTAT DE SESSION (Mémoire de l'application)
# =============================================================================

DEFAULT_DB_DIR = Path.home() / "Praxedo_Resultats"
DEFAULT_DB_DIR.mkdir(parents=True, exist_ok=True)
DEFAULT_DB_PATH = DEFAULT_DB_DIR / "praxedo.db"

# Initialisation des variables en mémoire
if "db_path" not in st.session_state: st.session_state.db_path = str(DEFAULT_DB_PATH)
if "output_dir" not in st.session_state: st.session_state.output_dir = str(DEFAULT_DB_DIR)
if "search_results" not in st.session_state: st.session_state.search_results = None
if "detail_id" not in st.session_state: st.session_state.detail_id = None
if "adv_criteria" not in st.session_state:
    st.session_state.adv_criteria = [{"field": "tipologie", "operator": "égal à", "value": ""}]
# Variable pour stocker le dernier fichier traité pour édition
if "df_a_corriger" not in st.session_state: st.session_state.df_a_corriger = None

# =============================================================================
# FONCTIONS UTILITAIRES
# =============================================================================

def fmt_date(s: Optional[str]) -> str:
    if not s: return "—"
    try: return datetime.fromisoformat(s).strftime("%d/%m/%Y %H:%M")
    except: return str(s)

def fmt_size(n: int) -> str:
    for unit in ("o", "Ko", "Mo", "Go"):
        if n < 1024: return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} To"

def get_db_summary():
    dbp = Path(st.session_state.db_path)
    if not dbp.exists(): return None
    try:
        db = PraxedoDB(dbp)
        s = db.summary()
        db.close()
        return s
    except: return None

# =============================================================================
# PAGE 1 : TRAITEMENT ET ÉDITION
# =============================================================================

def page_traitement():
    st.markdown("## 📥 Traitement et Correction")
    
    uploaded_file = st.file_uploader("Déposez votre export Praxedo (.xlsx / .xlsm)", type=["xlsx", "xlsm"])

    if uploaded_file:
        col1, col2 = st.columns(2)
        with col1:
            use_db = st.checkbox("💾 Ajouter à la base de données", value=True)
        
        if st.button("🚀 1. Lancer le traitement automatique", type="primary"):
            out_dir = Path(st.session_state.output_dir)
            out_dir.mkdir(parents=True, exist_ok=True)
            
            # Sauvegarde temporaire pour traitement
            tmp_path = out_dir / f"_tmp_{uploaded_file.name}"
            tmp_path.write_bytes(uploaded_file.getvalue())
            
            db = PraxedoDB(st.session_state.db_path) if use_db else None
            out_xlsx = out_dir / f"{Path(uploaded_file.name).stem}_traite.xlsx"
            
            try:
                # Traitement core
                stats = process_workbook(tmp_path, out_xlsx, db=db)
                st.success(f"✅ Traitement terminé : {stats['rows']} lignes analysées.")
                
                # Chargement du résultat pour permettre l'édition manuelle
                st.session_state.df_a_corriger = pd.read_excel(out_xlsx)
            except Exception as e:
                st.error(f"Erreur : {e}")
            finally:
                if tmp_path.exists(): tmp_path.unlink()
                if db: db.close()

    # SECTION ÉDITION MANUELLE (Apparaît seulement après traitement)
    if st.session_state.df_a_corriger is not None:
        st.divider()
        st.markdown("### 📝 2. Correction manuelle (POI / Tipologie)")
        st.info("Modifiez directement les cases vides ci-dessous. Les changements seront inclus dans votre export final.")
        
        # Éditeur interactif
        df_modifie = st.data_editor(
            st.session_state.df_a_corriger,
            use_container_width=True,
            hide_index=True,
            column_config={
                "POI": st.column_config.TextColumn("POI", help="Saisissez le POI si manquant"),
                "Tipologie": st.column_config.SelectboxColumn("Tipologie", options=["STD", "NSTD"])
            }
        )
        
        # Bouton de téléchargement final
        col_dl1, col_dl2 = st.columns([1, 4])
        with col_dl1:
            # Préparation du fichier Excel en mémoire
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_modifie.to_excel(writer, index=False, sheet_name='Données Corrigées')
            
            st.download_button(
                label="📥 Télécharger l'Excel Corrigé",
                data=output.getvalue(),
                file_name="export_praxedo_final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# =============================================================================
# AUTRES PAGES (Conservées pour la structure)
# =============================================================================

def page_recherche():
    # ... (Le reste de votre code de recherche reste identique)
    st.markdown("## 🔍 Recherche dans la base")
    # [Votre code précédent ici...]
    st.info("Utilisez l'onglet 'Traitement' pour importer des données.")

# =============================================================================
# LANCEMENT
# =============================================================================

def main():
    st.markdown('<div class="main-header"><h1>📊 Praxedo Processor</h1></div>', unsafe_allow_html=True)
    
    # Sidebar
    st.sidebar.title("Paramètres")
    st.session_state.db_path = st.sidebar.text_input("Base SQL", value=st.session_state.db_path)
    
    summary = get_db_summary()
    if summary:
        st.sidebar.metric("Interventions", f"{summary['nb_interventions']}")
    
    # Menu principal
    tab1, tab2, tab3 = st.tabs(["📥 Traitement & Correction", "🔍 Recherche", "📈 Dashboard"])
    
    with tab1:
        page_traitement()
    with tab2:
        page_recherche()
    with tab3:
        st.write("Le dashboard est disponible après traitement des données.")

if __name__ == "__main__":
    main()
