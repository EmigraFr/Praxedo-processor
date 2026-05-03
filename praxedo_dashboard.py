#!/usr/bin/env python3
"""
Praxedo Dashboard — Calculs KPI + génération graphiques + export PDF
=====================================================================
Module réutilisable par l'app Streamlit (Phase 3).

Fournit :
  - compute_kpi(filters) : KPI de synthèse
  - timeseries(period)   : évolution temporelle
  - repartition_tipo()   : donut STD/NSTD
  - stats_par_agence()   : barres empilées
  - top_techniciens(n)   : classement
  - top_caff(n)          : classement CAFF
  - top_clients(n)       : classement clients
  - nd_recurrents(n)     : ND avec doublons
  - alertes()            : anomalies détectées
  - generate_pdf_report() : PDF reporting prêt à diffuser
"""

from __future__ import annotations
from typing import Any, Dict, List, Optional, Tuple

import io
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path


# =============================================================================
# CALCULS — KPI ET AGRÉGATIONS
# =============================================================================

PERIODS = {
    "7 derniers jours":    7,
    "30 derniers jours":   30,
    "90 derniers jours":   90,
    "6 derniers mois":     182,
    "12 derniers mois":    365,
    "Toute la base":       None,
}


def _period_clause(days: Optional[int], col: str = "updated_at") -> str:
    """Construit une clause WHERE de période (retourne '' si None)."""
    if days is None:
        return ""
    cutoff = (datetime.now() - timedelta(days=days)).isoformat(timespec="seconds")
    return f" AND {col} >= '{cutoff}' "


def compute_kpi(conn: sqlite3.Connection,
                period_days: Optional[int] = None,
                agence: Optional[str] = None) -> Dict[str, Any]:
    """Calcule les 8 KPI principaux du dashboard."""

    where = "WHERE 1=1"
    if period_days is not None:
        cutoff = (datetime.now() - timedelta(days=period_days)).isoformat(timespec="seconds")
        where += f" AND updated_at >= '{cutoff}'"
    if agence:
        where += f" AND agence = '{agence.replace(chr(39), chr(39)*2)}'"

    def s(q: str) -> int:
        r = conn.execute(q).fetchone()
        return (r[0] if r and r[0] is not None else 0)

    total = s(f"SELECT COUNT(*) FROM interventions {where}")
    std   = s(f"SELECT SUM(standard)     FROM interventions {where}")
    nstd  = s(f"SELECT SUM(non_standard) FROM interventions {where}")
    arc   = s(f"SELECT SUM(arc)          FROM interventions {where}")
    no_poi = s(f"SELECT COUNT(*) FROM interventions {where} AND (poi IS NULL OR poi = '')")
    doublons = s(f"SELECT COUNT(*) FROM interventions {where} AND nb_versions > 1")

    # Évolution : comparer à la période précédente
    trend_std = trend_arc = trend_total = None
    if period_days:
        cutoff_prev = (datetime.now() - timedelta(days=period_days * 2)).isoformat(timespec="seconds")
        cutoff_start = (datetime.now() - timedelta(days=period_days)).isoformat(timespec="seconds")
        prev_total = s(
            f"SELECT COUNT(*) FROM interventions "
            f"WHERE updated_at >= '{cutoff_prev}' AND updated_at < '{cutoff_start}'"
        )
        prev_std = s(
            f"SELECT SUM(standard) FROM interventions "
            f"WHERE updated_at >= '{cutoff_prev}' AND updated_at < '{cutoff_start}'"
        )
        prev_arc = s(
            f"SELECT SUM(arc) FROM interventions "
            f"WHERE updated_at >= '{cutoff_prev}' AND updated_at < '{cutoff_start}'"
        )
        trend_total = total - prev_total
        trend_std = std - prev_std
        trend_arc = arc - prev_arc

    taux_std = round(100 * std / total, 1) if total > 0 else 0
    taux_nstd = round(100 * nstd / total, 1) if total > 0 else 0

    return {
        "total":         total,
        "std":           std,
        "nstd":          nstd,
        "arc":           arc,
        "no_poi":        no_poi,
        "doublons":      doublons,
        "taux_std":      taux_std,
        "taux_nstd":     taux_nstd,
        "trend_total":   trend_total,
        "trend_std":     trend_std,
        "trend_arc":     trend_arc,
        "period_days":   period_days,
        "agence":        agence,
    }


def timeseries(conn: sqlite3.Connection,
               period_days: Optional[int] = 90,
               granularity: str = "week") -> List[Tuple[str, int, int, int, int]]:
    """Série temporelle : (période, total, std, nstd, arc)."""
    where = ""
    if period_days:
        cutoff = (datetime.now() - timedelta(days=period_days)).isoformat(timespec="seconds")
        where = f"WHERE updated_at >= '{cutoff}'"

    fmt = {"day": "%Y-%m-%d", "week": "%Y-W%W", "month": "%Y-%m"}[granularity]

    rows = conn.execute(
        f"SELECT strftime('{fmt}', updated_at) AS p, "
        f"       COUNT(*), "
        f"       COALESCE(SUM(standard), 0), "
        f"       COALESCE(SUM(non_standard), 0), "
        f"       COALESCE(SUM(arc), 0) "
        f"FROM interventions {where} "
        f"GROUP BY p ORDER BY p"
    ).fetchall()
    return [tuple(r) for r in rows]


def repartition_tipo(conn: sqlite3.Connection,
                     period_days: Optional[int] = None) -> Dict[str, int]:
    where = ""
    if period_days:
        cutoff = (datetime.now() - timedelta(days=period_days)).isoformat(timespec="seconds")
        where = f"WHERE updated_at >= '{cutoff}'"

    std = conn.execute(
        f"SELECT SUM(standard) FROM interventions {where}"
    ).fetchone()[0] or 0
    nstd = conn.execute(
        f"SELECT SUM(non_standard) FROM interventions {where}"
    ).fetchone()[0] or 0

    total = conn.execute(
        f"SELECT COUNT(*) FROM interventions {where}"
    ).fetchone()[0] or 0
    inconnu = max(0, total - std - nstd)

    return {"STD": std, "NSTD": nstd, "Inconnu": inconnu}


def stats_par_agence(conn: sqlite3.Connection,
                     period_days: Optional[int] = None) -> List[Dict]:
    where = ""
    if period_days:
        cutoff = (datetime.now() - timedelta(days=period_days)).isoformat(timespec="seconds")
        where = f"AND updated_at >= '{cutoff}'"

    rows = conn.execute(
        f"SELECT COALESCE(agence, 'Non renseignée') AS agence, "
        f"       COUNT(*) AS total, "
        f"       COALESCE(SUM(standard), 0) AS std, "
        f"       COALESCE(SUM(non_standard), 0) AS nstd, "
        f"       COALESCE(SUM(arc), 0) AS arc "
        f"FROM interventions "
        f"WHERE 1=1 {where} "
        f"GROUP BY agence ORDER BY total DESC LIMIT 10"
    ).fetchall()
    return [{"agence": r[0], "total": r[1], "std": r[2],
             "nstd": r[3], "arc": r[4]} for r in rows]


def top_techniciens(conn: sqlite3.Connection, n: int = 10,
                    period_days: Optional[int] = None) -> List[Dict]:
    where = ""
    if period_days:
        cutoff = (datetime.now() - timedelta(days=period_days)).isoformat(timespec="seconds")
        where = f"AND updated_at >= '{cutoff}'"

    rows = conn.execute(
        f"SELECT COALESCE(technicien_nom, '') || ' ' || COALESCE(technicien_prenom, '') AS tech, "
        f"       COUNT(*) AS n "
        f"FROM interventions "
        f"WHERE technicien_nom IS NOT NULL AND technicien_nom != '' {where} "
        f"GROUP BY tech ORDER BY n DESC LIMIT ?", (n,)
    ).fetchall()
    return [{"nom": r[0].strip(), "nb": r[1]} for r in rows]


def top_caff(conn: sqlite3.Connection, n: int = 10,
             period_days: Optional[int] = None) -> List[Dict]:
    where = ""
    if period_days:
        cutoff = (datetime.now() - timedelta(days=period_days)).isoformat(timespec="seconds")
        where = f"AND updated_at >= '{cutoff}'"

    rows = conn.execute(
        f"SELECT caff_nom, COUNT(*) AS n "
        f"FROM interventions "
        f"WHERE caff_nom IS NOT NULL AND caff_nom != '' {where} "
        f"GROUP BY caff_nom ORDER BY n DESC LIMIT ?", (n,)
    ).fetchall()
    return [{"nom": r[0], "nb": r[1]} for r in rows]


def top_clients(conn: sqlite3.Connection, n: int = 10,
                period_days: Optional[int] = None) -> List[Dict]:
    where = ""
    if period_days:
        cutoff = (datetime.now() - timedelta(days=period_days)).isoformat(timespec="seconds")
        where = f"AND updated_at >= '{cutoff}'"

    rows = conn.execute(
        f"SELECT client, COUNT(*) AS n "
        f"FROM interventions "
        f"WHERE client IS NOT NULL AND client != '' {where} "
        f"GROUP BY client ORDER BY n DESC LIMIT ?", (n,)
    ).fetchall()
    return [{"nom": r[0], "nb": r[1]} for r in rows]


def nd_recurrents(conn: sqlite3.Connection, n: int = 10) -> List[Dict]:
    """ND qui reviennent souvent (plusieurs interventions différentes)."""
    rows = conn.execute(
        "SELECT nd, COUNT(*) AS n "
        "FROM interventions "
        "WHERE nd IS NOT NULL AND nd != '' "
        "GROUP BY nd HAVING COUNT(*) > 1 "
        "ORDER BY n DESC LIMIT ?", (n,)
    ).fetchall()
    return [{"nd": r[0], "nb": r[1]} for r in rows]


def alertes(conn: sqlite3.Connection) -> Dict[str, List]:
    """Détecte les anomalies."""
    now = datetime.now()
    week_ago = (now - timedelta(days=7)).isoformat(timespec="seconds")

    # Sans POI récents
    no_poi_recent = conn.execute(
        "SELECT refint, nd, updated_at FROM interventions "
        "WHERE (poi IS NULL OR poi = '') AND updated_at >= ? "
        "ORDER BY updated_at DESC LIMIT 20",
        (week_ago,)
    ).fetchall()

    # Refint avec plusieurs versions
    doublons = conn.execute(
        "SELECT refint, nb_versions, updated_at FROM interventions "
        "WHERE nb_versions > 1 ORDER BY nb_versions DESC LIMIT 20"
    ).fetchall()

    # Emails RPE non orange (potentiellement suspects)
    emails_suspects = conn.execute(
        "SELECT refint, email_rpe FROM interventions "
        "WHERE email_rpe IS NOT NULL AND email_rpe != '' "
        "AND email_rpe NOT LIKE '%orange%' LIMIT 20"
    ).fetchall()

    # Téléphone manquant alors que CAFF présent
    caff_sans_tel = conn.execute(
        "SELECT refint, caff_nom FROM interventions "
        "WHERE caff_nom IS NOT NULL AND caff_nom != '' "
        "AND (caff_tel IS NULL OR caff_tel = '') LIMIT 20"
    ).fetchall()

    return {
        "no_poi_recent": [dict(refint=r[0], nd=r[1], date=r[2]) for r in no_poi_recent],
        "doublons":      [dict(refint=r[0], versions=r[1], date=r[2]) for r in doublons],
        "emails_suspects": [dict(refint=r[0], email=r[1]) for r in emails_suspects],
        "caff_sans_tel": [dict(refint=r[0], caff=r[1]) for r in caff_sans_tel],
    }


def get_distinct_agences(conn: sqlite3.Connection) -> List[str]:
    rows = conn.execute(
        "SELECT DISTINCT agence FROM interventions "
        "WHERE agence IS NOT NULL AND agence != '' "
        "ORDER BY agence"
    ).fetchall()
    return [r[0] for r in rows]


# =============================================================================
# GRAPHIQUES PLOTLY
# =============================================================================

def fig_timeseries(data: List[Tuple], title: str = "Évolution dans le temps"):
    import plotly.graph_objects as go
    fig = go.Figure()
    if not data:
        fig.add_annotation(text="Aucune donnée", showarrow=False,
                           xref="paper", yref="paper", x=0.5, y=0.5)
    else:
        periodes = [r[0] for r in data]
        totaux = [r[1] for r in data]
        stds = [r[2] for r in data]
        nstds = [r[3] for r in data]

        fig.add_trace(go.Scatter(x=periodes, y=totaux, name="Total",
                                 mode="lines+markers", line=dict(color="#305496", width=3)))
        fig.add_trace(go.Scatter(x=periodes, y=stds, name="STD",
                                 mode="lines+markers", line=dict(color="#2e7d32", width=2)))
        fig.add_trace(go.Scatter(x=periodes, y=nstds, name="NSTD",
                                 mode="lines+markers", line=dict(color="#ef6c00", width=2)))

    fig.update_layout(
        title=title,
        xaxis_title="Période",
        yaxis_title="Nombre d'interventions",
        height=380,
        hovermode="x unified",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=40),
    )
    return fig


def fig_donut_tipo(repartition: Dict[str, int], title: str = "Répartition STD / NSTD"):
    import plotly.graph_objects as go
    labels = list(repartition.keys())
    values = list(repartition.values())
    colors = {"STD": "#2e7d32", "NSTD": "#ef6c00", "Inconnu": "#9e9e9e"}

    fig = go.Figure(data=[go.Pie(
        labels=labels, values=values, hole=0.55,
        marker=dict(colors=[colors.get(l, "#777") for l in labels]),
        textinfo="label+percent",
    )])
    fig.update_layout(
        title=title, height=380, template="plotly_white",
        margin=dict(l=20, r=20, t=50, b=20),
    )
    return fig


def fig_bar_agence(agences: List[Dict], title: str = "Volume par agence"):
    import plotly.graph_objects as go
    if not agences:
        fig = go.Figure()
        fig.add_annotation(text="Aucune donnée", showarrow=False,
                           xref="paper", yref="paper", x=0.5, y=0.5)
        fig.update_layout(title=title, height=380)
        return fig

    noms = [a["agence"][:30] for a in agences]
    stds = [a["std"] for a in agences]
    nstds = [a["nstd"] for a in agences]
    arcs = [a["arc"] for a in agences]

    fig = go.Figure()
    fig.add_trace(go.Bar(name="STD", x=noms, y=stds, marker_color="#2e7d32"))
    fig.add_trace(go.Bar(name="NSTD", x=noms, y=nstds, marker_color="#ef6c00"))
    fig.add_trace(go.Bar(name="ARC", x=noms, y=arcs, marker_color="#c62828"))
    fig.update_layout(
        title=title, barmode="stack", height=380,
        xaxis_tickangle=-30, template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=120),
    )
    return fig


def fig_top_bar(items: List[Dict], key_name="nom", key_val="nb",
                title: str = "Top 10", color: str = "#305496"):
    import plotly.graph_objects as go
    if not items:
        fig = go.Figure()
        fig.add_annotation(text="Aucune donnée", showarrow=False,
                           xref="paper", yref="paper", x=0.5, y=0.5)
        fig.update_layout(title=title, height=380)
        return fig

    noms = [str(i[key_name])[:40] for i in reversed(items)]
    vals = [i[key_val] for i in reversed(items)]
    fig = go.Figure(go.Bar(x=vals, y=noms, orientation="h",
                           marker_color=color,
                           text=vals, textposition="outside"))
    fig.update_layout(
        title=title, height=380, template="plotly_white",
        margin=dict(l=150, r=40, t=50, b=40),
        xaxis_title="Nombre",
    )
    return fig


# =============================================================================
# EXPORT PDF REPORTING
# =============================================================================

def generate_pdf_report(conn: sqlite3.Connection,
                        output_path: Path,
                        period_label: str = "Toute la base",
                        period_days: Optional[int] = None,
                        agence: Optional[str] = None) -> Path:
    """Génère un rapport PDF complet (8-10 pages A4)."""
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import cm
    from reportlab.lib.colors import HexColor, white
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                     Table, TableStyle, PageBreak, Image)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT

    # Calculs
    kpi = compute_kpi(conn, period_days=period_days, agence=agence)
    ts = timeseries(conn, period_days=period_days or 90, granularity="week")
    tipo = repartition_tipo(conn, period_days=period_days)
    agences = stats_par_agence(conn, period_days=period_days)
    techs = top_techniciens(conn, n=10, period_days=period_days)
    caffs = top_caff(conn, n=10, period_days=period_days)
    alerts = alertes(conn)

    # Convertir figures Plotly en PNG via kaleido
    try:
        import plotly.io as pio
        img_timeseries = pio.to_image(fig_timeseries(ts), format="png",
                                      width=1000, height=400, scale=2)
        img_donut = pio.to_image(fig_donut_tipo(tipo), format="png",
                                 width=600, height=500, scale=2)
        img_agence = pio.to_image(fig_bar_agence(agences), format="png",
                                  width=1000, height=500, scale=2)
        img_top_tech = pio.to_image(
            fig_top_bar(techs, title="Top 10 techniciens", color="#305496"),
            format="png", width=1000, height=500, scale=2)
        img_top_caff = pio.to_image(
            fig_top_bar(caffs, title="Top 10 CAFF", color="#2e7d32"),
            format="png", width=1000, height=500, scale=2)
    except Exception:
        # Fallback : si kaleido absent, pas de graphiques
        img_timeseries = img_donut = img_agence = None
        img_top_tech = img_top_caff = None

    # Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleStyle", parent=styles["Title"],
        fontSize=24, textColor=HexColor("#1f4e79"),
        alignment=TA_CENTER, spaceAfter=20,
    )
    h2 = ParagraphStyle(
        "H2", parent=styles["Heading2"],
        fontSize=14, textColor=HexColor("#305496"),
        spaceBefore=12, spaceAfter=8,
    )
    normal = styles["Normal"]
    center = ParagraphStyle("Center", parent=normal, alignment=TA_CENTER)
    small = ParagraphStyle("Small", parent=normal, fontSize=8,
                           textColor=HexColor("#666"))

    doc = SimpleDocTemplate(
        str(output_path), pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm,
    )
    story = []

    # ========== PAGE DE GARDE ==========
    story.append(Spacer(1, 4*cm))
    story.append(Paragraph("📊 Praxedo Processor", title_style))
    story.append(Paragraph("Rapport de tableau de bord", center))
    story.append(Spacer(1, 2*cm))

    meta_data = [
        ["Période couverte :", period_label],
        ["Agence :", agence or "Toutes"],
        ["Date d'édition :", datetime.now().strftime("%d/%m/%Y %H:%M")],
        ["Interventions totales :", f"{kpi['total']:,}".replace(",", " ")],
    ]
    meta_table = Table(meta_data, colWidths=[6*cm, 8*cm])
    meta_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 11),
        ("TEXTCOLOR", (0, 0), (0, -1), HexColor("#305496")),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LINEBELOW", (0, 0), (-1, -1), 0.5, HexColor("#e0e4e9")),
    ]))
    story.append(meta_table)
    story.append(PageBreak())

    # ========== PAGE KPI ==========
    story.append(Paragraph("Indicateurs clés", h2))
    story.append(Spacer(1, 0.3*cm))

    kpi_rows = [
        ["Total interventions",   f"{kpi['total']:,}".replace(",", " "),
         "Standards (STD)",       f"{kpi['std']}"],
        ["Taux standardisation",  f"{kpi['taux_std']} %",
         "Non Standards",         f"{kpi['nstd']}"],
        ["ARC détectés",          f"{kpi['arc']}",
         "POI manquants",         f"{kpi['no_poi']}"],
        ["Doublons (>= 2 versions)", f"{kpi['doublons']}",
         "Taux NSTD",             f"{kpi['taux_nstd']} %"],
    ]
    kpi_table = Table(kpi_rows, colWidths=[4.5*cm, 3*cm, 4.5*cm, 3*cm])
    kpi_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 11),
        ("BACKGROUND", (0, 0), (0, -1), HexColor("#eaf0f7")),
        ("BACKGROUND", (2, 0), (2, -1), HexColor("#eaf0f7")),
        ("TEXTCOLOR", (1, 0), (1, -1), HexColor("#1f4e79")),
        ("TEXTCOLOR", (3, 0), (3, -1), HexColor("#1f4e79")),
        ("FONTNAME", (1, 0), (1, -1), "Helvetica-Bold"),
        ("FONTNAME", (3, 0), (3, -1), "Helvetica-Bold"),
        ("FONTSIZE", (1, 0), (1, -1), 14),
        ("FONTSIZE", (3, 0), (3, -1), 14),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOX", (0, 0), (-1, -1), 0.5, HexColor("#c0c0c0")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, HexColor("#c0c0c0")),
    ]))
    story.append(kpi_table)
    story.append(PageBreak())

    # ========== GRAPHIQUES ==========
    def add_image_page(img_bytes, title):
        story.append(Paragraph(title, h2))
        story.append(Spacer(1, 0.3*cm))
        if img_bytes:
            img = Image(io.BytesIO(img_bytes), width=17*cm, height=8*cm,
                        kind="proportional")
            story.append(img)
        else:
            story.append(Paragraph(
                "(Graphique indisponible - module kaleido absent)", small))
        story.append(PageBreak())

    add_image_page(img_timeseries, "Évolution temporelle")
    add_image_page(img_donut, "Répartition STD / NSTD")
    add_image_page(img_agence, "Volume par agence")
    if techs:
        add_image_page(img_top_tech, "Top 10 techniciens")
    if caffs:
        add_image_page(img_top_caff, "Top 10 CAFF")

    # ========== TABLEAU SYNTHÉTIQUE ==========
    story.append(Paragraph("Synthèse par agence", h2))
    if agences:
        data = [["Agence", "Total", "STD", "NSTD", "ARC"]]
        for a in agences:
            data.append([a["agence"][:35], a["total"], a["std"],
                         a["nstd"], a["arc"]])
        t = Table(data, colWidths=[7*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), HexColor("#305496")),
            ("TEXTCOLOR", (0, 0), (-1, 0), white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("GRID", (0, 0), (-1, -1), 0.5, HexColor("#c0c0c0")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [white, HexColor("#f8fafc")]),
        ]))
        story.append(t)

    # ========== ALERTES ==========
    story.append(PageBreak())
    story.append(Paragraph("Alertes et anomalies", h2))
    story.append(Spacer(1, 0.3*cm))

    alert_info = [
        ("POI manquants (7 derniers jours)", len(alerts["no_poi_recent"])),
        ("Interventions avec plusieurs versions", len(alerts["doublons"])),
        ("Emails RPE non-Orange", len(alerts["emails_suspects"])),
        ("CAFF sans téléphone", len(alerts["caff_sans_tel"])),
    ]
    at = Table([["Type d'alerte", "Nombre"]] + [[l, n] for l, n in alert_info],
               colWidths=[11*cm, 3*cm])
    at.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), HexColor("#c62828")),
        ("TEXTCOLOR", (0, 0), (-1, 0), white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("GRID", (0, 0), (-1, -1), 0.5, HexColor("#c0c0c0")),
        ("ALIGN", (1, 0), (1, -1), "CENTER"),
    ]))
    story.append(at)

    # Pied de page
    def footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(HexColor("#777777"))
        canvas.drawString(2*cm, 1*cm,
                          f"Praxedo Processor — Rapport généré le "
                          f"{datetime.now().strftime('%d/%m/%Y %H:%M')}")
        canvas.drawRightString(A4[0] - 2*cm, 1*cm,
                               f"Page {doc_.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return output_path
