from __future__ import annotations

from hashlib import sha256
from html import escape
from pathlib import Path
from time import perf_counter

import streamlit as st

from processor import (
    GeneratedPackage,
    InvalidWorkbookError,
    LuctivError,
    process_uploaded_package,
)
from well_visualization import build_well_figure, calculate_impact_metrics


APP_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = APP_DIR / "assets" / "plantilla_datos_terminados.xlsx"
SUPPORTED_SUFFIXES = {".xlsm", ".xlsx", ".zip"}

st.set_page_config(
    page_title="LUCTIV",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    :root {
        --luctiv-ink: #202020;
        --luctiv-red: #CC0000;
        --luctiv-red-dark: #A80000;
        --luctiv-soft: #F7F7F7;
        --luctiv-border: #E0E0E0;
        --luctiv-muted: #5B5B5B;
    }
    .stApp { background: linear-gradient(180deg, #F7F7F7 0%, #FFFFFF 36%); }
    .block-container { max-width: 1180px; padding-top: 2.4rem; padding-bottom: 3rem; }
    .luctiv-hero {
        border: 1px solid var(--luctiv-border);
        border-left: 6px solid var(--luctiv-red);
        border-radius: 22px;
        padding: 2rem 2.1rem;
        background: rgba(255,255,255,0.94);
        box-shadow: 0 18px 48px rgba(32, 32, 32, 0.06);
        margin-bottom: 1.4rem;
    }
    .luctiv-kicker {
        color: var(--luctiv-red);
        font-size: 0.82rem;
        letter-spacing: 0.18em;
        font-weight: 800;
        margin-bottom: 0.35rem;
    }
    .luctiv-title {
        color: var(--luctiv-ink);
        font-size: clamp(2.6rem, 7vw, 4.6rem);
        letter-spacing: -0.06em;
        line-height: 0.95;
        font-weight: 850;
        margin: 0;
    }
    .luctiv-subtitle {
        color: var(--luctiv-muted);
        max-width: 680px;
        font-size: 1.04rem;
        margin: 1rem 0 0 0;
        line-height: 1.6;
    }
    div[data-testid="stFileUploader"] {
        border: 1px dashed #C75C5C;
        border-radius: 18px;
        padding: 0.45rem 0.75rem 0.15rem;
        background: var(--luctiv-soft);
    }
    div.stButton > button, div.stDownloadButton > button {
        border-radius: 12px;
        min-height: 3rem;
        font-weight: 750;
        border: none;
    }
    div.stButton > button[kind="primary"], div.stDownloadButton > button {
        background: var(--luctiv-red);
        color: white;
    }
    div.stButton > button[kind="primary"]:hover,
    div.stDownloadButton > button:hover {
        background: var(--luctiv-red-dark);
        color: white;
    }
    .luctiv-flow {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 0.7rem;
        margin: 1rem 0 0.45rem;
    }
    .luctiv-flow-step {
        border: 1px solid var(--luctiv-border);
        border-radius: 14px;
        background: var(--luctiv-soft);
        color: var(--luctiv-ink);
        padding: 0.85rem 0.9rem;
        font-size: 0.88rem;
        font-weight: 750;
        text-align: center;
    }
    .luctiv-flow-step span {
        display: inline-block;
        color: var(--luctiv-red);
        margin-right: 0.3rem;
    }
    .luctiv-note {
        color: var(--luctiv-muted);
        font-size: 0.86rem;
        border-top: 1px solid var(--luctiv-border);
        padding-top: 1rem;
        margin-top: 1.5rem;
    }
    .luctiv-result-heading {
        min-height: 3.35rem;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .luctiv-result-heading h1 {
        color: var(--luctiv-ink);
        font-size: 1.75rem;
        line-height: 1;
        letter-spacing: -0.03em;
        margin: 0;
    }
    .luctiv-result-heading p {
        color: var(--luctiv-muted);
        font-size: 0.78rem;
        margin: 0.28rem 0 0;
    }
    .luctiv-kpis {
        display: grid;
        grid-template-columns: repeat(6, minmax(0, 1fr));
        gap: 0.42rem;
    }
    .luctiv-kpi {
        min-width: 0;
        border-left: 3px solid var(--luctiv-red);
        background: var(--luctiv-soft);
        padding: 0.42rem 0.62rem;
    }
    .luctiv-kpi-label {
        color: var(--luctiv-muted);
        font-size: 0.66rem;
        letter-spacing: 0.04em;
        text-transform: uppercase;
        white-space: nowrap;
    }
    .luctiv-kpi-value {
        color: var(--luctiv-ink);
        font-size: 1.05rem;
        font-weight: 800;
        line-height: 1.15;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .luctiv-summary {
        height: 485px;
        padding: 0.75rem 0.2rem 0 0.85rem;
        border-left: 1px solid var(--luctiv-border);
        color: var(--luctiv-ink);
    }
    .luctiv-summary h3 {
        font-size: 1rem;
        margin: 0 0 0.65rem;
    }
    .luctiv-summary-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 0.45rem;
    }
    .luctiv-summary-item {
        background: var(--luctiv-soft);
        padding: 0.55rem 0.62rem;
        min-width: 0;
    }
    .luctiv-summary-label {
        color: var(--luctiv-muted);
        font-size: 0.65rem;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        white-space: nowrap;
    }
    .luctiv-summary-value {
        font-size: 1.12rem;
        font-weight: 800;
        line-height: 1.15;
    }
    .luctiv-status-list {
        margin-top: 0.7rem;
        border-top: 1px solid var(--luctiv-border);
        padding-top: 0.55rem;
    }
    .luctiv-status-row {
        display: flex;
        align-items: center;
        gap: 0.4rem;
        color: var(--luctiv-muted);
        font-size: 0.74rem;
        padding: 0.2rem 0;
    }
    .luctiv-status-dot {
        width: 0.48rem;
        height: 0.48rem;
        border-radius: 50%;
        background: var(--luctiv-red);
        flex: 0 0 auto;
    }
    .luctiv-source-note {
        color: var(--luctiv-muted);
        font-size: 0.7rem;
        line-height: 1.35;
        margin: 0.65rem 0 0;
    }
    .luctiv-chart-heading {
        color: var(--luctiv-ink);
        font-size: 1.05rem;
        font-weight: 750;
        line-height: 1.35;
        margin: 0.45rem 0 0;
    }
    @media (max-width: 720px) {
        .luctiv-flow { grid-template-columns: 1fr; }
        .luctiv-hero { padding: 1.5rem; }
        .luctiv-kpis { grid-template-columns: repeat(3, minmax(0, 1fr)); }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

def _render_upload_controls(compact: bool):
    controls = (
        st.popover("Archivo")
        if compact
        else st.container()
    )
    with controls:
        uploaded = st.file_uploader(
            "Excel del pozo o paquete del PAD",
            type=["xlsm", "xlsx", "zip"],
            help=(
                "Podés cargar un Excel individual o el ZIP recibido para procesar "
                "automáticamente todos los pozos compatibles."
            ),
            key="luctiv_workbook",
        )

        if compact:
            manual_time = st.number_input(
                "Tiempo manual de referencia (minutos)",
                min_value=1,
                max_value=480,
                value=45,
                step=5,
                help="Se usa solo para estimar el ahorro; no modifica el Excel.",
                key="luctiv_manual_minutes",
            )
        else:
            with st.expander("Referencia para estimar eficiencia", expanded=False):
                manual_time = st.number_input(
                    "Tiempo manual estimado para preparar y verificar un archivo (minutos)",
                    min_value=1,
                    max_value=480,
                    value=45,
                    step=5,
                    help="Este valor solo se usa para estimar el ahorro de tiempo; no modifica el Excel.",
                    key="luctiv_manual_minutes",
                )

        invalid = (
            uploaded is not None
            and Path(uploaded.name).suffix.lower() not in SUPPORTED_SUFFIXES
        )
        if invalid:
            st.error(
                "LUCTIV solo acepta Excel .xlsm/.xlsx o un paquete .zip.",
                icon="⛔",
            )
        if uploaded is not None:
            size_mb = uploaded.size / (1024 * 1024)
            st.caption(f"Archivo seleccionado: **{uploaded.name}** · {size_mb:.2f} MB")

        clicked = st.button(
            "Procesar archivo" if not compact else "Volver a procesar",
            type="primary",
            use_container_width=True,
            disabled=uploaded is None or invalid,
            key="luctiv_process_button",
        )
    return uploaded, manual_time, clicked


cached_generated = st.session_state.get("luctiv_generated")
has_cached_result = cached_generated is not None

if isinstance(cached_generated, GeneratedPackage):
    elapsed_seconds = float(st.session_state.get("luctiv_elapsed_seconds", 0.0))
    total_stages = sum(
        len(item.generated.result.stages) for item in cached_generated.workbooks
    )
    total_clusters = sum(
        len(item.generated.result.clusters) for item in cached_generated.workbooks
    )
    total_survey = sum(
        len(item.generated.result.survey) for item in cached_generated.workbooks
    )
    failed_items = [
        issue for issue in cached_generated.issues if issue.category == "error"
    ]
    ignored_items = [
        issue for issue in cached_generated.issues if issue.category == "ignored"
    ]

    title_col, new_col, download_col = st.columns(
        [5.0, 1.0, 1.35],
        gap="small",
        vertical_alignment="center",
    )
    with title_col:
        st.markdown(
            f"""
            <div class="luctiv-result-heading">
                <h1>LUCTIV · PAD procesado</h1>
                <p>{cached_generated.processed_count} Excel terminados ·
                {cached_generated.failed_count} archivos requieren revisión</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with new_col:
        if st.button("Nuevo archivo", use_container_width=True):
            st.session_state.pop("luctiv_generated", None)
            st.session_state.pop("luctiv_elapsed_seconds", None)
            st.session_state.pop("luctiv_upload_fingerprint", None)
            st.rerun()
    with download_col:
        st.download_button(
            "Descargar ZIP",
            data=cached_generated.data,
            file_name=cached_generated.output_filename,
            mime="application/zip",
            on_click="ignore",
            use_container_width=True,
            type="primary",
        )

    metrics = st.columns(6)
    metrics[0].metric("Procesados", cached_generated.processed_count)
    metrics[1].metric("Con error", cached_generated.failed_count)
    metrics[2].metric("Auxiliares", cached_generated.ignored_count)
    metrics[3].metric("Etapas", total_stages)
    metrics[4].metric("Clústeres", f"{total_clusters:,}".replace(",", "."))
    metrics[5].metric("Tiempo", f"{elapsed_seconds:.2f} s")

    if failed_items:
        subject = "pozo no se pudo generar" if len(failed_items) == 1 else "pozos no se pudieron generar"
        st.warning(
            f"{len(failed_items)} {subject}. Revisá el aviso antes de usar el lote. "
            "Los demás Excel están disponibles en el ZIP de descarga.",
            icon="⚠️",
        )

    st.subheader("Resultados por pozo")
    for item in cached_generated.workbooks:
        result = item.generated.result
        st.success(
            f"{result.well_name}: {len(result.stages)} etapas, "
            f"{len(result.clusters)} clústeres y {len(result.survey)} registros Survey.",
            icon="✅",
        )
    if failed_items:
        with st.expander(
            f"Archivos que requieren revisión ({len(failed_items)})",
            expanded=True,
        ):
            for issue in failed_items:
                st.warning(f"**{issue.source_filename}:** {issue.message}", icon="⚠️")
    if ignored_items:
        with st.expander(f"Archivos auxiliares ignorados ({len(ignored_items)})"):
            for issue in ignored_items:
                st.info(f"**{issue.source_filename}:** {issue.message}")
    st.caption(
        f"Survey total leído: {total_survey:,} registros. El ZIP descargable también "
        "incluye Survey TXT/CSV por pozo y RESUMEN_PROCESAMIENTO.txt."
    )
    st.stop()

if has_cached_result:
    st.markdown(
        """
        <style>
        .block-container {
            max-width: 100%;
            min-height: 100vh;
            padding: 2.2rem 1.15rem 0.15rem;
        }
        .block-container > div:first-child > div[data-testid="stVerticalBlock"] {
            gap: 0.34rem;
        }
        div[data-testid="stPlotlyChart"] { margin-top: 0; }
        @media (max-width: 900px) {
            .luctiv-kpis { grid-template-columns: repeat(3, minmax(0, 1fr)); }
            .luctiv-summary { height: auto; padding-bottom: 1rem; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    cached_result = cached_generated.result
    title_col, file_col, controls_col, download_col, txt_col, csv_col = st.columns(
        [3.8, 0.8, 0.85, 1.15, 0.9, 0.9],
        gap="small",
        vertical_alignment="center",
    )
    with file_col:
        uploaded_file, manual_minutes, process_clicked = _render_upload_controls(
            compact=True
        )

    elapsed_seconds = float(st.session_state.get("luctiv_elapsed_seconds", 0.0))
    cached_impact = calculate_impact_metrics(
        cached_result,
        elapsed_seconds,
        manual_minutes,
    )
    with title_col:
        st.markdown(
            f"""
            <div class="luctiv-result-heading">
                <h1>LUCTIV · {escape(cached_result.well_name)}</h1>
                <p>Análisis completado · Excel terminado listo para revisar</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with controls_col:
        with st.popover("Controles"):
            st.markdown(
                f"**{cached_impact.checks_passed}/{cached_impact.checks_passed} validaciones aprobadas**"
            )
            for check in cached_result.checks:
                st.caption(f"✅ {check}")
            if cached_result.warnings:
                st.divider()
                for warning in cached_result.warnings:
                    st.warning(warning, icon="⚠️")
            else:
                st.caption("✅ Sin sobreescrituras u observaciones")
            if cached_result.source_sheets:
                st.divider()
                st.caption("**Origen detectado**")
                for logical_name, location in cached_result.source_sheets.items():
                    st.caption(f"{logical_name}: {location}")
            st.divider()
            st.caption(
                f"Referencia manual: {manual_minutes} min · "
                f"automatizado: {cached_impact.elapsed_seconds:.2f} s · "
                f"eficiencia estimada: {cached_impact.estimated_efficiency_percent:.2f}%"
            )
    with download_col:
        st.download_button(
            "Descargar Excel",
            data=cached_generated.data,
            file_name=cached_result.output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click="ignore",
            use_container_width=True,
            type="primary",
        )
    with txt_col:
        st.download_button(
            "Survey TXT",
            data=cached_generated.survey_txt_data,
            file_name=cached_generated.survey_txt_filename,
            mime="text/plain",
            on_click="ignore",
            use_container_width=True,
        )
    with csv_col:
        st.download_button(
            "Survey CSV",
            data=cached_generated.survey_csv_data,
            file_name=cached_generated.survey_csv_filename,
            mime="text/csv",
            on_click="ignore",
            use_container_width=True,
        )
else:
    st.markdown(
        """
        <section class="luctiv-hero">
            <div class="luctiv-kicker">EXCEL PROCESSOR</div>
            <h1 class="luctiv-title">LUCTIV</h1>
            <p class="luctiv-subtitle">
                Cargá el archivo original del pozo, procesalo y descargá el Excel terminado
                con Datos Fractura, Survey, Smart Staging y Wellbore IFS.
            </p>
        </section>
        """,
        unsafe_allow_html=True,
    )
    uploaded_file, manual_minutes, process_clicked = _render_upload_controls(
        compact=False
    )

if uploaded_file is not None:

    uploaded_bytes = uploaded_file.getvalue()
    upload_fingerprint = sha256(uploaded_bytes).hexdigest()
    if st.session_state.get("luctiv_upload_fingerprint") != upload_fingerprint:
        st.session_state["luctiv_upload_fingerprint"] = upload_fingerprint
        st.session_state.pop("luctiv_generated", None)
        st.session_state.pop("luctiv_elapsed_seconds", None)
else:
    uploaded_bytes = None
    upload_fingerprint = None
    st.session_state.pop("luctiv_upload_fingerprint", None)
    st.session_state.pop("luctiv_generated", None)
    st.session_state.pop("luctiv_elapsed_seconds", None)

if process_clicked and uploaded_file is not None:
    with st.spinner("Analizando configuraciones, Survey y punzados…"):
        started_at = perf_counter()
        try:
            generated = process_uploaded_package(
                file_bytes=uploaded_bytes,
                filename=uploaded_file.name,
                template_path=TEMPLATE_PATH,
            )
        except InvalidWorkbookError as exc:
            st.error(str(exc), icon="⛔")
            st.stop()
        except LuctivError as exc:
            st.error(str(exc), icon="⛔")
            st.stop()
        except Exception:
            st.error(
                "Ocurrió un error inesperado al procesar el archivo. "
                "Revisá que el Excel no esté dañado y que mantenga la estructura esperada.",
                icon="⛔",
            )
            st.stop()

        st.session_state["luctiv_generated"] = generated
        st.session_state["luctiv_elapsed_seconds"] = perf_counter() - started_at
        st.rerun()

generated = st.session_state.get("luctiv_generated")
if generated is not None and uploaded_file is not None:
    result = generated.result
    elapsed_seconds = float(st.session_state.get("luctiv_elapsed_seconds", 0.0))
    impact = calculate_impact_metrics(result, elapsed_seconds, manual_minutes)

    kpis = [
        ("Procesamiento", f"{impact.elapsed_seconds:.2f} s"),
        ("Valores preparados", f"{impact.values_prepared:,}".replace(",", ".")),
        ("Controles", f"{impact.checks_passed}/{impact.checks_passed}"),
        ("Ahorro estimado", f"{impact.estimated_minutes_saved:.2f} min"),
        ("Etapas", str(len(result.stages))),
        ("Clústeres", f"{len(result.clusters):,}".replace(",", ".")),
    ]
    st.markdown(
        '<div class="luctiv-kpis">'
        + "".join(
            f'<div class="luctiv-kpi"><div class="luctiv-kpi-label">{label}</div>'
            f'<div class="luctiv-kpi-value">{value}</div></div>'
            for label, value in kpis
        )
        + "</div>",
        unsafe_allow_html=True,
    )

    trajectory_col, summary_col = st.columns([2.75, 0.85], gap="small")
    with trajectory_col:
        st.markdown(
            '<div class="luctiv-chart-heading">Trayectoria 3D · pozo completo</div>',
            unsafe_allow_html=True,
        )
        figure, trajectory_source = build_well_figure(result, height=485)
        st.plotly_chart(
            figure,
            use_container_width=True,
            theme="streamlit",
            key=f"well-trajectory-{upload_fingerprint[:12]}",
            config={
                "displaylogo": False,
                "scrollZoom": True,
                "modeBarButtonsToRemove": ["select2d", "lasso2d"],
            },
        )
    with summary_col:
        source_note = (
            "Trayectoria construida con DX/DY y TVD del Survey."
            if trajectory_source == "survey"
            else "Trayectoria relativa reconstruida por curvatura mínima."
        )
        summary_items = [
            ("Etapas", len(result.stages)),
            ("Clústeres", len(result.clusters)),
            ("Survey", len(result.survey)),
            ("Wellbore", result.wellbore_row_count),
            ("Configuraciones", len(result.fracture_configs)),
            ("Sobreescrituras", result.override_count),
        ]
        st.markdown(
            '<section class="luctiv-summary"><h3>Resumen técnico</h3>'
            '<div class="luctiv-summary-grid">'
            + "".join(
                '<div class="luctiv-summary-item">'
                f'<div class="luctiv-summary-label">{label}</div>'
                f'<div class="luctiv-summary-value">{value}</div></div>'
                for label, value in summary_items
            )
            + '</div><div class="luctiv-status-list">'
            + "".join(
                '<div class="luctiv-status-row"><span class="luctiv-status-dot"></span>'
                f'{step}</div>'
                for step in ("Lectura", "Validación", "Generación", "Verificación")
            )
            + f'</div><p class="luctiv-source-note">{source_note}</p></section>',
            unsafe_allow_html=True,
        )

if not has_cached_result:
    st.markdown(
        """
        <div class="luctiv-note">
            Los archivos se procesan en memoria durante la sesión. LUCTIV no necesita que
            el usuario instale Excel, Python ni ningún programa adicional.
        </div>
        """,
        unsafe_allow_html=True,
    )
