from __future__ import annotations

from pathlib import Path

import streamlit as st

from processor import InvalidWorkbookError, LuctivError, process_uploaded_workbook


APP_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = APP_DIR / "assets" / "plantilla_datos_terminados.xlsx"
SUPPORTED_SUFFIXES = {".xlsm", ".xlsx"}

st.set_page_config(
    page_title="LUCTIV",
    page_icon="⚙️",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    :root {
        --luctiv-ink: #13202B;
        --luctiv-blue: #1677A3;
        --luctiv-cyan: #55BED2;
        --luctiv-soft: #EFF8FA;
        --luctiv-border: #D8E6EA;
    }
    .stApp { background: linear-gradient(180deg, #F7FBFC 0%, #FFFFFF 36%); }
    .block-container { max-width: 920px; padding-top: 2.4rem; padding-bottom: 3rem; }
    .luctiv-hero {
        border: 1px solid var(--luctiv-border);
        border-radius: 22px;
        padding: 2rem 2.1rem;
        background: rgba(255,255,255,0.94);
        box-shadow: 0 18px 48px rgba(21, 66, 83, 0.08);
        margin-bottom: 1.4rem;
    }
    .luctiv-kicker {
        color: var(--luctiv-blue);
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
        color: #526672;
        max-width: 680px;
        font-size: 1.04rem;
        margin: 1rem 0 0 0;
        line-height: 1.6;
    }
    div[data-testid="stFileUploader"] {
        border: 1px dashed #8BC8D5;
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
        background: linear-gradient(135deg, var(--luctiv-blue), var(--luctiv-cyan));
        color: white;
    }
    .luctiv-note {
        color: #617581;
        font-size: 0.86rem;
        border-top: 1px solid var(--luctiv-border);
        padding-top: 1rem;
        margin-top: 1.5rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

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

uploaded_file = st.file_uploader(
    "Archivo del pozo",
    type=["xlsm", "xlsx"],
    help="El archivo debe contener las hojas Input, Survey y Punzados.",
)

invalid_extension = (
    uploaded_file is not None
    and Path(uploaded_file.name).suffix.lower() not in SUPPORTED_SUFFIXES
)
if invalid_extension:
    st.error("LUCTIV solo acepta archivos Excel .xlsm o .xlsx.", icon="⛔")

process_clicked = st.button(
    "Procesar archivo",
    type="primary",
    use_container_width=True,
    disabled=uploaded_file is None or invalid_extension,
)

if uploaded_file is not None:
    size_mb = uploaded_file.size / (1024 * 1024)
    st.caption(f"Archivo seleccionado: **{uploaded_file.name}** · {size_mb:.2f} MB")

if process_clicked and uploaded_file is not None:
    with st.spinner("Analizando configuraciones, Survey y punzados…"):
        try:
            generated = process_uploaded_workbook(
                file_bytes=uploaded_file.getvalue(),
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

    result = generated.result
    st.success(f"{result.well_name} fue procesado correctamente.", icon="✅")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Pozo", result.well_name)
    col2.metric("Etapas", len(result.stages))
    col3.metric("Clústeres", len(result.clusters))
    col4.metric("Survey", len(result.survey))

    col5, col6, col7 = st.columns(3)
    col5.metric("Configuraciones", len(result.fracture_configs))
    col6.metric("Filas Wellbore", result.wellbore_row_count)
    col7.metric("Sobreescrituras", result.override_count)

    with st.expander("Ver validaciones", expanded=True):
        for check in result.checks:
            st.write(f"✅ {check}")
        if result.warnings:
            st.divider()
            for warning in result.warnings:
                st.warning(warning, icon="⚠️")
        else:
            st.write("✅ Sin sobreescrituras u observaciones detectadas")

    st.download_button(
        "Descargar Excel terminado",
        data=generated.data,
        file_name=result.output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown(
    """
    <div class="luctiv-note">
        Los archivos se procesan en memoria durante la sesión. LUCTIV no necesita que
        el usuario instale Excel, Python ni ningún programa adicional.
    </div>
    """,
    unsafe_allow_html=True,
)
