from __future__ import annotations

from pathlib import Path

import streamlit as st

from processor import InvalidWorkbookError, LuctivError, process_uploaded_workbook


APP_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = APP_DIR / "assets" / "plantilla_datos_terminados.xlsx"
SUPPORTED_SUFFIXES = {".xlsm", ".xlsx"}

st.set_page_config(
    page_title="LUCTIV",
    page_icon="L",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    :root {
        --luctiv-ink: #121820;
        --luctiv-muted: #52616B;
        --luctiv-red: #D71920;
        --luctiv-teal: #0F8B8D;
        --luctiv-gold: #F4B942;
        --luctiv-soft: #F5F8FA;
        --luctiv-border: #D9E2E7;
    }
    .stApp {
        background:
            radial-gradient(circle at 8% 0%, rgba(215, 25, 32, 0.10), transparent 28rem),
            radial-gradient(circle at 95% 8%, rgba(15, 139, 141, 0.10), transparent 26rem),
            linear-gradient(180deg, #F7F9FA 0%, #FFFFFF 46%);
    }
    .block-container { max-width: 960px; padding-top: 2rem; padding-bottom: 3rem; }
    .luctiv-hero {
        border: 1px solid var(--luctiv-border);
        border-left: 6px solid var(--luctiv-red);
        border-radius: 8px;
        padding: clamp(1.45rem, 4vw, 2.35rem);
        background: rgba(255, 255, 255, 0.96);
        box-shadow: 0 20px 50px rgba(18, 24, 32, 0.08);
        margin-bottom: 1.25rem;
        overflow: hidden;
    }
    .luctiv-brand-row {
        align-items: center;
        display: flex;
        gap: 0.8rem;
        margin-bottom: 0.9rem;
    }
    .luctiv-mark {
        align-items: center;
        background: linear-gradient(135deg, var(--luctiv-red) 0 52%, var(--luctiv-teal) 52% 100%);
        border-radius: 8px;
        box-shadow: 0 10px 24px rgba(215, 25, 32, 0.18);
        color: white;
        display: flex;
        font-size: 1.18rem;
        font-weight: 900;
        height: 44px;
        justify-content: center;
        line-height: 1;
        width: 44px;
    }
    .luctiv-title {
        color: var(--luctiv-ink);
        font-size: clamp(2.7rem, 8vw, 5.1rem);
        letter-spacing: 0;
        line-height: 0.95;
        font-weight: 900;
        margin: 0;
    }
    .luctiv-subtitle {
        color: var(--luctiv-muted);
        max-width: 760px;
        font-size: 1.08rem;
        margin: 0;
        line-height: 1.62;
    }
    .luctiv-rule {
        background: linear-gradient(90deg, var(--luctiv-red), var(--luctiv-gold), var(--luctiv-teal));
        border-radius: 999px;
        height: 4px;
        margin-top: 1.15rem;
        width: 132px;
    }
    div[data-testid="stFileUploader"] {
        border: 1px dashed #9CB7C3;
        border-radius: 8px;
        padding: 0.65rem 0.85rem 0.25rem;
        background: rgba(255, 255, 255, 0.92);
        box-shadow: 0 12px 32px rgba(18, 24, 32, 0.05);
    }
    div[data-testid="stMetric"] {
        background: #FFFFFF;
        border: 1px solid var(--luctiv-border);
        border-radius: 8px;
        padding: 0.72rem 0.8rem;
    }
    div.stButton > button, div.stDownloadButton > button {
        border-radius: 8px;
        min-height: 3rem;
        font-weight: 750;
        border: none;
    }
    div.stButton > button[kind="primary"], div.stDownloadButton > button {
        background: linear-gradient(135deg, var(--luctiv-red), var(--luctiv-teal));
        color: white;
    }
    div[data-testid="stExpander"] {
        border: 1px solid var(--luctiv-border);
        border-radius: 8px;
        box-shadow: 0 10px 26px rgba(18, 24, 32, 0.04);
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
        <div class="luctiv-brand-row">
            <div class="luctiv-mark">L</div>
            <h1 class="luctiv-title">LUCTIV</h1>
        </div>
        <p class="luctiv-subtitle">
            Hola inge! Carg&aacute; el archivo del pozo, procesalo y descarg&aacute; el archivo excel terminado con los datos para hacer Smart Staging y IFS. Enjoy it.
        </p>
        <div class="luctiv-rule"></div>
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
