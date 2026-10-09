from __future__ import annotations

from io import BytesIO
import os
from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
import pytest

from processor import (
    GeneratedPackage,
    InvalidWorkbookError,
    ProcessingDecisionRequired,
    analyze_workbook,
    generate_finished_workbook,
    process_uploaded_package,
    process_uploaded_workbook,
)


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "assets" / "plantilla_datos_terminados.xlsx"
SAMPLES = Path(os.environ.get("LUCTIV_SAMPLE_DIR", ROOT / "samples"))


def _xlsx_bytes(workbook: Workbook) -> bytes:
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def _default_configs(stage_count: int, clusters: int = 10, spf: int = 4) -> list[dict[str, int | str]]:
    return [
        {
            "label": f"Etapas 1-{stage_count}",
            "start": 1,
            "end": stage_count,
            "clusters": clusters,
            "spf": spf,
        }
    ]


def _config_for_stage(configs: list[dict[str, int | str]], stage: int) -> dict[str, int | str]:
    for config in configs:
        if int(config["start"]) <= stage <= int(config["end"]):
            return config
    return {"label": "sin config", "start": stage, "end": stage, "clusters": 10, "spf": 4}


def make_source_workbook(
    *,
    well_name: str = "LajE-32(h)",
    stage_count: int = 3,
    survey_rows: int = 5,
    configs: list[dict[str, int | str]] | None = None,
    stage_numbers: list[int] | None = None,
    clusters_by_stage: dict[int, int] | None = None,
    spf_by_stage: dict[int, int] | None = None,
    duplicate_cluster_number: bool = False,
    bad_depth_stage: int | None = None,
    overrides: dict[tuple[int, int], str] | None = None,
    omit_sheets: set[str] | None = None,
    empty_punzados: bool = False,
    stage_top_md: dict[int, float] | None = None,
) -> bytes:
    input_configs = _default_configs(stage_count) if configs is None else configs
    cluster_configs = input_configs or _default_configs(stage_count)
    stage_numbers = stage_numbers or list(range(1, stage_count + 1))
    clusters_by_stage = clusters_by_stage or {}
    spf_by_stage = spf_by_stage or {}
    overrides = overrides or {}
    omit_sheets = omit_sheets or set()
    stage_top_md = stage_top_md or {}

    workbook = Workbook()
    input_sheet = workbook.active
    input_sheet.title = "Input"
    survey_sheet = workbook.create_sheet("Survey")
    punzados_sheet = workbook.create_sheet("Punzados")

    input_sheet["A1"] = "Nombre"
    input_sheet["B1"] = well_name
    input_sheet.append([])
    input_sheet.append(["Etapas", "Inicio", "Fin", "N° Cl", "SPF"])
    for config in input_configs:
        input_sheet.append(
            [
                config["label"],
                config["start"],
                config["end"],
                config["clusters"],
                config["spf"],
            ]
        )

    survey_sheet.append(["fila separadora", None, None, None])
    survey_sheet.append(["MD", "INCL", "AZIM_TN", "TVD"])
    survey_sheet.append(["texto", "sin", "md", "numerico"])
    for idx in range(survey_rows):
        survey_sheet.append(
            [
                f"{1000 + idx * 10:.2f}" if idx == 0 else 1000 + idx * 10,
                80 + idx * 0.1,
                120 + idx * 0.2,
                900 + idx * 8.5,
            ]
        )

    punzados_sheet.append(
        [
            "# Cluster",
            "Tope Cluster MD (m)",
            "Base Cluster MD (m)",
            "Número etapa",
            "N° de tiros x cluster",
            "Longitud de etapa",
            "Cantidad de clusters",
            "En caso de cambio de punzados, sobreescribir datos (NO BORRAR)",
        ]
    )
    cluster_number = 1
    if not empty_punzados:
        for stage in stage_numbers:
            config = _config_for_stage(cluster_configs, stage)
            cluster_count = clusters_by_stage.get(stage, int(config["clusters"]))
            spf = spf_by_stage.get(stage, int(config["spf"]))
            for index in range(cluster_count):
                top_md = stage_top_md.get(stage, 6000 + stage * 20) + index * 1.5
                base_md = top_md + 1.1
                if bad_depth_stage == stage:
                    base_md = top_md - 0.5
                number = 1 if duplicate_cluster_number and cluster_number == 2 else cluster_number
                punzados_sheet.append(
                    [
                        number,
                        top_md,
                        base_md,
                        stage,
                        spf,
                        50,
                        int(config["clusters"]),
                        overrides.get((stage, index + 1)),
                    ]
                )
                cluster_number += 1

    for sheet_name in omit_sheets:
        del workbook[sheet_name]

    return _xlsx_bytes(workbook)


def test_missing_required_sheet_names_the_exact_sheet():
    data = make_source_workbook(omit_sheets={"Punzados"})

    with pytest.raises(InvalidWorkbookError, match='No se encontró la hoja "Punzados"'):
        analyze_workbook(data, "pozo.xlsx")


def test_empty_survey_is_rejected():
    data = make_source_workbook(survey_rows=0)

    with pytest.raises(InvalidWorkbookError, match="Survey"):
        analyze_workbook(data, "pozo.xlsx")


def test_optional_survey_displacements_are_preserved_for_visualization():
    workbook = load_workbook(BytesIO(make_source_workbook(survey_rows=2)))
    survey = workbook["Survey"]
    survey["E2"] = "DX"
    survey["F2"] = "DY"
    survey["E4"] = 12.345
    survey["F4"] = 67.891
    survey["E5"] = 20.0
    survey["F5"] = 80.0

    result = analyze_workbook(_xlsx_bytes(workbook), "pozo.xlsx")

    assert result.survey[0].dx == 12.35
    assert result.survey[0].dy == 67.89


def test_alternative_spanish_survey_headers_are_detected():
    workbook = load_workbook(BytesIO(make_source_workbook(survey_rows=2)))
    survey = workbook["Survey"]
    survey["A2"] = "Prof [m]"
    survey["B2"] = "Desviac. Vert [°]"
    survey["C2"] = "Azimuth [°]"
    survey["D2"] = "Pfv [m]"
    survey["E2"] = "Dis NS [m]"
    survey["F2"] = "Dis EO [m]"
    survey["E4"] = 11.25
    survey["F4"] = -4.75
    survey["E5"] = 14.0
    survey["F5"] = -8.0

    result = analyze_workbook(_xlsx_bytes(workbook), "pozo.xlsx")

    assert len(result.survey) == 2
    assert result.survey[0].dx == -4.75
    assert result.survey[0].dy == 11.25


def test_incomplete_survey_requires_confirmation_for_matching_report():
    workbook = load_workbook(BytesIO(make_source_workbook(survey_rows=3)))
    survey = workbook["Survey"]
    report = workbook.create_sheet("Survey report")
    for col, header in enumerate(("MD", "Inclination", "Azimuth", "TVD"), start=1):
        report.cell(63, col, header)
    for offset, source_row in enumerate(range(4, 7), start=65):
        for col in range(1, 5):
            report.cell(offset, col, survey.cell(source_row, col).value)
        survey.cell(source_row, 2).value = None
        survey.cell(source_row, 3).value = None

    data = _xlsx_bytes(workbook)
    with pytest.raises(ProcessingDecisionRequired) as pending:
        process_uploaded_workbook(data, "pozo.xlsx", TEMPLATE)
    assert any("Survey report" in detail for detail in pending.value.details)

    generated = process_uploaded_workbook(
        data, "pozo.xlsx", TEMPLATE, accept_recoveries=True
    )
    assert generated.result.source_sheets["Survey"] == "Survey report (fila 63)"
    assert len(generated.result.survey) == 3


def test_incomplete_survey_rejects_unrelated_report():
    workbook = load_workbook(BytesIO(make_source_workbook(survey_rows=3)))
    survey = workbook["Survey"]
    report = workbook.create_sheet("Survey report")
    report.append(["MD", "Inclination", "Azimuth", "TVD"])
    for source_row in range(4, 7):
        report.append(
            [
                survey.cell(source_row, 1).value,
                survey.cell(source_row, 2).value,
                survey.cell(source_row, 3).value,
                survey.cell(source_row, 4).value + 100,
            ]
        )
        survey.cell(source_row, 2).value = None
        survey.cell(source_row, 3).value = None

    with pytest.raises(InvalidWorkbookError, match="INCL"):
        analyze_workbook(_xlsx_bytes(workbook), "pozo.xlsx")


@pytest.mark.parametrize("base_header", ["Base Cluster MD (m)", "Fondo Cluster MD (m)"])
def test_base_and_fondo_cluster_headers_are_detected(base_header: str):
    workbook = load_workbook(BytesIO(make_source_workbook(stage_count=1, survey_rows=2)))
    workbook["Punzados"]["C1"] = base_header

    result = analyze_workbook(_xlsx_bytes(workbook), "pozo.xlsx")

    assert len(result.clusters) == 10
    assert result.clusters[0].base_md == 6021.1
    assert len(result.stages) == 1


def test_tables_are_detected_when_sheet_names_change():
    workbook = load_workbook(BytesIO(make_source_workbook(stage_count=2, survey_rows=2)))
    workbook["Input"].title = "Plan de fractura"
    workbook["Survey"].title = "Trayectoria direccional"
    workbook["Punzados"].title = "Perforaciones propuestas"

    result = analyze_workbook(_xlsx_bytes(workbook), "pozo.xlsx")

    assert result.source_sheets == {
        "Survey": "Trayectoria direccional (fila 2)",
        "Punzados": "Perforaciones propuestas (fila 1)",
        "Input": "Plan de fractura (fila 3)",
    }


def test_fracture_configs_are_inferred_when_input_table_is_missing():
    workbook = load_workbook(BytesIO(make_source_workbook(stage_count=3, survey_rows=2)))
    del workbook["Input"]

    result = analyze_workbook(_xlsx_bytes(workbook), "LLL-2000.xlsx")

    assert result.well_name == "LLL-2000"
    assert len(result.fracture_configs) == 1
    inferred = result.fracture_configs[0]
    assert (inferred.start_stage, inferred.end_stage, inferred.clusters, inferred.spf) == (
        1,
        3,
        10,
        4,
    )
    assert any("se infirió" in warning for warning in result.warnings)


def test_empty_punzados_is_rejected():
    data = make_source_workbook(empty_punzados=True)

    with pytest.raises(InvalidWorkbookError, match="Punzados"):
        analyze_workbook(data, "pozo.xlsx")


def test_non_consecutive_stages_are_rejected():
    data = make_source_workbook(stage_count=3, stage_numbers=[1, 3])

    with pytest.raises(InvalidWorkbookError, match="Etapas faltantes: \\[2\\]"):
        analyze_workbook(data, "pozo.xlsx")


def test_wrong_cluster_count_is_rejected():
    data = make_source_workbook(stage_count=2, clusters_by_stage={2: 8})

    with pytest.raises(InvalidWorkbookError, match="se encontraron 8 clústeres"):
        analyze_workbook(data, "pozo.xlsx")


def test_punzados_count_requires_confirmation_when_input_is_stale():
    workbook = load_workbook(BytesIO(make_source_workbook(stage_count=2)))
    workbook["Input"]["D4"] = 11
    data = _xlsx_bytes(workbook)

    with pytest.raises(ProcessingDecisionRequired) as pending:
        analyze_workbook(data, "pozo.xlsx")
    assert any("Punzados confirma 10" in detail for detail in pending.value.details)

    result = analyze_workbook(data, "pozo.xlsx", accept_recoveries=True)
    assert result.fracture_configs[0].clusters == 10
    assert any("Se usó Punzados" in warning for warning in result.warnings)


def test_wrong_spf_is_rejected():
    data = make_source_workbook(stage_count=2, spf_by_stage={2: 5})

    with pytest.raises(InvalidWorkbookError, match="SPF"):
        analyze_workbook(data, "pozo.xlsx")


def test_stage_with_eleven_clusters_is_valid():
    configs = [{"label": "Etapa especial", "start": 1, "end": 1, "clusters": 11, "spf": 4}]
    data = make_source_workbook(stage_count=1, configs=configs)

    result = analyze_workbook(data, "pozo.xlsx")

    assert len(result.stages) == 1
    assert len(result.clusters) == 11


def test_different_number_of_fracture_configurations_is_detected():
    configs = [
        {"label": "A", "start": 1, "end": 2, "clusters": 8, "spf": 4},
        {"label": "B", "start": 3, "end": 4, "clusters": 9, "spf": 4},
        {"label": "C", "start": 5, "end": 6, "clusters": 10, "spf": 5},
        {"label": "D", "start": 7, "end": 8, "clusters": 11, "spf": 5},
    ]
    data = make_source_workbook(stage_count=8, configs=configs)

    result = analyze_workbook(data, "pozo.xlsx")

    assert len(result.fracture_configs) == 4
    assert len(result.stages) == 8


@pytest.mark.parametrize("stage_count", [81, 82])
def test_dynamic_stage_counts_generate_valid_output(stage_count):
    data = make_source_workbook(stage_count=stage_count, survey_rows=12)
    result = analyze_workbook(data, f"LajE-{stage_count}.xlsm")

    generated = generate_finished_workbook(result, TEMPLATE)
    output_workbook = load_workbook(BytesIO(generated.data), data_only=False)
    sheet = output_workbook["Datos terminados"]

    assert len(result.stages) == stage_count
    assert sheet.cell(4, 13).value == stage_count
    assert sheet.cell(3 + stage_count, 13).value == 1
    assert sheet.cell(4, 18).value == "Treatment Interval"
    assert generated.data[:2] == b"PK"


def test_generated_workbook_cleans_residual_template_values(tmp_path):
    dirty_template = tmp_path / "dirty_template.xlsx"
    workbook = load_workbook(TEMPLATE)
    sheet = workbook["Datos terminados"]
    for row in range(12, 18):
        for col in [1, 2, 3, 4, 5, 8, 9, 10, 11, 13, 14, 15, 16, 18, 19, 20]:
            sheet.cell(row, col, "residual")
    workbook.save(dirty_template)

    data = make_source_workbook(stage_count=2, survey_rows=3)
    result = analyze_workbook(data, "pozo.xlsx")
    generated = generate_finished_workbook(result, dirty_template)
    output_workbook = load_workbook(BytesIO(generated.data), data_only=False)
    output_sheet = output_workbook["Datos terminados"]

    for row in range(12, 18):
        for col in [1, 2, 3, 4, 5, 8, 9, 10, 11, 13, 14, 15, 16, 18, 19, 20]:
            assert output_sheet.cell(row, col).value is None


def test_wellbore_has_two_rows_per_stage_in_reverse_order():
    data = make_source_workbook(stage_count=3)
    result = analyze_workbook(data, "pozo.xlsx")
    generated = generate_finished_workbook(result, TEMPLATE)
    sheet = load_workbook(BytesIO(generated.data), data_only=False)["Datos terminados"]

    rows = [sheet.cell(row, 18).value for row in range(4, 4 + result.wellbore_row_count)]

    assert result.wellbore_row_count == 6
    assert rows == [
        "Treatment Interval",
        "Perforations",
        "Treatment Interval",
        "Perforations",
        "Treatment Interval",
        "Perforations",
    ]
    assert sheet.cell(4, 19).value == result.stages[-1].top_md
    assert sheet.cell(8, 19).value == result.stages[0].top_md


def test_smart_staging_and_wellbore_keep_published_sort_order():
    data = make_source_workbook(
        stage_count=3,
        stage_top_md={1: 7000, 2: 6500, 3: 6800},
    )
    generated = process_uploaded_workbook(data, "pozo.xlsx", TEMPLATE)
    sheet = load_workbook(BytesIO(generated.data), data_only=False)["Datos terminados"]

    assert [sheet.cell(row, 13).value for row in range(4, 7)] == [3, 2, 1]
    assert [sheet.cell(row, 14).value for row in range(4, 7)] == [6800, 6500, 7000]
    assert [sheet.cell(row, 19).value for row in range(4, 10)] == [
        7000, 7000, 6800, 6800, 6500, 6500,
    ]


def test_survey_txt_and_csv_match_published_columns():
    data = make_source_workbook(stage_count=2, survey_rows=3)
    generated = process_uploaded_workbook(data, "pozo.xlsm", TEMPLATE)

    assert generated.survey_txt_filename == "LajE-32h_survey_md_inclination_azimuth.txt"
    assert generated.survey_csv_filename == "LajE-32h_survey_md_inclination_azimuth.csv"
    assert generated.survey_txt_data.decode("utf-8").splitlines() == [
        "MD\tINCLINATION\tAZIMUTH",
        "1000\t80\t120",
        "1010\t80.1\t120.2",
        "1020\t80.2\t120.4",
    ]
    assert generated.survey_csv_data.decode("utf-8").splitlines() == [
        "MD,INCLINATION,AZIMUTH",
        "1000,80,120",
        "1010,80.1,120.2",
        "1020,80.2,120.4",
    ]


def test_output_filename_is_sanitized():
    data = make_source_workbook(well_name="LajE-32(h) PAD / Norte")

    result = analyze_workbook(data, "entrada.xlsm")

    assert result.output_filename == "LajE-32h-PAD-Norte_datos_terminados.xlsx"


def test_generated_xlsx_integrity_and_formula_error_scan():
    data = make_source_workbook(stage_count=4, survey_rows=7)
    generated = process_uploaded_workbook(data, "pozo.xlsm", TEMPLATE)
    workbook = load_workbook(BytesIO(generated.data), data_only=False)
    sheet = workbook["Datos terminados"]
    values = [cell.value for row in sheet.iter_rows() for cell in row]

    assert workbook.sheetnames == ["Datos terminados"]
    assert not {"#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A"} & set(values)
    assert "Archivo .xlsx verificado con openpyxl" in generated.result.checks


def test_zip_package_generates_one_output_per_compatible_workbook():
    package_bytes = BytesIO()
    unrelated = Workbook()
    unrelated.active.title = "Notas"
    unrelated.active["A1"] = "Documento auxiliar"
    with ZipFile(package_bytes, "w") as package:
        package.writestr(
            "pozos/LLL-1.xlsx",
            make_source_workbook(well_name="LLL-1", stage_count=2, survey_rows=3),
        )
        package.writestr(
            "pozos/LLL-2.xlsm",
            make_source_workbook(well_name="LLL-2", stage_count=3, survey_rows=4),
        )
        package.writestr(
            "pozos/LLL-3-incompleto.xlsx",
            make_source_workbook(
                well_name="LLL-3",
                stage_count=2,
                stage_numbers=[1],
                survey_rows=3,
            ),
        )
        package.writestr("documentacion/notas.xlsx", _xlsx_bytes(unrelated))
        package.writestr("mapa.png", b"not needed")

    generated = process_uploaded_package(
        package_bytes.getvalue(),
        "PAD nuevo.zip",
        TEMPLATE,
    )

    assert isinstance(generated, GeneratedPackage)
    assert generated.processed_count == 2
    assert generated.issue_count == 2
    assert generated.failed_count == 1
    assert generated.ignored_count == 1
    assert generated.output_filename == "PAD-nuevo_terminados.zip"
    with ZipFile(BytesIO(generated.data)) as output:
        names = set(output.namelist())
        assert names == {
            "LLL-1_datos_terminados.xlsx",
            "LLL-1_survey_md_inclination_azimuth.txt",
            "LLL-1_survey_md_inclination_azimuth.csv",
            "LLL-2_datos_terminados.xlsx",
            "LLL-2_survey_md_inclination_azimuth.txt",
            "LLL-2_survey_md_inclination_azimuth.csv",
            "RESUMEN_PROCESAMIENTO.txt",
        }
        summary = output.read("RESUMEN_PROCESAMIENTO.txt").decode("utf-8")
        assert "Archivos generados: 2" in summary
        assert "Archivos con error: 1" in summary
        assert "LLL-3-incompleto.xlsx" in summary
        assert "notas.xlsx" in summary


def test_generated_numeric_outputs_use_expected_number_formats():
    data = make_source_workbook(stage_count=3, survey_rows=5)
    generated = process_uploaded_workbook(data, "pozo.xlsm", TEMPLATE)
    sheet = load_workbook(BytesIO(generated.data), data_only=False)["Datos terminados"]

    numeric_ranges = (
        (3, 2 + len(generated.result.fracture_configs), 2, 5, "0"),
        (4, 3 + len(generated.result.survey), 8, 11, "General"),
        (4, 3 + len(generated.result.stages), 13, 13, "0"),
        (4, 3 + len(generated.result.stages), 14, 16, "General"),
        (4, 3 + generated.result.wellbore_row_count, 19, 20, "General"),
    )
    for min_row, max_row, min_col, max_col, expected_format in numeric_ranges:
        for row in sheet.iter_rows(
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
        ):
            for cell in row:
                assert isinstance(cell.value, (int, float))
                assert cell.number_format == expected_format

    assert sheet["M4"].value == 3
    assert sheet["M4"].number_format == "0"
    assert sheet["N4"].number_format == "General"
    assert sheet["P4"].number_format == "General"


def test_duplicate_cluster_numbers_are_rejected():
    data = make_source_workbook(duplicate_cluster_number=True)

    with pytest.raises(InvalidWorkbookError, match="duplicados"):
        analyze_workbook(data, "pozo.xlsx")


def test_tope_must_be_less_than_fondo():
    data = make_source_workbook(bad_depth_stage=1)

    with pytest.raises(InvalidWorkbookError, match="Tope"):
        analyze_workbook(data, "pozo.xlsx")


def test_overrides_are_reported_as_warning():
    data = make_source_workbook(overrides={(1, 1): "Revisar punzado"})

    result = analyze_workbook(data, "pozo.xlsx")

    assert result.override_count == 1
    assert any("sobreescrituras" in warning for warning in result.warnings)


def test_invalid_upload_extension_is_rejected():
    data = make_source_workbook()

    with pytest.raises(InvalidWorkbookError, match=".xlsm o .xlsx"):
        process_uploaded_workbook(data, "pozo.txt", TEMPLATE)


@pytest.mark.parametrize(
    "filename,expected",
    [
        ("LajE-35(h)_version2.xlsm", (82, 784, 225, 3)),
        ("LajE-34(h)_version2.xlsm", (82, 784, 229, 3)),
        ("LajE-33(h)_version2.xlsm", (82, 785, 238, 4)),
        ("LajE-32(h)_version2.xlsm", (81, 779, 612, 4)),
    ],
)
def test_known_wells_if_available(filename, expected):
    path = SAMPLES / filename
    if not path.exists():
        pytest.skip(f"Sample not available: {path}")

    result = analyze_workbook(path.read_bytes(), filename)

    assert (
        len(result.stages),
        len(result.clusters),
        len(result.survey),
        len(result.fracture_configs),
    ) == expected
