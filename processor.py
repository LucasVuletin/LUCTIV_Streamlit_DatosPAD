from __future__ import annotations

from copy import copy
from dataclasses import dataclass, field, replace
from io import BytesIO
import math
from pathlib import Path
from pathlib import PurePosixPath
import re
import unicodedata
from typing import BinaryIO, Iterable
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


REQUIRED_SHEETS = ("Input", "Survey", "Punzados")
TEMPLATE_SHEET = "Datos terminados"
ALLOWED_EXTENSIONS = {".xlsm", ".xlsx"}
ALLOWED_UPLOAD_EXTENSIONS = ALLOWED_EXTENSIONS | {".zip"}
FORMULA_ERROR_VALUES = {"#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A"}
MAX_PACKAGE_EXCEL_FILES = 50
MAX_PACKAGE_ENTRY_BYTES = 25 * 1024 * 1024
MAX_PACKAGE_TOTAL_BYTES = 150 * 1024 * 1024

CONFIG_HEADER_GROUPS = (
    ("Etapas", "Rango de etapas", "Stage range"),
    ("Inicio", "Etapa inicial", "Start stage", "From stage"),
    ("Fin", "Etapa final", "End stage", "To stage"),
    ("N° Cl", "N Cl", "Nro Cl", "Cantidad de clusters", "Clusters", "Cluster count"),
    ("SPF", "Tiros por pie", "Shots per foot"),
)
SURVEY_HEADER_GROUPS = (
    ("MD", "Measured depth", "Prof [m]", "Profundidad medida"),
    ("TVD", "True vertical depth", "Pfv [m]", "Profundidad vertical"),
    ("INCL", "Inclination", "Inclinacion", "Desviac. Vert", "Desviacion vertical"),
    ("AZIM_TN", "Azimuth", "Azimut", "AZIM"),
)
CLUSTER_BASE_HEADER_ALIASES = (
    "Base Cluster MD",
    "Fondo Cluster MD",
    "Base cluster",
    "Fondo cluster",
    "Cluster base",
    "Bottom MD",
)
CLUSTER_HEADER_GROUPS = (
    ("# Cluster", "Numero de cluster", "Cluster number"),
    ("Tope Cluster MD", "Tope cluster", "Cluster top", "Top MD"),
    CLUSTER_BASE_HEADER_ALIASES,
    ("Numero etapa", "Etapa", "Stage number"),
    ("N° de tiros x cluster", "Tiros x cluster", "SPF", "Shots per foot"),
)

TABLE_SHEET_ALIASES = {
    "Input": ("Input", "Entrada", "Datos de entrada", "Configuracion"),
    "Survey": ("Survey", "Directional Survey", "Direccional", "Trayectoria"),
    "Punzados": ("Punzados", "Perforaciones", "Perforation", "Clusters"),
}


class LuctivError(Exception):
    """Base error shown to the end user."""


class InvalidWorkbookError(LuctivError):
    """Raised when the uploaded workbook does not match the expected structure."""


class ProcessingDecisionRequired(LuctivError):
    """A recoverable source discrepancy needs the user's choice before generating output."""

    def __init__(self, details: Iterable[str]):
        self.details = tuple(details)
        super().__init__("Se encontraron diferencias entre las hojas del archivo.")


@dataclass(frozen=True)
class FractureConfig:
    label: str
    start_stage: int
    end_stage: int
    clusters: int
    spf: int


@dataclass(frozen=True)
class SurveyPoint:
    md: float
    inclination: float
    azimuth: float
    tvd: float
    dx: float | None = None
    dy: float | None = None


@dataclass(frozen=True)
class Cluster:
    cluster_number: int
    top_md: float
    base_md: float
    stage: int
    spf: int
    expected_clusters: int | None = None
    stage_length: float | None = None
    override_note: str | None = None


@dataclass(frozen=True)
class StageInterval:
    stage: int
    top_md: float
    base_md: float
    plug_md: float


@dataclass
class ProcessingResult:
    well_name: str
    output_filename: str
    fracture_configs: list[FractureConfig]
    survey: list[SurveyPoint]
    clusters: list[Cluster]
    stages: list[StageInterval]
    warnings: list[str] = field(default_factory=list)
    checks: list[str] = field(default_factory=list)
    source_sheets: dict[str, str] = field(default_factory=dict)

    @property
    def wellbore_row_count(self) -> int:
        return len(self.stages) * 2

    @property
    def override_count(self) -> int:
        return sum(1 for cluster in self.clusters if cluster.override_note)

@dataclass(frozen=True)
class GeneratedWorkbook:
    result: ProcessingResult
    data: bytes
    survey_txt_data: bytes
    survey_csv_data: bytes

    @property
    def survey_txt_filename(self) -> str:
        stem = self.result.output_filename.removesuffix("_datos_terminados.xlsx")
        return f"{stem}_survey_md_inclination_azimuth.txt"

    @property
    def survey_csv_filename(self) -> str:
        stem = self.result.output_filename.removesuffix("_datos_terminados.xlsx")
        return f"{stem}_survey_md_inclination_azimuth.csv"


@dataclass(frozen=True)
class PackageWorkbook:
    source_filename: str
    generated: GeneratedWorkbook


@dataclass(frozen=True)
class PackageIssue:
    source_filename: str
    message: str
    category: str = "error"


@dataclass(frozen=True)
class GeneratedPackage:
    output_filename: str
    data: bytes
    workbooks: tuple[PackageWorkbook, ...]
    issues: tuple[PackageIssue, ...]

    @property
    def processed_count(self) -> int:
        return len(self.workbooks)

    @property
    def issue_count(self) -> int:
        return len(self.issues)

    @property
    def failed_count(self) -> int:
        return sum(1 for issue in self.issues if issue.category == "error")

    @property
    def ignored_count(self) -> int:
        return sum(1 for issue in self.issues if issue.category == "ignored")


def _normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^a-z0-9#]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _is_number(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _coerce_float(value: object) -> float | None:
    if _is_number(value):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(" ", "")
        if not cleaned:
            return None
        if "," in cleaned and "." in cleaned:
            if cleaned.rfind(",") > cleaned.rfind("."):
                cleaned = cleaned.replace(".", "").replace(",", ".")
            else:
                cleaned = cleaned.replace(",", "")
        else:
            cleaned = cleaned.replace(",", ".")
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def _as_float(value: object, field_name: str, row_number: int) -> float:
    number = _coerce_float(value)
    if number is not None:
        return number
    raise InvalidWorkbookError(
        f"El campo '{field_name}' de la fila {row_number} no es numérico: {value!r}."
    )


def _as_int(value: object, field_name: str, row_number: int) -> int:
    number = _as_float(value, field_name, row_number)
    if not number.is_integer():
        raise InvalidWorkbookError(
            f"El campo '{field_name}' de la fila {row_number} debe ser entero: {value!r}."
        )
    return int(number)


def _find_header_row(
    sheet: Worksheet,
    required_terms: Iterable[str],
    max_scan_rows: int = 60,
) -> int:
    required = [_normalize_text(term) for term in required_terms]
    for row_idx in range(1, min(sheet.max_row, max_scan_rows) + 1):
        values = [_normalize_text(cell.value) for cell in sheet[row_idx]]
        if all(any(term == value or term in value for value in values) for term in required):
            return row_idx
    raise InvalidWorkbookError(
        f"No se pudo localizar el encabezado esperado en la hoja '{sheet.title}'."
    )


def _find_header_row_by_alias_groups(
    sheet: Worksheet,
    required_groups: Iterable[Iterable[str]],
    max_scan_rows: int = 60,
    max_scan_columns: int = 80,
) -> int:
    """Find a header row by canonical fields instead of literal column names."""
    groups = tuple(tuple(group) for group in required_groups)
    last_row = min(sheet.max_row, max_scan_rows)
    last_column = min(sheet.max_column, max_scan_columns)
    for row_idx in range(1, last_row + 1):
        headers = {
            cell.column: _normalize_text(cell.value)
            for cell in sheet[row_idx][:last_column]
            if cell.value not in (None, "")
        }
        if headers and all(_find_column_or_none(headers, aliases) is not None for aliases in groups):
            return row_idx
    raise InvalidWorkbookError(
        f"No se pudo localizar el encabezado esperado en la hoja '{sheet.title}'."
    )


def _sheet_name_priority(sheet_name: str, aliases: Iterable[str]) -> int:
    normalized_name = _normalize_text(sheet_name)
    normalized_aliases = tuple(_normalize_text(alias) for alias in aliases)
    if normalized_name in normalized_aliases:
        return 2
    if any(alias and alias in normalized_name for alias in normalized_aliases):
        return 1
    return 0


def _locate_table(
    workbook: Workbook,
    logical_name: str,
    required_groups: Iterable[Iterable[str]],
    *,
    max_scan_rows: int,
) -> tuple[Worksheet, int] | None:
    """Locate one logical table anywhere in the workbook, preferring familiar sheet names."""
    aliases = TABLE_SHEET_ALIASES[logical_name]
    candidates: list[tuple[int, int, Worksheet, int]] = []
    for sheet_index, sheet in enumerate(workbook.worksheets):
        try:
            header_row = _find_header_row_by_alias_groups(
                sheet,
                required_groups,
                max_scan_rows=max_scan_rows,
            )
        except InvalidWorkbookError:
            continue
        candidates.append(
            (_sheet_name_priority(sheet.title, aliases), sheet_index, sheet, header_row)
        )

    if not candidates:
        return None

    candidates.sort(key=lambda item: (-item[0], item[1], item[3]))
    best_priority = candidates[0][0]
    best = [candidate for candidate in candidates if candidate[0] == best_priority]
    if len(best) > 1:
        sheet_names = ", ".join(f"'{candidate[2].title}'" for candidate in best[:8])
        raise InvalidWorkbookError(
            f"Se encontraron varias tablas posibles para {logical_name}: {sheet_names}. "
            "No se puede elegir una automáticamente sin riesgo."
        )
    return candidates[0][2], candidates[0][3]


def _find_column(headers: dict[int, str], aliases: Iterable[str]) -> int:
    normalized_aliases = [_normalize_text(alias) for alias in aliases]
    for alias in normalized_aliases:
        for col_idx, header in headers.items():
            if header == alias:
                return col_idx
    for alias in normalized_aliases:
        for col_idx, header in headers.items():
            if alias in header:
                return col_idx
    raise InvalidWorkbookError(
        f"Falta una columna requerida. Se buscó: {', '.join(aliases)}."
    )


def _find_column_or_none(headers: dict[int, str], aliases: Iterable[str]) -> int | None:
    try:
        return _find_column(headers, aliases)
    except InvalidWorkbookError:
        return None


def _extract_well_name(input_sheet: Worksheet, fallback_filename: str) -> str:
    for row in input_sheet.iter_rows(min_row=1, max_row=min(input_sheet.max_row, 20)):
        for idx, cell in enumerate(row):
            if _normalize_text(cell.value).rstrip(":") == "nombre":
                if idx + 1 < len(row) and row[idx + 1].value:
                    return str(row[idx + 1].value).strip()
    return Path(fallback_filename).stem.replace("_version2", "")


def _safe_output_filename(well_name: str) -> str:
    cleaned = well_name.replace("(h)", "h").replace("(H)", "H")
    cleaned = re.sub(r"[()]", "", cleaned)
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "-", cleaned).strip("-_.")
    return f"{cleaned or 'pozo'}_datos_terminados.xlsx"


def _format_plain_number(value: float | int) -> str:
    number = float(value)
    if math.isclose(number, round(number), abs_tol=0.000000001):
        return str(int(round(number)))
    return f"{number:.10f}".rstrip("0").rstrip(".")


def _survey_delimited_data(result: ProcessingResult, delimiter: str) -> bytes:
    lines = [delimiter.join(("MD", "INCLINATION", "AZIMUTH"))]
    for point in result.survey:
        lines.append(
            delimiter.join(
                (
                    _format_plain_number(point.md),
                    _format_plain_number(point.inclination),
                    _format_plain_number(point.azimuth),
                )
            )
        )
    return ("\r\n".join(lines) + "\r\n").encode("utf-8")


def _smart_staging_stages(result: ProcessingResult) -> list[StageInterval]:
    return sorted(result.stages, key=lambda stage: stage.stage, reverse=True)


def _wellbore_ifs_stages(result: ProcessingResult) -> list[StageInterval]:
    return sorted(
        result.stages,
        key=lambda stage: (stage.top_md, stage.base_md, stage.stage),
        reverse=True,
    )


def _extract_fracture_configs(
    sheet: Worksheet,
    header_row: int | None = None,
) -> list[FractureConfig]:
    header_row = header_row or _find_header_row_by_alias_groups(
        sheet,
        CONFIG_HEADER_GROUPS,
    )
    headers = {
        cell.column: _normalize_text(cell.value)
        for cell in sheet[header_row]
        if cell.value is not None
    }
    col_label = _find_column(headers, ("Etapas", "Rango de etapas", "Stage range"))
    col_start = _find_column(headers, ("Inicio", "Etapa inicial", "Start stage", "From stage"))
    col_end = _find_column(headers, ("Fin", "Etapa final", "End stage", "To stage"))
    col_clusters = _find_column(
        headers,
        ("N° Cl", "N Cl", "Nro Cl", "Cantidad de clusters", "Clusters", "Cluster count"),
    )
    col_spf = _find_column(headers, ("SPF", "Tiros por pie", "Shots per foot"))

    configs: list[FractureConfig] = []
    blank_run = 0
    for row_idx in range(header_row + 1, sheet.max_row + 1):
        label = sheet.cell(row_idx, col_label).value
        start = sheet.cell(row_idx, col_start).value
        end = sheet.cell(row_idx, col_end).value

        if label in (None, "") and start in (None, "") and end in (None, ""):
            blank_run += 1
            if configs and blank_run >= 3:
                break
            continue
        blank_run = 0

        if not (_is_number(start) or isinstance(start, str)):
            continue
        if not (_is_number(end) or isinstance(end, str)):
            continue

        configs.append(
            FractureConfig(
                label=str(label).strip() if label not in (None, "") else f"Etapas {start}-{end}",
                start_stage=_as_int(start, "Inicio", row_idx),
                end_stage=_as_int(end, "Fin", row_idx),
                clusters=_as_int(sheet.cell(row_idx, col_clusters).value, "N° Cl", row_idx),
                spf=_as_int(sheet.cell(row_idx, col_spf).value, "SPF", row_idx),
            )
        )

    if not configs:
        raise InvalidWorkbookError(
            f"No se encontraron configuraciones de fractura en '{sheet.title}'."
        )
    return configs


def _extract_survey(
    sheet: Worksheet,
    header_row: int | None = None,
) -> list[SurveyPoint]:
    header_row = header_row or _find_header_row_by_alias_groups(
        sheet,
        SURVEY_HEADER_GROUPS,
        max_scan_rows=60,
    )
    headers = {
        cell.column: _normalize_text(cell.value)
        for cell in sheet[header_row]
        if cell.value is not None
    }
    col_md = _find_column(headers, ("MD", "Measured depth", "Prof [m]", "Profundidad medida"))
    col_tvd = _find_column(headers, ("TVD", "True vertical depth", "Pfv [m]", "Profundidad vertical"))
    col_incl = _find_column(
        headers,
        ("INCL", "Inclination", "Inclinacion", "Desviac. Vert", "Desviacion vertical"),
    )
    col_azim = _find_column(headers, ("AZIM_TN", "Azimuth", "Azimut", "AZIM"))
    col_dx = _find_column_or_none(
        headers,
        ("DX", "Dis EO", "Desplazamiento Este Oeste", "Easting", "East West"),
    )
    col_dy = _find_column_or_none(
        headers,
        ("DY", "Dis NS", "Desplazamiento Norte Sur", "Northing", "North South"),
    )

    survey: list[SurveyPoint] = []
    for row_idx in range(header_row + 1, sheet.max_row + 1):
        md = sheet.cell(row_idx, col_md).value
        md_number = _coerce_float(md)
        if md_number is None:
            continue
        survey.append(
            SurveyPoint(
                md=round(md_number, 2),
                inclination=round(_as_float(sheet.cell(row_idx, col_incl).value, "INCL", row_idx), 2),
                azimuth=round(_as_float(sheet.cell(row_idx, col_azim).value, "AZIM", row_idx), 2),
                tvd=round(_as_float(sheet.cell(row_idx, col_tvd).value, "TVD", row_idx), 2),
                dx=(
                    round(dx_number, 2)
                    if col_dx is not None
                    and (dx_number := _coerce_float(sheet.cell(row_idx, col_dx).value)) is not None
                    else None
                ),
                dy=(
                    round(dy_number, 2)
                    if col_dy is not None
                    and (dy_number := _coerce_float(sheet.cell(row_idx, col_dy).value)) is not None
                    else None
                ),
            )
        )

    if not survey:
        raise InvalidWorkbookError(
            f"No se encontraron registros numéricos en la hoja '{sheet.title}'."
        )
    return survey


def _survey_depth_pairs(sheet: Worksheet, header_row: int) -> set[tuple[float, float]]:
    """Identify a survey by its measured and vertical depths, even if angles are blank."""
    headers = {
        cell.column: _normalize_text(cell.value)
        for cell in sheet[header_row]
        if cell.value is not None
    }
    col_md = _find_column(headers, SURVEY_HEADER_GROUPS[0])
    col_tvd = _find_column(headers, SURVEY_HEADER_GROUPS[1])
    pairs = set()
    for row_idx in range(header_row + 1, sheet.max_row + 1):
        md = _coerce_float(sheet.cell(row_idx, col_md).value)
        tvd = _coerce_float(sheet.cell(row_idx, col_tvd).value)
        if md is not None and tvd is not None:
            pairs.add((round(md, 2), round(tvd, 2)))
    return pairs


def _matching_survey_source(
    workbook: Workbook,
    incomplete_sheet: Worksheet,
    incomplete_header_row: int,
) -> tuple[Worksheet, int, list[SurveyPoint]] | None:
    """Use another complete survey only when its depths match the incomplete table."""
    reference = _survey_depth_pairs(incomplete_sheet, incomplete_header_row)
    if len(reference) < 2:
        return None

    matches = []
    for sheet in workbook.worksheets:
        if sheet is incomplete_sheet:
            continue
        try:
            header_row = _find_header_row_by_alias_groups(
                sheet, SURVEY_HEADER_GROUPS, max_scan_rows=100
            )
            points = _extract_survey(sheet, header_row)
        except InvalidWorkbookError:
            continue
        candidate = {(point.md, point.tvd) for point in points}
        overlap = len(reference & candidate)
        if overlap >= 2 and overlap / len(reference) >= 0.95:
            matches.append((overlap, sheet, header_row, points))

    if not matches:
        return None
    matches.sort(key=lambda item: item[0], reverse=True)
    if len(matches) > 1 and matches[0][0] == matches[1][0]:
        names = ", ".join(f"'{item[1].title}'" for item in matches if item[0] == matches[0][0])
        raise InvalidWorkbookError(
            f"Hay varias hojas Survey completas que coinciden con '{incomplete_sheet.title}': {names}."
        )
    _, sheet, header_row, points = matches[0]
    return sheet, header_row, points


def _extract_clusters(
    sheet: Worksheet,
    header_row: int | None = None,
) -> list[Cluster]:
    header_row = header_row or _find_header_row_by_alias_groups(
        sheet,
        CLUSTER_HEADER_GROUPS,
        max_scan_rows=60,
    )
    headers = {
        cell.column: _normalize_text(cell.value)
        for cell in sheet[header_row]
        if cell.value is not None
    }
    col_number = _find_column(headers, ("# Cluster", "Numero de cluster", "Cluster number"))
    col_top = _find_column(headers, ("Tope Cluster MD", "Tope cluster", "Cluster top", "Top MD"))
    col_base = _find_column(headers, CLUSTER_BASE_HEADER_ALIASES)
    col_stage = _find_column(headers, ("Número etapa", "Numero etapa", "Etapa", "Stage number"))
    col_spf = _find_column(
        headers,
        ("N° de tiros x cluster", "tiros x cluster", "SPF", "Shots per foot"),
    )
    col_stage_length = _find_column_or_none(headers, ("Longitud de etapa", "Stage length"))
    col_expected = _find_column_or_none(headers, ("Cantidad de clusters", "Cluster count"))

    override_col: int | None = None
    for col_idx, header in headers.items():
        if "sobreescribir" in header or "cambio de punzados" in header:
            override_col = col_idx
            break

    clusters: list[Cluster] = []
    for row_idx in range(header_row + 1, sheet.max_row + 1):
        cluster_number = sheet.cell(row_idx, col_number).value
        stage = sheet.cell(row_idx, col_stage).value
        cluster_number_value = _coerce_float(cluster_number)
        stage_value = _coerce_float(stage)
        if cluster_number_value is None or stage_value is None:
            continue

        override_value = sheet.cell(row_idx, override_col).value if override_col else None
        expected_value = sheet.cell(row_idx, col_expected).value if col_expected else None
        stage_length_value = sheet.cell(row_idx, col_stage_length).value if col_stage_length else None
        clusters.append(
            Cluster(
                cluster_number=_as_int(cluster_number_value, "# Cluster", row_idx),
                top_md=_as_float(sheet.cell(row_idx, col_top).value, "Tope Cluster MD", row_idx),
                base_md=_as_float(sheet.cell(row_idx, col_base).value, "Base Cluster MD", row_idx),
                stage=_as_int(stage_value, "Número etapa", row_idx),
                spf=_as_int(sheet.cell(row_idx, col_spf).value, "N° de tiros x cluster", row_idx),
                expected_clusters=(
                    _as_int(expected_value, "Cantidad de clusters", row_idx)
                    if expected_value not in (None, "")
                    else None
                ),
                stage_length=(
                    _as_float(stage_length_value, "Longitud de etapa", row_idx)
                    if stage_length_value not in (None, "")
                    else None
                ),
                override_note=(str(override_value).strip() if override_value not in (None, "") else None),
            )
        )

    if not clusters:
        raise InvalidWorkbookError(
            f"No se encontraron clústeres numéricos en la hoja '{sheet.title}'."
        )
    return clusters


def _infer_fracture_configs(clusters: list[Cluster]) -> list[FractureConfig]:
    """Build the Datos Fractura ranges when no separate configuration table exists."""
    grouped: dict[int, list[Cluster]] = {}
    for cluster in clusters:
        grouped.setdefault(cluster.stage, []).append(cluster)

    signatures: list[tuple[int, int, int]] = []
    for stage in sorted(grouped):
        spfs = {cluster.spf for cluster in grouped[stage]}
        if len(spfs) != 1:
            raise InvalidWorkbookError(
                f"Etapa {stage}: no se puede inferir la configuración porque hay valores SPF inconsistentes."
            )
        signatures.append((stage, len(grouped[stage]), next(iter(spfs))))

    configs: list[FractureConfig] = []
    start_stage: int | None = None
    end_stage: int | None = None
    active_clusters: int | None = None
    active_spf: int | None = None

    for stage, cluster_count, spf in signatures:
        is_continuation = (
            end_stage is not None
            and stage == end_stage + 1
            and cluster_count == active_clusters
            and spf == active_spf
        )
        if start_stage is None or not is_continuation:
            if start_stage is not None:
                configs.append(
                    FractureConfig(
                        label=f"ET {start_stage}-{end_stage}",
                        start_stage=start_stage,
                        end_stage=end_stage or start_stage,
                        clusters=active_clusters or 0,
                        spf=active_spf or 0,
                    )
                )
            start_stage = stage
            active_clusters = cluster_count
            active_spf = spf
        end_stage = stage

    if start_stage is not None:
        configs.append(
            FractureConfig(
                label=f"ET {start_stage}-{end_stage}",
                start_stage=start_stage,
                end_stage=end_stage or start_stage,
                clusters=active_clusters or 0,
                spf=active_spf or 0,
            )
        )
    if not configs:
        raise InvalidWorkbookError(
            "No se pudo inferir la configuración de fractura desde los punzados."
        )
    return configs


def _reconcile_fracture_configs(
    configs: list[FractureConfig], clusters: list[Cluster]
) -> tuple[list[FractureConfig], list[str], list[str]]:
    """Trust Punzados counts only when its declared and actual counts agree."""
    grouped: dict[int, list[Cluster]] = {}
    for cluster in clusters:
        grouped.setdefault(cluster.stage, []).append(cluster)

    reconciled = []
    warnings = []
    decision_details = []
    for config in configs:
        counts = set()
        for stage in range(config.start_stage, config.end_stage + 1):
            stage_clusters = grouped.get(stage, [])
            declared = {cluster.expected_clusters for cluster in stage_clusters}
            if not stage_clusters or declared != {len(stage_clusters)}:
                break
            counts.add(len(stage_clusters))
        else:
            if len(counts) == 1:
                actual_count = counts.pop()
                if actual_count != config.clusters:
                    detail = (
                        f"Input, configuración '{config.label}': N° Cl indica "
                        f"{config.clusters}, pero Punzados confirma {actual_count} "
                        "clústeres por etapa."
                    )
                    decision_details.append(detail)
                    warnings.append(f"{detail} Se usó Punzados.")
                    config = replace(config, clusters=actual_count)
        reconciled.append(config)
    return reconciled, warnings, decision_details


def _build_stages(
    clusters: list[Cluster],
    configs: list[FractureConfig],
) -> tuple[list[StageInterval], list[str], list[str]]:
    grouped: dict[int, list[Cluster]] = {}
    for cluster in clusters:
        grouped.setdefault(cluster.stage, []).append(cluster)

    stages: list[StageInterval] = []
    warnings: list[str] = []
    checks: list[str] = []

    actual_stage_numbers = sorted(grouped)
    expected_stage_numbers = list(range(1, max(actual_stage_numbers) + 1))
    if actual_stage_numbers != expected_stage_numbers:
        missing = sorted(set(expected_stage_numbers) - set(actual_stage_numbers))
        raise InvalidWorkbookError(
            "La numeración de etapas no es consecutiva. "
            f"Etapas faltantes: {missing or 'no identificadas'}."
        )

    for cluster in clusters:
        if cluster.top_md >= cluster.base_md:
            raise InvalidWorkbookError(
                f"Clúster {cluster.cluster_number}, etapa {cluster.stage}: "
                f"el Tope ({cluster.top_md}) debe ser menor que el Fondo ({cluster.base_md})."
            )

    seen_cluster_numbers: set[int] = set()
    duplicates: list[int] = []
    for cluster in clusters:
        if cluster.cluster_number in seen_cluster_numbers:
            duplicates.append(cluster.cluster_number)
        seen_cluster_numbers.add(cluster.cluster_number)
    if duplicates:
        raise InvalidWorkbookError(
            f"Hay números de clúster duplicados: {sorted(set(duplicates))[:20]}."
        )

    for stage_number in actual_stage_numbers:
        stage_clusters = grouped[stage_number]
        top_md = min(cluster.top_md for cluster in stage_clusters)
        base_md = max(cluster.base_md for cluster in stage_clusters)
        expected_counts = {
            cluster.expected_clusters
            for cluster in stage_clusters
            if cluster.expected_clusters is not None
        }
        spfs = {cluster.spf for cluster in stage_clusters}
        stage_lengths = {
            round(cluster.stage_length, 6)
            for cluster in stage_clusters
            if cluster.stage_length is not None
        }

        if len(expected_counts) > 1:
            raise InvalidWorkbookError(
                f"Etapa {stage_number}: la cantidad esperada de clústeres no es consistente."
            )
        if expected_counts and len(stage_clusters) != next(iter(expected_counts)):
            raise InvalidWorkbookError(
                f"Etapa {stage_number}: se encontraron {len(stage_clusters)} clústeres "
                f"y se esperaban {next(iter(expected_counts))} según Punzados."
            )
        if len(spfs) != 1:
            raise InvalidWorkbookError(f"Etapa {stage_number}: hay valores SPF inconsistentes.")
        if len(stage_lengths) != 1:
            warnings.append(
                f"Etapa {stage_number}: aparecen longitudes de etapa diferentes dentro del mismo grupo."
            )
        if top_md >= base_md:
            raise InvalidWorkbookError(
                f"Etapa {stage_number}: el Tope ({top_md}) debe ser menor que el Fondo ({base_md})."
            )

        stages.append(
            StageInterval(
                stage=stage_number,
                top_md=round(top_md, 2),
                base_md=round(base_md, 2),
                plug_md=round(base_md + 3.7, 2),
            )
        )

    # Cross-check every configured range against the actual clusters and SPF.
    configured_stages: set[int] = set()
    for config in configs:
        if config.start_stage > config.end_stage:
            raise InvalidWorkbookError(
                f"Configuración '{config.label}': Inicio es mayor que Fin."
            )
        for stage_number in range(config.start_stage, config.end_stage + 1):
            if stage_number in configured_stages:
                raise InvalidWorkbookError(
                    f"La etapa {stage_number} aparece en más de una configuración de fractura."
                )
            configured_stages.add(stage_number)
            stage_clusters = grouped.get(stage_number)
            if not stage_clusters:
                raise InvalidWorkbookError(
                    f"La configuración '{config.label}' incluye la etapa {stage_number}, "
                    "pero no hay punzados para esa etapa."
                )
            if len(stage_clusters) != config.clusters:
                raise InvalidWorkbookError(
                    f"{config.label}, etapa {stage_number}: se encontraron "
                    f"{len(stage_clusters)} clústeres y la configuración indica {config.clusters}."
                )
            actual_spf = {cluster.spf for cluster in stage_clusters}
            if actual_spf != {config.spf}:
                raise InvalidWorkbookError(
                    f"{config.label}, etapa {stage_number}: SPF {sorted(actual_spf)} "
                    f"y la configuración indica {config.spf}."
                )

    unconfigured = sorted(set(actual_stage_numbers) - configured_stages)
    if unconfigured:
        raise InvalidWorkbookError(
            f"Hay etapas sin configuración de fractura: {unconfigured}."
        )

    overrides = [cluster for cluster in clusters if cluster.override_note]
    if overrides:
        warnings.append(
            f"Se detectaron {len(overrides)} celdas con observaciones o sobreescrituras "
            "en la hoja Punzados. Revisar el archivo original si el cambio modifica Tope o Fondo."
        )

    checks.extend(
        [
            f"{len(stages)} etapas consecutivas",
            f"{len(clusters)} clústeres procesados",
            "Cantidad de clústeres por etapa correcta",
            "SPF consistente con las configuraciones",
            f"{len(configs)} configuraciones de fractura",
        ]
    )
    return stages, warnings, checks


def analyze_workbook(
    file_obj: bytes | BinaryIO,
    filename: str,
    *,
    accept_recoveries: bool = False,
) -> ProcessingResult:
    stream = BytesIO(file_obj) if isinstance(file_obj, bytes) else file_obj
    try:
        workbook = load_workbook(
            stream,
            read_only=False,
            data_only=True,
            keep_vba=False,
            keep_links=False,
        )
    except Exception as exc:  # openpyxl exposes several parser exceptions
        raise InvalidWorkbookError(
            "No se pudo abrir el archivo. Verificá que sea un Excel .xlsm o .xlsx válido."
        ) from exc

    try:
        survey_location = _locate_table(
            workbook,
            "Survey",
            SURVEY_HEADER_GROUPS,
            max_scan_rows=100,
        )
        if survey_location is None:
            raise InvalidWorkbookError(
                'No se encontró la hoja "Survey" ni otra tabla con MD, INCL, AZIM y TVD.'
            )

        cluster_location = _locate_table(
            workbook,
            "Punzados",
            CLUSTER_HEADER_GROUPS,
            max_scan_rows=60,
        )
        if cluster_location is None:
            raise InvalidWorkbookError(
                'No se encontró la hoja "Punzados" ni otra tabla con etapa, clúster, Tope, Fondo y SPF.'
            )

        config_location = _locate_table(
            workbook,
            "Input",
            CONFIG_HEADER_GROUPS,
            max_scan_rows=80,
        )

        survey_sheet, survey_header_row = survey_location
        cluster_sheet, cluster_header_row = cluster_location
        clusters = _extract_clusters(cluster_sheet, cluster_header_row)

        inferred_configs = config_location is None
        config_warnings: list[str] = []
        decision_details: list[str] = []
        if config_location is None:
            configs = _infer_fracture_configs(clusters)
            well_name = Path(filename).stem.replace("_version2", "")
        else:
            config_sheet, config_header_row = config_location
            configs = _extract_fracture_configs(config_sheet, config_header_row)
            configs, config_warnings, config_decisions = _reconcile_fracture_configs(
                configs, clusters
            )
            decision_details.extend(config_decisions)
            well_name = _extract_well_name(config_sheet, filename)

        survey_source_warning = None
        try:
            survey = _extract_survey(survey_sheet, survey_header_row)
        except InvalidWorkbookError:
            matched_survey = _matching_survey_source(
                workbook, survey_sheet, survey_header_row
            )
            if matched_survey is None:
                raise
            original_sheet_name = survey_sheet.title
            survey_sheet, survey_header_row, survey = matched_survey
            survey_source_warning = (
                f"La hoja '{original_sheet_name}' tiene datos Survey incompletos; "
                f"se usó '{survey_sheet.title}' porque coinciden MD y TVD."
            )
            decision_details.append(
                f"La hoja '{original_sheet_name}' tiene datos Survey incompletos. "
                f"'{survey_sheet.title}' contiene {len(survey)} registros y coincide en MD y TVD."
            )
        stages, warnings, checks = _build_stages(clusters, configs)
        if decision_details and not accept_recoveries:
            raise ProcessingDecisionRequired(decision_details)
        warnings[:0] = config_warnings
        if survey_source_warning:
            warnings.insert(0, survey_source_warning)
        if inferred_configs:
            warnings.insert(
                0,
                "No se encontró una tabla de configuración separada; Datos Fractura se infirió "
                "desde la cantidad de clústeres y SPF de cada etapa.",
            )
            checks.append("Configuraciones de fractura inferidas desde Punzados")

        source_sheets = {
            "Survey": f"{survey_sheet.title} (fila {survey_header_row})",
            "Punzados": f"{cluster_sheet.title} (fila {cluster_header_row})",
        }
        if config_location is not None:
            source_sheets["Input"] = f"{config_sheet.title} (fila {config_header_row})"

        checks.extend(
            [
                f"{len(survey)} registros Survey",
                f"{len(stages) * 2} filas Wellbore IFS",
                "Tapones calculados como Fondo + 3,7 m",
                "Tablas de origen detectadas por encabezados",
            ]
        )

        return ProcessingResult(
            well_name=well_name,
            output_filename=_safe_output_filename(well_name),
            fracture_configs=configs,
            survey=survey,
            clusters=clusters,
            stages=stages,
            warnings=warnings,
            checks=checks,
            source_sheets=source_sheets,
        )
    finally:
        workbook.close()


def _check_generated_workbook(data: bytes, result: ProcessingResult) -> list[str]:
    try:
        workbook = load_workbook(BytesIO(data), data_only=False)
    except Exception as exc:
        raise LuctivError("El archivo generado no pudo volver a abrirse con openpyxl.") from exc

    if TEMPLATE_SHEET not in workbook.sheetnames:
        raise LuctivError(f"El archivo generado no contiene la hoja '{TEMPLATE_SHEET}'.")

    sheet = workbook[TEMPLATE_SHEET]
    errors: list[str] = []
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip() in FORMULA_ERROR_VALUES:
                errors.append(f"{cell.coordinate}: {cell.value.strip()}")
    if errors:
        raise LuctivError(
            "El archivo generado contiene errores visibles de Excel: " + ", ".join(errors[:20])
        )

    for row_idx, stage in enumerate(_smart_staging_stages(result), start=4):
        if sheet.cell(row_idx, 13).value != stage.stage:
            raise LuctivError(f"Smart Staging tiene una etapa incorrecta en la fila {row_idx}.")
        plug = sheet.cell(row_idx, 16).value
        fondo = sheet.cell(row_idx, 15).value
        if not isinstance(plug, (int, float)) or not isinstance(fondo, (int, float)):
            raise LuctivError(f"El tapón de la fila {row_idx} no quedó como valor numérico.")
        if not math.isclose(float(plug), float(fondo) + 3.7, abs_tol=0.001):
            raise LuctivError(
                f"El tapón de la etapa {stage.stage} no coincide con Fondo + 3,7 m."
            )

    expected_wellbore: list[tuple[str, float, float]] = []
    for stage in _wellbore_ifs_stages(result):
        expected_wellbore.append(("Treatment Interval", stage.top_md, stage.base_md))
        expected_wellbore.append(("Perforations", stage.top_md, stage.base_md))

    for offset, expected in enumerate(expected_wellbore):
        row_idx = 4 + offset
        label, top_md, base_md = expected
        if sheet.cell(row_idx, 18).value != label:
            raise LuctivError(f"Wellbore IFS no tiene la etiqueta esperada en la fila {row_idx}.")
        if not math.isclose(float(sheet.cell(row_idx, 19).value), top_md, abs_tol=0.001):
            raise LuctivError(f"Wellbore IFS tiene un Tope incorrecto en la fila {row_idx}.")
        if not math.isclose(float(sheet.cell(row_idx, 20).value), base_md, abs_tol=0.001):
            raise LuctivError(f"Wellbore IFS tiene un Fondo incorrecto en la fila {row_idx}.")

    variable_blocks = (
        ("Datos Fractura", 3 + len(result.fracture_configs), 1, 5),
        ("Datos Survey", 4 + len(result.survey), 8, 11),
        ("Smart Staging", 4 + len(result.stages), 13, 16),
        ("Wellbore IFS", 4 + result.wellbore_row_count, 18, 20),
    )
    for block_name, first_empty_row, min_col, max_col in variable_blocks:
        for row_idx in range(first_empty_row, sheet.max_row + 1):
            for col_idx in range(min_col, max_col + 1):
                value = sheet.cell(row_idx, col_idx).value
                if value not in (None, ""):
                    raise LuctivError(
                        f"Quedaron datos residuales en {block_name}: "
                        f"{sheet.cell(row_idx, col_idx).coordinate}."
                    )

    return [
        "Smart Staging ordenado por ETAPA de mayor a menor",
        "Wellbore IFS ordenado por MD de mayor a menor",
        "Wellbore IFS verificado con dos filas por etapa",
        "Archivo .xlsx verificado con openpyxl",
        "Sin datos residuales en los rangos variables",
        "Sin errores críticos",
    ]


def _copy_cell_style(source: Cell, target: Cell) -> None:
    if source.has_style:
        target._style = copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    if source.alignment:
        target.alignment = copy(source.alignment)
    if source.protection:
        target.protection = copy(source.protection)


def _copy_row_style(
    sheet: Worksheet,
    source_row: int,
    target_row: int,
    min_col: int,
    max_col: int,
) -> None:
    for col_idx in range(min_col, max_col + 1):
        _copy_cell_style(sheet.cell(source_row, col_idx), sheet.cell(target_row, col_idx))
    if source_row in sheet.row_dimensions:
        sheet.row_dimensions[target_row].height = sheet.row_dimensions[source_row].height


def _clear_values(sheet: Worksheet, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    for row in sheet.iter_rows(
        min_row=min_row,
        max_row=max_row,
        min_col=min_col,
        max_col=max_col,
    ):
        for cell in row:
            cell.value = None


def _write_output_value(
    sheet: Worksheet,
    row: int,
    column: int,
    value: object,
) -> None:
    """Write an output value without inheriting custom numeric formats."""
    cell = sheet.cell(row, column, value)
    if _is_number(value):
        cell.number_format = "General"


def _ensure_styles(sheet: Worksheet, result: ProcessingResult) -> None:
    design_end = 2 + len(result.fracture_configs)
    survey_end = 3 + len(result.survey)
    stage_end = 3 + len(result.stages)
    wellbore_end = 3 + result.wellbore_row_count

    # The template has representative styles in these rows. Copy only styling, never values.
    for row_idx in range(3, design_end + 1):
        source_row = 3 if row_idx == 3 else 4
        _copy_row_style(sheet, source_row, row_idx, 1, 5)

    for row_idx in range(4, survey_end + 1):
        _copy_row_style(sheet, 4, row_idx, 8, 11)

    for row_idx in range(4, stage_end + 1):
        _copy_row_style(sheet, 4, row_idx, 13, 16)

    for row_idx in range(4, wellbore_end + 1):
        source_row = 4 if (row_idx - 4) % 2 == 0 else 5
        _copy_row_style(sheet, source_row, row_idx, 18, 20)


def _set_number_format(
    sheet: Worksheet,
    min_row: int,
    max_row: int,
    columns: Iterable[int],
    number_format: str,
) -> None:
    if max_row < min_row:
        return
    for row_idx in range(min_row, max_row + 1):
        for col_idx in columns:
            sheet.cell(row_idx, col_idx).number_format = number_format


def _normalize_output_number_formats(sheet: Worksheet, result: ProcessingResult) -> None:
    _set_number_format(sheet, 3, 2 + len(result.fracture_configs), range(2, 6), "0")
    _set_number_format(sheet, 4, 3 + len(result.survey), range(8, 12), "General")
    _set_number_format(sheet, 4, 3 + len(result.stages), (13,), "0")
    _set_number_format(sheet, 4, 3 + len(result.stages), range(14, 17), "General")
    _set_number_format(sheet, 4, 3 + result.wellbore_row_count, range(19, 21), "General")


def generate_finished_workbook(
    result: ProcessingResult,
    template_path: str | Path,
) -> GeneratedWorkbook:
    template_path = Path(template_path)
    if not template_path.exists():
        raise LuctivError("No se encontró la plantilla de salida de LUCTIV.")

    workbook = load_workbook(template_path)
    if TEMPLATE_SHEET not in workbook.sheetnames:
        raise LuctivError(
            f"La plantilla no contiene la hoja '{TEMPLATE_SHEET}'."
        )
    sheet = workbook[TEMPLATE_SHEET]

    required_max_row = max(
        3 + len(result.survey),
        3 + len(result.stages),
        3 + result.wellbore_row_count,
        2 + len(result.fracture_configs),
        sheet.max_row,
    )
    _clear_values(sheet, 3, required_max_row, 1, 5)
    _clear_values(sheet, 4, required_max_row, 8, 11)
    _clear_values(sheet, 4, required_max_row, 13, 16)
    _clear_values(sheet, 4, required_max_row, 18, 20)
    _ensure_styles(sheet, result)

    for row_idx, config in enumerate(result.fracture_configs, start=3):
        values = (
            config.label,
            config.start_stage,
            config.end_stage,
            config.clusters,
            config.spf,
        )
        for col_idx, value in enumerate(values, start=1):
            _write_output_value(sheet, row_idx, col_idx, value)

    for row_idx, point in enumerate(result.survey, start=4):
        _write_output_value(sheet, row_idx, 8, point.md)
        _write_output_value(sheet, row_idx, 9, point.inclination)
        _write_output_value(sheet, row_idx, 10, point.azimuth)
        _write_output_value(sheet, row_idx, 11, point.tvd)

    for row_idx, stage in enumerate(_smart_staging_stages(result), start=4):
        _write_output_value(sheet, row_idx, 13, stage.stage)
        _write_output_value(sheet, row_idx, 14, stage.top_md)
        _write_output_value(sheet, row_idx, 15, stage.base_md)
        _write_output_value(sheet, row_idx, 16, stage.plug_md)

    wellbore_row = 4
    for stage in _wellbore_ifs_stages(result):
        for label in ("Treatment Interval", "Perforations"):
            sheet.cell(wellbore_row, 18, label)
            _write_output_value(sheet, wellbore_row, 19, stage.top_md)
            _write_output_value(sheet, wellbore_row, 20, stage.base_md)
            wellbore_row += 1

    _normalize_output_number_formats(sheet, result)

    # Ensure Excel recalculates any formulas that could exist in the template.
    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
    except AttributeError:
        pass

    output = BytesIO()
    workbook.save(output)
    data = output.getvalue()
    for check in _check_generated_workbook(data, result):
        if check not in result.checks:
            result.checks.append(check)
    txt_check = "Survey TXT/CSV generado con MD, INCLINATION y AZIMUTH"
    if txt_check not in result.checks:
        result.checks.append(txt_check)
    return GeneratedWorkbook(
        result=result,
        data=data,
        survey_txt_data=_survey_delimited_data(result, "\t"),
        survey_csv_data=_survey_delimited_data(result, ","),
    )


def process_uploaded_workbook(
    file_bytes: bytes,
    filename: str,
    template_path: str | Path,
    *,
    accept_recoveries: bool = False,
) -> GeneratedWorkbook:
    if Path(filename).suffix.lower() not in ALLOWED_EXTENSIONS:
        raise InvalidWorkbookError("LUCTIV solo acepta archivos Excel .xlsm o .xlsx.")
    result = analyze_workbook(
        file_bytes, filename, accept_recoveries=accept_recoveries
    )
    return generate_finished_workbook(result, template_path)


def _safe_package_filename(filename: str) -> str:
    stem = Path(filename).stem
    cleaned = re.sub(r"[^A-Za-z0-9._()\-]+", "-", stem).strip("-_.")
    return f"{cleaned or 'PAD'}_terminados.zip"


def _package_member_filename(member_name: str) -> str:
    normalized = member_name.replace("\\", "/")
    return PurePosixPath(normalized).name


def _unique_output_filename(filename: str, used_names: set[str]) -> str:
    candidate = filename
    stem = Path(filename).stem
    suffix = Path(filename).suffix
    counter = 2
    while candidate.lower() in used_names:
        candidate = f"{stem}-{counter}{suffix}"
        counter += 1
    used_names.add(candidate.lower())
    return candidate


def _build_package_summary(
    workbooks: Iterable[PackageWorkbook],
    issues: Iterable[PackageIssue],
) -> str:
    workbook_items = tuple(workbooks)
    issue_items = tuple(issues)
    lines = [
        "LUCTIV - Resumen de procesamiento del PAD",
        f"Archivos generados: {len(workbook_items)}",
        f"Archivos con error: {sum(1 for issue in issue_items if issue.category == 'error')}",
        f"Archivos auxiliares ignorados: {sum(1 for issue in issue_items if issue.category == 'ignored')}",
        "",
        "GENERADOS",
    ]
    for item in workbook_items:
        result = item.generated.result
        lines.append(
            f"- {item.source_filename} -> {result.output_filename} | "
            f"pozo {result.well_name} | {len(result.stages)} etapas | "
            f"{len(result.clusters)} clusters | {len(result.survey)} registros Survey"
        )
    failed_items = tuple(issue for issue in issue_items if issue.category == "error")
    ignored_items = tuple(issue for issue in issue_items if issue.category == "ignored")
    if failed_items:
        lines.extend(("", "NO PROCESADOS / REQUIEREN REVISION"))
        for issue in failed_items:
            lines.append(f"- {issue.source_filename}: {issue.message}")
    if ignored_items:
        lines.extend(("", "ARCHIVOS AUXILIARES IGNORADOS"))
        for issue in ignored_items:
            lines.append(f"- {issue.source_filename}: {issue.message}")
    return "\n".join(lines) + "\n"


def _looks_like_single_well_workbook(file_bytes: bytes) -> bool:
    """Differentiate an intended well workbook from auxiliary spreadsheets in a PAD ZIP."""
    try:
        workbook = load_workbook(
            BytesIO(file_bytes),
            read_only=False,
            data_only=True,
            keep_vba=False,
            keep_links=False,
        )
    except Exception:
        return True
    try:
        survey_location = _locate_table(
            workbook,
            "Survey",
            SURVEY_HEADER_GROUPS,
            max_scan_rows=100,
        )
        cluster_location = _locate_table(
            workbook,
            "Punzados",
            CLUSTER_HEADER_GROUPS,
            max_scan_rows=60,
        )
        return survey_location is not None and cluster_location is not None
    except InvalidWorkbookError:
        return False
    finally:
        workbook.close()


def process_uploaded_package(
    file_bytes: bytes,
    filename: str,
    template_path: str | Path,
    *,
    accept_recoveries: bool = False,
) -> GeneratedWorkbook | GeneratedPackage:
    """Process either one workbook or every compatible workbook inside a ZIP package."""
    suffix = Path(filename).suffix.lower()
    if suffix in ALLOWED_EXTENSIONS:
        return process_uploaded_workbook(
            file_bytes, filename, template_path, accept_recoveries=accept_recoveries
        )
    if suffix != ".zip":
        raise InvalidWorkbookError(
            "LUCTIV solo acepta archivos Excel .xlsm/.xlsx o un paquete .zip."
        )

    try:
        source_archive = ZipFile(BytesIO(file_bytes))
    except BadZipFile as exc:
        raise InvalidWorkbookError("El archivo .zip está dañado o no es un ZIP válido.") from exc

    workbooks: list[PackageWorkbook] = []
    issues: list[PackageIssue] = []
    decision_details: list[str] = []
    try:
        excel_entries = [
            info
            for info in source_archive.infolist()
            if not info.is_dir()
            and Path(_package_member_filename(info.filename)).suffix.lower() in ALLOWED_EXTENSIONS
        ]
        if not excel_entries:
            raise InvalidWorkbookError(
                "El paquete no contiene archivos Excel .xlsm o .xlsx."
            )
        if len(excel_entries) > MAX_PACKAGE_EXCEL_FILES:
            raise InvalidWorkbookError(
                f"El paquete contiene {len(excel_entries)} Excel; el máximo permitido es "
                f"{MAX_PACKAGE_EXCEL_FILES}."
            )
        total_uncompressed = sum(info.file_size for info in excel_entries)
        if total_uncompressed > MAX_PACKAGE_TOTAL_BYTES:
            raise InvalidWorkbookError(
                "El contenido Excel descomprimido del paquete supera el máximo permitido."
            )

        for info in excel_entries:
            member_filename = _package_member_filename(info.filename) or "archivo.xlsx"
            if info.flag_bits & 0x1:
                issues.append(
                    PackageIssue(member_filename, "El archivo está cifrado con contraseña.")
                )
                continue
            if info.file_size > MAX_PACKAGE_ENTRY_BYTES:
                issues.append(
                    PackageIssue(
                        member_filename,
                        "El archivo supera el tamaño máximo individual permitido.",
                    )
                )
                continue
            contents = b""
            try:
                contents = source_archive.read(info)
                generated = process_uploaded_workbook(
                    contents,
                    member_filename,
                    template_path,
                    accept_recoveries=accept_recoveries,
                )
            except ProcessingDecisionRequired as exc:
                decision_details.extend(
                    f"{member_filename}: {detail}" for detail in exc.details
                )
                continue
            except (InvalidWorkbookError, LuctivError, BadZipFile, RuntimeError) as exc:
                category = (
                    "error"
                    if not contents or _looks_like_single_well_workbook(contents)
                    else "ignored"
                )
                issues.append(PackageIssue(member_filename, str(exc), category))
                continue
            workbooks.append(PackageWorkbook(member_filename, generated))
    finally:
        source_archive.close()

    if decision_details:
        raise ProcessingDecisionRequired(decision_details)

    if not workbooks:
        details = "; ".join(
            f"{issue.source_filename}: {issue.message}" for issue in issues[:3]
        )
        raise InvalidWorkbookError(
            "No se pudo generar ningún Excel terminado desde el paquete. " + details
        )

    output = BytesIO()
    used_names: set[str] = set()
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as output_archive:
        packaged_workbooks: list[PackageWorkbook] = []
        for item in workbooks:
            survey_txt_name = item.generated.survey_txt_filename
            survey_csv_name = item.generated.survey_csv_filename
            output_name = _unique_output_filename(
                item.generated.result.output_filename,
                used_names,
            )
            if output_name != item.generated.result.output_filename:
                item.generated.result.output_filename = output_name
            output_archive.writestr(output_name, item.generated.data)
            output_archive.writestr(
                _unique_output_filename(survey_txt_name, used_names),
                item.generated.survey_txt_data,
            )
            output_archive.writestr(
                _unique_output_filename(survey_csv_name, used_names),
                item.generated.survey_csv_data,
            )
            packaged_workbooks.append(item)
        output_archive.writestr(
            "RESUMEN_PROCESAMIENTO.txt",
            _build_package_summary(packaged_workbooks, issues),
        )

    return GeneratedPackage(
        output_filename=_safe_package_filename(filename),
        data=output.getvalue(),
        workbooks=tuple(workbooks),
        issues=tuple(issues),
    )
