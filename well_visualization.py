from __future__ import annotations

from bisect import bisect_left
from dataclasses import dataclass
import math
from typing import Sequence

from processor import ProcessingResult, SurveyPoint


@dataclass(frozen=True)
class TrajectoryPoint:
    md: float
    x: float
    y: float
    tvd: float
    inclination: float
    azimuth: float


@dataclass(frozen=True)
class ImpactMetrics:
    logical_rows: int
    values_prepared: int
    checks_passed: int
    elapsed_seconds: float
    estimated_minutes_saved: float
    estimated_efficiency_percent: float


def calculate_impact_metrics(
    result: ProcessingResult,
    elapsed_seconds: float,
    manual_minutes: float,
) -> ImpactMetrics:
    logical_rows = (
        len(result.fracture_configs)
        + len(result.survey)
        + len(result.stages)
        + result.wellbore_row_count
    )
    values_prepared = (
        len(result.fracture_configs) * 5
        + len(result.survey) * 4
        + len(result.stages) * 4
        + result.wellbore_row_count * 3
    )
    manual_seconds = max(float(manual_minutes), 0.0) * 60
    saved_seconds = max(manual_seconds - max(elapsed_seconds, 0.0), 0.0)
    efficiency = (saved_seconds / manual_seconds * 100) if manual_seconds else 0.0
    return ImpactMetrics(
        logical_rows=logical_rows,
        values_prepared=values_prepared,
        checks_passed=len(result.checks),
        elapsed_seconds=max(elapsed_seconds, 0.0),
        estimated_minutes_saved=saved_seconds / 60,
        estimated_efficiency_percent=efficiency,
    )


def _minimum_curvature_trajectory(survey: Sequence[SurveyPoint]) -> list[TrajectoryPoint]:
    if not survey:
        return []

    trajectory = [
        TrajectoryPoint(
            md=survey[0].md,
            x=0.0,
            y=0.0,
            tvd=survey[0].tvd,
            inclination=survey[0].inclination,
            azimuth=survey[0].azimuth,
        )
    ]
    x = 0.0
    y = 0.0

    for previous, current in zip(survey, survey[1:]):
        delta_md = current.md - previous.md
        if delta_md < 0:
            raise ValueError("El Survey debe estar ordenado por MD para construir la trayectoria 3D.")

        inc_1 = math.radians(previous.inclination)
        inc_2 = math.radians(current.inclination)
        azi_1 = math.radians(previous.azimuth)
        azi_2 = math.radians(current.azimuth)
        cosine_dogleg = (
            math.cos(inc_1) * math.cos(inc_2)
            + math.sin(inc_1) * math.sin(inc_2) * math.cos(azi_2 - azi_1)
        )
        dogleg = math.acos(max(-1.0, min(1.0, cosine_dogleg)))
        ratio_factor = 1.0 if dogleg < 1e-10 else 2 * math.tan(dogleg / 2) / dogleg

        x += (
            delta_md
            / 2
            * (math.sin(inc_1) * math.sin(azi_1) + math.sin(inc_2) * math.sin(azi_2))
            * ratio_factor
        )
        y += (
            delta_md
            / 2
            * (math.sin(inc_1) * math.cos(azi_1) + math.sin(inc_2) * math.cos(azi_2))
            * ratio_factor
        )
        trajectory.append(
            TrajectoryPoint(
                md=current.md,
                x=x,
                y=y,
                tvd=current.tvd,
                inclination=current.inclination,
                azimuth=current.azimuth,
            )
        )
    return trajectory


def build_trajectory(survey: Sequence[SurveyPoint]) -> tuple[list[TrajectoryPoint], str]:
    if not survey:
        return [], "reconstructed"
    if all(point.dx is not None and point.dy is not None for point in survey):
        return (
            [
                TrajectoryPoint(
                    md=point.md,
                    x=float(point.dx),
                    y=float(point.dy),
                    tvd=point.tvd,
                    inclination=point.inclination,
                    azimuth=point.azimuth,
                )
                for point in survey
            ],
            "survey",
        )
    return _minimum_curvature_trajectory(survey), "reconstructed"


def interpolate_trajectory(
    trajectory: Sequence[TrajectoryPoint],
    md: float,
) -> TrajectoryPoint:
    if not trajectory:
        raise ValueError("No hay puntos de Survey para interpolar.")
    if md <= trajectory[0].md:
        return trajectory[0]
    if md >= trajectory[-1].md:
        return trajectory[-1]

    measured_depths = [point.md for point in trajectory]
    upper_index = bisect_left(measured_depths, md)
    lower = trajectory[upper_index - 1]
    upper = trajectory[upper_index]
    span = upper.md - lower.md
    fraction = 0.0 if span == 0 else (md - lower.md) / span

    def lerp(start: float, end: float) -> float:
        return start + (end - start) * fraction

    azimuth_delta = (upper.azimuth - lower.azimuth + 180) % 360 - 180
    interpolated_azimuth = (lower.azimuth + azimuth_delta * fraction) % 360

    return TrajectoryPoint(
        md=md,
        x=lerp(lower.x, upper.x),
        y=lerp(lower.y, upper.y),
        tvd=lerp(lower.tvd, upper.tvd),
        inclination=lerp(lower.inclination, upper.inclination),
        azimuth=interpolated_azimuth,
    )


def _trajectory_slice(
    trajectory: Sequence[TrajectoryPoint],
    start_md: float,
    end_md: float,
) -> list[TrajectoryPoint]:
    start, end = sorted((start_md, end_md))
    return [
        interpolate_trajectory(trajectory, start),
        *(point for point in trajectory if start < point.md < end),
        interpolate_trajectory(trajectory, end),
    ]


def build_well_figure(
    result: ProcessingResult,
    *,
    height: int = 485,
) -> tuple[go.Figure, str]:
    import plotly.graph_objects as go

    trajectory, source = build_trajectory(result.survey)
    if not trajectory:
        raise ValueError("No hay Survey disponible para visualizar la trayectoria 3D.")

    figure = go.Figure()
    hover = [
        [point.md, point.tvd, point.inclination, point.azimuth]
        for point in trajectory
    ]
    figure.add_trace(
        go.Scatter3d(
            x=[point.x for point in trajectory],
            y=[point.y for point in trajectory],
            z=[point.tvd for point in trajectory],
            mode="lines",
            name="Trayectoria Survey",
            line={"color": "#6B7E8C", "width": 5},
            customdata=hover,
            hovertemplate=(
                "<b>Survey</b><br>MD %{customdata[0]:,.1f} m"
                "<br>TVD %{customdata[1]:,.1f} m"
                "<br>INCL %{customdata[2]:.2f}°"
                "<br>AZIM %{customdata[3]:.2f}°<extra></extra>"
            ),
        )
    )

    if result.stages:
        stage_start = min(stage.top_md for stage in result.stages)
        stage_end = max(stage.base_md for stage in result.stages)
        stimulated = _trajectory_slice(trajectory, stage_start, stage_end)
        figure.add_trace(
            go.Scatter3d(
                x=[point.x for point in stimulated],
                y=[point.y for point in stimulated],
                z=[point.tvd for point in stimulated],
                mode="lines",
                name="Intervalo estimulado",
                line={"color": "#55BED2", "width": 10},
                hoverinfo="skip",
            )
        )

        stage_points = [
            interpolate_trajectory(trajectory, (stage.top_md + stage.base_md) / 2)
            for stage in result.stages
        ]
        stage_customdata = [
            [stage.stage, stage.top_md, stage.base_md, point.tvd]
            for stage, point in zip(result.stages, stage_points)
        ]
        figure.add_trace(
            go.Scatter3d(
                x=[point.x for point in stage_points],
                y=[point.y for point in stage_points],
                z=[point.tvd for point in stage_points],
                mode="markers",
                name="Etapas",
                marker={
                    "size": 5,
                    "color": [stage.stage for stage in result.stages],
                    "colorscale": "Turbo",
                    "showscale": True,
                    "colorbar": {
                        "title": "Etapa",
                        "thickness": 10,
                        "len": 0.65,
                    },
                    "line": {"color": "#FFFFFF", "width": 0.7},
                },
                customdata=stage_customdata,
                hovertemplate=(
                    "<b>Etapa %{customdata[0]}</b>"
                    "<br>Tope %{customdata[1]:,.1f} m"
                    "<br>Fondo %{customdata[2]:,.1f} m"
                    "<br>TVD aprox. %{customdata[3]:,.1f} m<extra></extra>"
                ),
            )
        )

    figure.add_trace(
        go.Scatter3d(
            x=[trajectory[0].x],
            y=[trajectory[0].y],
            z=[trajectory[0].tvd],
            mode="markers+text",
            name="Superficie",
            text=["Superficie"],
            textposition="top center",
            marker={"size": 7, "color": "#1677A3", "symbol": "diamond"},
            hovertemplate="<b>Superficie</b><extra></extra>",
        )
    )

    figure.update_layout(
        height=height,
        margin={"l": 0, "r": 0, "t": 44, "b": 0},
        title={
            "text": "Trayectoria 3D · pozo completo",
            "x": 0.01,
            "xanchor": "left",
            "font": {"size": 16},
        },
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        legend={
            "orientation": "h",
            "yanchor": "bottom",
            "y": 1.01,
            "x": 0,
            "font": {"size": 10},
        },
        scene={
            "xaxis": {"title": "Desplazamiento X (m)", "showbackground": False},
            "yaxis": {"title": "Desplazamiento Y (m)", "showbackground": False},
            "zaxis": {
                "title": "Profundidad TVD (m)",
                "autorange": "reversed",
                "showbackground": False,
            },
            "aspectmode": "data",
            "camera": {"eye": {"x": 1.45, "y": 1.45, "z": 0.85}},
        },
        hoverlabel={"namelength": -1},
        uirevision=result.well_name,
    )
    return figure, source
