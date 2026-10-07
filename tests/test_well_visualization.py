from __future__ import annotations

import pytest

from processor import (
    FractureConfig,
    ProcessingResult,
    StageInterval,
    SurveyPoint,
)
from well_visualization import (
    TrajectoryPoint,
    build_trajectory,
    build_well_figure,
    calculate_impact_metrics,
    interpolate_trajectory,
)


def _result() -> ProcessingResult:
    return ProcessingResult(
        well_name="LajE-TEST(h)",
        output_filename="LajE-TEST(h)_terminado.xlsx",
        fracture_configs=[FractureConfig("ET 1-2", 1, 2, 12, 2)],
        survey=[
            SurveyPoint(0, 0, 0, 0, dx=0, dy=0),
            SurveyPoint(100, 45, 0, 85, dx=0, dy=35),
            SurveyPoint(200, 90, 0, 135, dx=0, dy=120),
        ],
        clusters=[],
        stages=[
            StageInterval(1, 140, 175, 178.7),
            StageInterval(2, 100, 135, 138.7),
        ],
        checks=["control 1", "control 2"],
    )


def test_build_trajectory_prefers_survey_displacements() -> None:
    trajectory, source = build_trajectory(_result().survey)

    assert source == "survey"
    assert trajectory[-1].x == 0
    assert trajectory[-1].y == 120
    assert trajectory[-1].tvd == 135


def test_build_trajectory_reconstructs_missing_displacements() -> None:
    survey = [
        SurveyPoint(0, 0, 0, 0),
        SurveyPoint(100, 0, 0, 100),
    ]

    trajectory, source = build_trajectory(survey)

    assert source == "reconstructed"
    assert trajectory[-1].x == pytest.approx(0)
    assert trajectory[-1].y == pytest.approx(0)
    assert trajectory[-1].tvd == 100


def test_interpolate_trajectory_uses_measured_depth() -> None:
    trajectory, _ = build_trajectory(_result().survey)

    midpoint = interpolate_trajectory(trajectory, 150)

    assert midpoint.x == pytest.approx(0)
    assert midpoint.y == pytest.approx(77.5)
    assert midpoint.tvd == pytest.approx(110)


def test_interpolate_trajectory_wraps_azimuth_through_north() -> None:
    trajectory = [
        TrajectoryPoint(0, 0, 0, 0, 90, 359),
        TrajectoryPoint(100, 0, 100, 0, 90, 1),
    ]

    midpoint = interpolate_trajectory(trajectory, 50)

    assert midpoint.azimuth == pytest.approx(0)


def test_impact_metrics_count_generated_work() -> None:
    metrics = calculate_impact_metrics(_result(), elapsed_seconds=2, manual_minutes=10)

    assert metrics.logical_rows == 10
    assert metrics.values_prepared == 37
    assert metrics.checks_passed == 2
    assert metrics.estimated_minutes_saved == pytest.approx(9.9666667)
    assert metrics.estimated_efficiency_percent == pytest.approx(99.6666667)


def test_well_figure_contains_path_interval_stages_and_surface() -> None:
    figure, source = build_well_figure(_result())

    assert source == "survey"
    assert len(figure.data) == 4
    stages_trace = next(trace for trace in figure.data if trace.name == "Etapas")
    assert [row[0] for row in stages_trace.customdata] == [1, 2]
    assert figure.layout.scene.zaxis.autorange == "reversed"
    assert figure.layout.height == 485
