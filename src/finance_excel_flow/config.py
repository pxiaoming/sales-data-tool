from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


@dataclass(slots=True)
class InputConfig:
    sheet_name: str | int | None = 0
    header_row: int = 1


@dataclass(slots=True)
class OutputConfig:
    sheet_name: str = "结果"
    filename_template: str = "{source_stem}_processed_{yyyymmdd}.xlsx"
    template_path: str | None = None
    cell_values: dict[str, Any] = field(default_factory=dict)
    freeze_header: bool = True
    auto_fit_columns: bool = True


@dataclass(slots=True)
class TransformConfig:
    rename_columns: dict[str, str] = field(default_factory=dict)
    required_columns: list[str] = field(default_factory=list)
    drop_empty_rows: bool = True
    filters: list[str] = field(default_factory=list)
    calculated_columns: dict[str, str] = field(default_factory=dict)
    select_columns: list[str] = field(default_factory=list)
    sort_by: list[str] = field(default_factory=list)


@dataclass(slots=True)
class BusinessRuleConfig:
    name: str
    source_filters: list[str] = field(default_factory=list)
    group_by: list[str] = field(default_factory=list)
    sum_field: str = "Due Amount"
    percentage: float = 0.0
    group_percentages: dict[str, float] = field(default_factory=dict)
    name_template: str = "{source_yyyymm}"
    group_row_pairs: list[tuple[int, int]] = field(default_factory=list)
    template_cells: dict[str, str] = field(default_factory=dict)
    summary_template_cells: dict[str, str] = field(default_factory=dict)


@dataclass(slots=True)
class AppConfig:
    input: InputConfig = field(default_factory=InputConfig)
    output: OutputConfig = field(default_factory=OutputConfig)
    transform: TransformConfig = field(default_factory=TransformConfig)
    business_rules: list[BusinessRuleConfig] = field(default_factory=list)


def _as_mapping(value: Any, section_name: str) -> dict[str, Any]:
    if value is None:
        return {}
    if not isinstance(value, dict):
        raise ValueError(f"{section_name} must be a mapping")
    return value


def _as_list(value: Any, section_name: str) -> list[Any]:
    if value is None:
        return []
    if not isinstance(value, list):
        raise ValueError(f"{section_name} must be a list")
    return value


def _as_row_pairs(value: Any, section_name: str) -> list[tuple[int, int]]:
    items = _as_list(value, section_name)
    pairs: list[tuple[int, int]] = []
    for index, item in enumerate(items, start=1):
        if not isinstance(item, list) or len(item) != 2:
            raise ValueError(f"{section_name}[{index}] must be a list with exactly 2 items")
        left, right = item
        pairs.append((int(left), int(right)))
    return pairs


def load_config(path: str | Path) -> AppConfig:
    config_path = Path(path)
    raw = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
    if not isinstance(raw, dict):
        raise ValueError("Configuration root must be a mapping")

    input_raw = _as_mapping(raw.get("input"), "input")
    output_raw = _as_mapping(raw.get("output"), "output")
    transform_raw = _as_mapping(raw.get("transform"), "transform")
    business_rules_raw = _as_list(raw.get("business_rules"), "business_rules")

    input_cfg = InputConfig(
        sheet_name=input_raw.get("sheet_name", 0),
        header_row=int(input_raw.get("header_row", 1)),
    )
    if input_cfg.header_row < 1:
        raise ValueError("input.header_row must be >= 1")
    output_cfg = OutputConfig(
        sheet_name=str(output_raw.get("sheet_name", "结果")),
        filename_template=str(
            output_raw.get("filename_template", "{source_stem}_processed_{yyyymmdd}.xlsx")
        ),
        template_path=(
            str(output_raw["template_path"]).strip()
            if output_raw.get("template_path")
            else None
        ),
        cell_values=_as_mapping(output_raw.get("cell_values"), "output.cell_values"),
        freeze_header=bool(output_raw.get("freeze_header", True)),
        auto_fit_columns=bool(output_raw.get("auto_fit_columns", True)),
    )
    transform_cfg = TransformConfig(
        rename_columns=_as_mapping(transform_raw.get("rename_columns"), "transform.rename_columns"),
        required_columns=[str(item) for item in _as_list(transform_raw.get("required_columns"), "transform.required_columns")],
        drop_empty_rows=bool(transform_raw.get("drop_empty_rows", True)),
        filters=[str(item) for item in _as_list(transform_raw.get("filters"), "transform.filters")],
        calculated_columns=_as_mapping(transform_raw.get("calculated_columns"), "transform.calculated_columns"),
        select_columns=[str(item) for item in _as_list(transform_raw.get("select_columns"), "transform.select_columns")],
        sort_by=[str(item) for item in _as_list(transform_raw.get("sort_by"), "transform.sort_by")],
    )

    business_rules: list[BusinessRuleConfig] = []
    for index, item in enumerate(business_rules_raw, start=1):
        if not isinstance(item, dict):
            raise ValueError(f"business_rules[{index}] must be a mapping")
        rule_name = str(item.get("name", f"rule_{index}"))
        group_percentages_raw = _as_mapping(
            item.get("group_percentages"),
            f"business_rules[{index}].group_percentages",
        )
        business_rules.append(
            BusinessRuleConfig(
                name=rule_name,
                source_filters=[str(expr) for expr in _as_list(item.get("source_filters"), f"business_rules[{index}].source_filters")],
                group_by=[str(field) for field in _as_list(item.get("group_by"), f"business_rules[{index}].group_by")],
                sum_field=str(item.get("sum_field", "Due Amount")),
                percentage=float(item.get("percentage", 0.0)),
                group_percentages={str(key): float(value) for key, value in group_percentages_raw.items()},
                name_template=str(item.get("name_template", "{source_yyyymm}")),
                group_row_pairs=_as_row_pairs(item.get("group_row_pairs"), f"business_rules[{index}].group_row_pairs"),
                template_cells={
                    str(cell_ref): str(expr)
                    for cell_ref, expr in _as_mapping(item.get("template_cells"), f"business_rules[{index}].template_cells").items()
                },
                summary_template_cells={
                    str(cell_ref): str(expr)
                    for cell_ref, expr in _as_mapping(
                        item.get("summary_template_cells"),
                        f"business_rules[{index}].summary_template_cells",
                    ).items()
                },
            )
        )

    return AppConfig(input=input_cfg, output=output_cfg, transform=transform_cfg, business_rules=business_rules)
