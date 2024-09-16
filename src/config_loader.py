from __future__ import annotations

from pathlib import Path

import yaml
from pydantic import BaseModel, Field, ValidationError


class Column(BaseModel):
    name: str = Field(..., max_length=255)
    md_pattern: str | None = Field(None, max_length=255)
    length: int = Field(..., ge=0)
    horizontal: str = Field(..., pattern="center|left|right")
    vertical: str = Field(..., pattern="center|top|bottom")
    multi_idx: bool | None = Field(None, alias="multi-idx")


class Columns(BaseModel):
    section: Column = Field(...)
    subsection: Column = Field(...)
    testcase: Column = Field(...)
    number: Column = Field(...)
    pos_neg: Column = Field(..., alias="pos-neg")
    result: Column = Field(...)
    step: Column = Field(...)
    expectation: Column = Field(...)
    notes: Column = Field(...)


class SheetName(BaseModel):
    summary: str = Field(..., max_length=255)
    test: str = Field(..., max_length=255)


class ExcelSettings(BaseModel):
    font_name: str = Field(..., max_length=255)
    sheet_name: SheetName = Field(...)


class Config(BaseModel):
    columns: Columns = Field(...)
    excel_settings: ExcelSettings = Field(...)


def load_config(file_path: Path):
    if not file_path.exists():
        raise FileNotFoundError(f"設定ファイルが見つかりません: {file_path}")
    with open(file_path, 'r', encoding='utf-8_sig') as f:
        config_data = yaml.safe_load(f)
        try:
            return Config(**config_data)
        except ValidationError as e:
            raise ValueError(f"設定ファイルの形式が正しくありません\n{e}")


def load_column_names(config: Config) -> list[str]:
    return [col["name"] for col in config.columns.model_dump().values()]
