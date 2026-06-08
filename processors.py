from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO, StringIO
from typing import Callable

import pandas as pd


UNKNOWN_VENDOR = "미인식"


VENDOR_ALIASES: dict[str, list[str]] = {
    # 추후 거래처별 규칙 입력 시 여기에 파일명/첫 행 키워드를 추가합니다.
    # 예: "sample_vendor": ["샘플거래처", "sample"],
}


@dataclass
class InputRecord:
    source_type: str
    source_name: str
    dataframe: pd.DataFrame
    first_row_text: str
    detected_vendor: str
    selected_vendor: str


Processor = Callable[[pd.DataFrame], pd.DataFrame]


def _pass_through(dataframe: pd.DataFrame) -> pd.DataFrame:
    return dataframe.copy()


PROCESSOR_REGISTRY: dict[str, Processor] = {
    UNKNOWN_VENDOR: _pass_through,
}


def available_vendors() -> list[str]:
    vendors = sorted(set(VENDOR_ALIASES) | set(PROCESSOR_REGISTRY))
    if UNKNOWN_VENDOR not in vendors:
        vendors.insert(0, UNKNOWN_VENDOR)
    return vendors


def first_row_as_text(dataframe: pd.DataFrame) -> str:
    if dataframe.empty:
        return ""
    first_row = dataframe.iloc[0].fillna("").astype(str).tolist()
    return " ".join(value.strip() for value in first_row if value.strip())


def detect_vendor(source_name: str, dataframe: pd.DataFrame) -> str:
    haystack = f"{source_name} {first_row_as_text(dataframe)}".casefold()
    for vendor_id, aliases in VENDOR_ALIASES.items():
        if any(alias.casefold() in haystack for alias in aliases):
            return vendor_id
    return UNKNOWN_VENDOR


def read_excel_upload(uploaded_file) -> pd.DataFrame:
    return pd.read_excel(BytesIO(uploaded_file.getvalue()), dtype=object)


def read_clipboard_table(raw_text: str) -> pd.DataFrame:
    cleaned = raw_text.strip()
    if not cleaned:
        return pd.DataFrame()
    return pd.read_csv(StringIO(cleaned), sep="\t", dtype=object)


def build_record(source_type: str, source_name: str, dataframe: pd.DataFrame) -> InputRecord:
    detected_vendor = detect_vendor(source_name, dataframe)
    return InputRecord(
        source_type=source_type,
        source_name=source_name,
        dataframe=dataframe,
        first_row_text=first_row_as_text(dataframe),
        detected_vendor=detected_vendor,
        selected_vendor=detected_vendor,
    )


def process_record(record: InputRecord) -> pd.DataFrame:
    processor = PROCESSOR_REGISTRY.get(record.selected_vendor, _pass_through)
    processed = processor(record.dataframe)
    processed.insert(0, "거래처", record.selected_vendor)
    processed.insert(1, "입력출처", record.source_type)
    processed.insert(2, "원본명", record.source_name)
    return processed


def combine_records(records: list[InputRecord]) -> pd.DataFrame:
    processed_frames = [process_record(record) for record in records]
    if not processed_frames:
        return pd.DataFrame()
    return pd.concat(processed_frames, ignore_index=True)


def dataframe_to_excel_bytes(dataframe: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="취합결과")
    buffer.seek(0)
    return buffer.getvalue()
