from __future__ import annotations

import streamlit as st

from processors import (
    UNKNOWN_VENDOR,
    InputRecord,
    available_vendors,
    build_record,
    combine_records,
    dataframe_to_excel_bytes,
    read_clipboard_table,
    read_excel_upload,
)


st.set_page_config(page_title="식품이력등록 데이터 취합", page_icon="📄", layout="wide")


def read_inputs() -> list[InputRecord]:
    records: list[InputRecord] = []

    pasted_text = st.text_area(
        "엑셀 붙여넣기",
        height=260,
        placeholder="엑셀에서 필요한 범위를 복사한 뒤 여기에 붙여넣으세요.",
        help="엑셀 복사 데이터는 탭으로 구분된 표 형태로 인식합니다.",
    )
    pasted_name = st.text_input("붙여넣기 원본명", value="클립보드 입력")

    if pasted_text.strip():
        try:
            dataframe = read_clipboard_table(pasted_text)
            if not dataframe.empty:
                records.append(build_record("클립보드", pasted_name, dataframe))
        except Exception as exc:
            st.error(f"붙여넣기 데이터를 읽지 못했습니다: {exc}")

    with st.expander("엑셀 파일 업로드", expanded=False):
        uploaded_files = st.file_uploader(
            "엑셀 파일 업로드",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            help="여러 파일을 한 번에 선택하거나 드래그 앤 드롭할 수 있습니다.",
        )

        if uploaded_files:
            for uploaded_file in uploaded_files:
                try:
                    dataframe = read_excel_upload(uploaded_file)
                    records.append(build_record("파일업로드", uploaded_file.name, dataframe))
                except Exception as exc:
                    st.error(f"{uploaded_file.name} 파일을 읽지 못했습니다: {exc}")

    return records


def vendor_review(records: list[InputRecord]) -> list[InputRecord]:
    vendors = available_vendors()
    reviewed_records: list[InputRecord] = []

    st.subheader("거래처 인식 결과")
    for index, record in enumerate(records, start=1):
        expanded = record.detected_vendor == UNKNOWN_VENDOR
        with st.expander(f"{index}. {record.source_name}", expanded=expanded):
            st.write(f"입력출처: {record.source_type}")
            st.write(f"자동 인식 거래처: {record.detected_vendor}")
            if record.first_row_text:
                st.caption(f"첫 행: {record.first_row_text}")

            default_index = vendors.index(record.detected_vendor) if record.detected_vendor in vendors else 0
            selected_vendor = st.selectbox(
                "거래처 선택",
                vendors,
                index=default_index,
                key=f"vendor_{index}_{record.source_type}_{record.source_name}",
            )

            reviewed_records.append(
                InputRecord(
                    source_type=record.source_type,
                    source_name=record.source_name,
                    dataframe=record.dataframe,
                    first_row_text=record.first_row_text,
                    detected_vendor=record.detected_vendor,
                    selected_vendor=selected_vendor,
                )
            )

    return reviewed_records


def main() -> None:
    st.title("식품이력등록 데이터 취합")
    st.caption("엑셀 데이터를 붙여넣거나 파일을 업로드한 뒤 산출물만 다운로드합니다. 서버에는 파일을 저장하지 않습니다.")

    records = read_inputs()

    if not records:
        st.info("엑셀 데이터를 붙여넣거나 엑셀 파일을 업로드하세요.")
        return

    reviewed_records = vendor_review(records)

    if st.button("가공 실행", type="primary"):
        combined = combine_records(reviewed_records)
        st.session_state["processed_dataframe"] = combined
        st.session_state["processed_excel"] = dataframe_to_excel_bytes(combined)

    processed_dataframe = st.session_state.get("processed_dataframe")
    processed_excel = st.session_state.get("processed_excel")
    if processed_dataframe is not None and processed_excel is not None:
        st.success("가공 완료! 다운로드하세요")
        st.dataframe(processed_dataframe, use_container_width=True)
        st.download_button(
            "최종 엑셀 파일 다운로드",
            data=processed_excel,
            file_name="식품이력등록_취합결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
