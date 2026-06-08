from __future__ import annotations

import pandas as pd
import streamlit as st

from processors import (
    UNKNOWN_VENDOR,
    InputRecord,
    OUTPUT_COLUMNS,
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

    st.subheader("데이터 입력")
    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("**📋 이지어드민 출고 붙여넣기**")
        ezadmin_text = st.text_area(
            "이지어드민",
            height=220,
            placeholder="이지어드민 출고 내역을 엑셀에서 복사 후 붙여넣기",
            label_visibility="collapsed",
            key="ezadmin_paste",
        )
        if ezadmin_text.strip():
            try:
                df = read_clipboard_table(ezadmin_text)
                if not df.empty:
                    records.append(build_record("붙여넣기", "이지어드민", df))
                    st.success(f"이지어드민: {len(df)}행 인식")
            except Exception as exc:
                st.error(f"이지어드민 읽기 실패: {exc}")

    with col_right:
        st.markdown("**📋 롯데마트 출고 붙여넣기**")
        lotte_text = st.text_area(
            "롯데마트",
            height=220,
            placeholder="롯데마트 출고 내역을 엑셀에서 복사 후 붙여넣기",
            label_visibility="collapsed",
            key="lotte_paste",
        )
        if lotte_text.strip():
            try:
                df = read_clipboard_table(lotte_text)
                if not df.empty:
                    records.append(build_record("붙여넣기", "롯데마트", df))
                    st.success(f"롯데마트: {len(df)}행 인식")
            except Exception as exc:
                st.error(f"롯데마트 읽기 실패: {exc}")

    col3, col4 = st.columns(2)

    with col3:
        st.markdown("**📋 명현유통 출고 붙여넣기**")
        myunghyun_text = st.text_area(
            "명현유통",
            height=120,
            placeholder="명현유통 출고 내역 (없으면 비워두세요)",
            label_visibility="collapsed",
            key="myunghyun_paste",
        )
        if myunghyun_text.strip():
            try:
                df = read_clipboard_table(myunghyun_text)
                if not df.empty:
                    records.append(build_record("붙여넣기", "명현유통", df))
                    st.success(f"명현유통: {len(df)}행 인식")
            except Exception as exc:
                st.error(f"명현유통 읽기 실패: {exc}")

    with col4:
        st.markdown("**📋 본에프디 출고 붙여넣기**")
        bonfd_text = st.text_area(
            "본에프디",
            height=120,
            placeholder="본에프디 출고 내역 (없으면 비워두세요)",
            label_visibility="collapsed",
            key="bonfd_paste",
        )
        if bonfd_text.strip():
            try:
                df = read_clipboard_table(bonfd_text)
                if not df.empty:
                    records.append(build_record("붙여넣기", "본에프디", df))
                    st.success(f"본에프디: {len(df)}행 인식")
            except Exception as exc:
                st.error(f"본에프디 읽기 실패: {exc}")

    with st.expander("엑셀 파일 직접 업로드", expanded=False):
        uploaded_files = st.file_uploader(
            "엑셀 파일",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
        )
        if uploaded_files:
            for f in uploaded_files:
                try:
                    df = read_excel_upload(f)
                    records.append(build_record("파일업로드", f.name, df))
                    st.success(f"{f.name}: {len(df)}행 인식")
                except Exception as exc:
                    st.error(f"{f.name} 읽기 실패: {exc}")

    return records


def vendor_review(records: list[InputRecord]) -> list[InputRecord]:
    vendors = available_vendors()
    reviewed: list[InputRecord] = []

    st.subheader("거래처 인식 확인")
    for idx, record in enumerate(records, 1):
        needs_check = record.detected_vendor == UNKNOWN_VENDOR
        with st.expander(
            f"{idx}. {record.source_name}  ({len(record.dataframe)}행)  "
            + ("⚠️ 거래처 미인식" if needs_check else f"→ {record.detected_vendor}"),
            expanded=needs_check,
        ):
            if record.first_row_text:
                st.caption(f"첫 행 내용: {record.first_row_text[:120]}")
            st.caption(f"컬럼: {list(record.dataframe.columns)}")

            default_idx = vendors.index(record.detected_vendor) if record.detected_vendor in vendors else 0
            selected = st.selectbox(
                "거래처",
                vendors,
                index=default_idx,
                key=f"vendor_{idx}_{record.source_name}",
            )
            reviewed.append(
                InputRecord(
                    source_type=record.source_type,
                    source_name=record.source_name,
                    dataframe=record.dataframe,
                    first_row_text=record.first_row_text,
                    detected_vendor=record.detected_vendor,
                    selected_vendor=selected,
                )
            )

    return reviewed


def show_result_summary(df) -> None:
    if df.empty:
        return

    total = int(df["출고수량"].apply(pd.to_numeric, errors="coerce").sum())
    unrecognized = df[df["전송여부"].str.startswith("⚠", na=False)]

    col1, col2, col3 = st.columns(3)
    col1.metric("총 행 수", len(df))
    col2.metric("출고수량 합계", f"{total:,}")
    col3.metric("미인식 행", len(unrecognized), delta_color="inverse")

    if not unrecognized.empty:
        st.warning("인식 못 한 상품이 있습니다. 제품 코드 맵을 업데이트하세요.")
        st.dataframe(unrecognized, use_container_width=True)


def main() -> None:
    st.title("📄 식품이력등록 출고 데이터 취합")
    st.caption("원본 데이터는 서버에 저장하지 않습니다.")

    records = read_inputs()

    if not records:
        st.info("각 거래처 데이터를 붙여넣거나 파일을 업로드하세요.")
        return

    st.divider()
    reviewed_records = vendor_review(records)

    st.divider()
    if st.button("▶ 가공 실행", type="primary", use_container_width=True):
        try:
            combined = combine_records(reviewed_records)
            st.session_state["processed_df"] = combined
            st.session_state["processed_excel"] = dataframe_to_excel_bytes(combined)
        except Exception as exc:
            st.error(f"가공 중 오류: {exc}")

    df = st.session_state.get("processed_df")
    excel = st.session_state.get("processed_excel")

    if df is not None and excel is not None:
        st.subheader("가공 결과")
        show_result_summary(df)
        st.dataframe(df, use_container_width=True, height=400)
        st.download_button(
            "⬇ 엑셀 다운로드 (이력사이트 등록용)",
            data=excel,
            file_name="출고이력등록.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
