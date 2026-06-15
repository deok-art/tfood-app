from __future__ import annotations

import pandas as pd
import streamlit as st

from processors import (
    NEEDS_REVIEW_CHANNEL,
    UNKNOWN_VENDOR,
    InputRecord,
    available_vendors,
    build_export_filename,
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

    st.markdown("**📋 이지어드민 출고 붙여넣기**")
    ezadmin_text = st.text_area(
        "이지어드민",
        height=260,
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

    with st.expander("💡 이지어드민 데이터 가공 규칙"):
        st.markdown("""
        - **출고처 분류 (메모 기준)**: `쿠팡출하`→쿠팡, `컬리출하`→컬리, `네이버출하`→네이버, `ezpos`/`송장출력`/`cs 배송처리`→개인 (규칙 외 `❓확인필요`)
        - **제외 대상 (작업)**: `배송` 또는 `출고`가 아닌 건 (예: 입고)
        - **제외 대상 (상품)**: 상품명에 `직납` 또는 `영양밥`이 포함된 건
        - **수량 집계**: 소비기한, 상품명, 출고처가 동일한 여러 건은 하나의 행으로 수량 합산
        """)

    # 롯데마트 / 명현유통 / 본에프디 붙여넣기는 추후 구현 — 현재 숨김

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


def _qty_col(df: pd.DataFrame) -> str | None:
    for name in ("작업수량", "출고수량"):
        if name in df.columns:
            return name
    return None


def show_result(df: pd.DataFrame, excel: bytes) -> None:
    if df.empty:
        st.info("가공된 데이터가 없습니다.")
        return

    qty_col = _qty_col(df)
    total = int(df[qty_col].apply(pd.to_numeric, errors="coerce").sum()) if qty_col else 0

    col1, col2 = st.columns(2)
    col1.metric("총 행 수", len(df))
    col2.metric("수량 합계", f"{total:,}")

    # 휴먼터치 필요 행 경고
    if "출고처" in df.columns:
        review_rows = df[df["출고처"] == NEEDS_REVIEW_CHANNEL]
        if not review_rows.empty:
            st.warning(
                f"⚠️ 출고처 미분류 {len(review_rows)}건 (메모 패턴 미등록). "
                f"엑셀에서 빨강 셀 확인 후 규칙 추가 필요."
            )

    # 합계 행을 마지막에 붙여서 표시
    display_df = df.copy()
    if qty_col:
        total_row: dict = {c: "" for c in display_df.columns}
        total_row[qty_col] = total
        if "작업일자" in total_row:
            total_row["작업일자"] = "합계"
        elif "출고일" in total_row:
            total_row["출고일"] = "합계"
        display_df = pd.concat(
            [display_df, pd.DataFrame([total_row])], ignore_index=True
        )

    st.dataframe(display_df, use_container_width=True, height=420)
    st.download_button(
        "⬇ 엑셀 다운로드",
        data=excel,
        file_name=build_export_filename(df),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


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
        show_result(df, excel)


if __name__ == "__main__":
    main()
