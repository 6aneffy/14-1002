from __future__ import annotations

from datetime import date as _date, datetime

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from src.yeobi.exporter import build_xlsx, bundle_receipts_pdf
from src.yeobi.extractor import extract_from_file
from src.yeobi.form_exporter import (
    FormMeta,
    SettlementChoice,
    VehicleChoice,
    build_settlement_workbook,
    format_trip_period_kr,
    sheet_title_for_travel_date,
    yeobi_sheet_total,
)
from src.yeobi.schema import (
    TRANSPORT_OPTIONS,
    TRAVEL_CLASS_SUGGESTIONS,
    Receipt,
    TransportType,
)
from src.yeobi.thumbnail import make_view_data_url

load_dotenv()

# 로그인 사용자(데모): 출장부서(C10)·소속(B26)에 동일 문자열, 직급·성명만 별도
PERSON_PROFILES: dict[str, dict[str, str]] = {
    "임영주": {
        "name": "임영주",
        "dept_affiliation": "국고실 조달계약정책관",
        "position": "주무관",
    },
    "염철민": {
        "name": "염철민",
        "dept_affiliation": "인공지능경제과 전략경제정책관",
        "position": "과장",
    },
    "김만수": {
        "name": "김만수",
        "dept_affiliation": "세제실 재산세제과",
        "position": "과장",
    },
}


def _per_diem_and_meal(vehicle: str | None) -> tuple[int, int]:
    """공무용차량 미이용 시 일비 25,000원, 이용 시 12,500원. 식비는 항상 25,000원."""
    meal = 25_000
    if vehicle == "이용":
        return 12_500, meal
    return 25_000, meal


def _date_suffix(d: _date | None) -> str:
    return d.isoformat() if d is not None else "__NONE__"


def _receipts_travel_date_sort(receipts: list[Receipt]) -> list[Receipt]:
    """운임일자 오름차순(4/1 → 4/30). 일자 없음은 맨 뒤, 같은 일자는 기존 행 순서 유지."""
    return [
        r
        for _, r in sorted(
            enumerate(receipts),
            key=lambda ir: (
                ir[1].travel_date is None,
                ir[1].travel_date or _date.min,
                ir[0],
            ),
        )
    ]


def _distinct_travel_dates(receipts: list[Receipt]) -> list[_date | None]:
    """운임일자별 탭/시트 — 날짜 오름차순(과거→미래). 일자 없음은 맨 뒤 한 시트."""
    known = sorted({r.travel_date for r in receipts if r.travel_date is not None})
    if any(r.travel_date is None for r in receipts):
        known.append(None)
    return known


def _init_per_date_form_keys(dates: list[_date | None]) -> None:
    for d in dates:
        suf = _date_suffix(d)
        st.session_state.setdefault(f"f_trip_type_{suf}", "")
        st.session_state.setdefault(f"f_purpose_{suf}", "")
        st.session_state.setdefault(f"f_grade_{suf}", "제1호")


def _filename_month_part(receipts: list[Receipt]) -> str:
    """파일명용: 영수증 운임일자에 나온 연·월 기준. 한 달이면 '4월', 여러 달이면 '3·4월' 등."""
    ym = sorted({(r.travel_date.year, r.travel_date.month) for r in receipts if r.travel_date})
    if not ym:
        return f"{datetime.now().month}월"
    if len(ym) == 1:
        return f"{ym[0][1]}월"
    years = {y for y, _ in ym}
    if len(years) == 1:
        months = [m for _, m in ym]
        return "·".join(str(m) for m in months) + "월"
    return "·".join(f"{y}년{m}월" for y, m in ym)


st.set_page_config(page_title="일사천리 (ILSACHUNLI)", page_icon="🚅", layout="wide")

st.title("🚅 일사천리 앱")
st.caption("일사천리는 출장 영수증을 업로드하면 자동으로 여비정산신청서를 생성해주는 서비스입니다.")


def _init_state() -> None:
    st.session_state.setdefault("receipts", [])
    st.session_state.setdefault("uploaded_files", {})
    st.session_state.setdefault("file_urls", {})
    st.session_state.setdefault("selected_person", "임영주")


_init_state()


# ──────────────────────── Sidebar ────────────────────────
with st.sidebar:
    st.subheader("👤 담당자")
    st.selectbox(
        "성명 (로그인 사용자)",
        options=list(PERSON_PROFILES.keys()),
        key="selected_person",
        help="선택한 사람의 출장부서·소속·직급이 여비정산신청서에 자동 반영됩니다.",
    )
    st.divider()
    st.header("📎 영수증 업로드")
    uploads = st.file_uploader(
        "사진(jpg/png) 또는 PDF 여러 개 선택",
        type=["jpg", "jpeg", "png", "webp", "pdf"],
        accept_multiple_files=True,
    )
    if uploads:
        new_files = [u for u in uploads if u.name not in st.session_state.uploaded_files]
        if new_files and st.button(f"📤 {len(new_files)}개 영수증 분석", type="primary"):
            progress = st.progress(0.0, text="분석 준비중...")
            for idx, up in enumerate(new_files, start=1):
                progress.progress(
                    (idx - 1) / len(new_files), text=f"분석중: {up.name}"
                )
                data = up.getvalue()
                try:
                    extracted = extract_from_file(up.name, data)
                    st.session_state.receipts.extend(extracted)
                    st.session_state.uploaded_files[up.name] = data
                    try:
                        st.session_state.file_urls[up.name] = (
                            make_view_data_url(up.name, data)
                        )
                    except Exception:
                        st.session_state.file_urls[up.name] = ""
                except Exception as e:
                    st.error(f"{up.name} 분석 실패: {e}")
            progress.progress(1.0, text="완료")
            st.rerun()

    if st.session_state.uploaded_files:
        st.divider()
        st.caption(f"업로드됨: {len(st.session_state.uploaded_files)}개")
        for name in st.session_state.uploaded_files:
            st.write(f"· {name}")
        if st.button("🗑️ 전체 초기화"):
            st.session_state.receipts = []
            st.session_state.uploaded_files = {}
            st.session_state.file_urls = {}
            st.rerun()


# ──────────────────────── Main: results ────────────────────────
receipts: list[Receipt] = st.session_state.receipts

st.subheader("📊 영수증 분석 결과")

if not receipts:
    st.info("왼쪽 사이드바에서 영수증을 업로드하고 **분석** 버튼을 누르세요.")
else:
    receipts = _receipts_travel_date_sort(list(receipts))
    st.session_state.receipts = receipts
    rows = [
        {
            "No.": i,
            "운임일자": r.travel_date.isoformat() if r.travel_date else "",
            "교통편": r.transport.value,
            "등급": r.travel_class or "",
            "매수": r.quantity,
            "금액": r.amount,
            "출발지": r.origin or "",
            "도착지": r.destination or "",
            "원본": r.source_file or "",
        }
        for i, r in enumerate(receipts, start=1)
    ]
    df = pd.DataFrame(rows)

    edited = st.data_editor(
        df,
        column_config={
            "No.": st.column_config.NumberColumn(width="small", disabled=True),
            "운임일자": st.column_config.TextColumn(
                help="YYYY-MM-DD", width="small"
            ),
            "교통편": st.column_config.SelectboxColumn(
                options=TRANSPORT_OPTIONS, required=True, width="medium"
            ),
            "등급": st.column_config.TextColumn(
                help=f"예: {', '.join(TRAVEL_CLASS_SUGGESTIONS[:4])} 등",
                width="medium",
            ),
            "매수": st.column_config.NumberColumn(min_value=1, step=1, width="small"),
            "금액": st.column_config.NumberColumn(
                format="%d 원", min_value=0, width="small"
            ),
            "출발지": st.column_config.TextColumn(width="small"),
            "도착지": st.column_config.TextColumn(width="small"),
            "원본": st.column_config.TextColumn(
                "원본", width="medium", disabled=True
            ),
        },
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        key="receipt_editor",
    )

    # Sync edits back to session_state
    updated: list[Receipt] = []
    for i, row in edited.iterrows():
        original = receipts[i]
        try:
            td = _date.fromisoformat(row["운임일자"]) if row["운임일자"] else None
        except Exception:
            td = original.travel_date
        try:
            transport = TransportType(row["교통편"])
        except ValueError:
            transport = original.transport
        try:
            updated.append(
                Receipt(
                    transport=transport,
                    travel_class=(row["등급"] or None) if row["등급"] != "" else None,
                    quantity=int(row["매수"] or 1),
                    travel_date=td,
                    amount=int(row["금액"] or 0),
                    origin=row["출발지"] or None,
                    destination=row["도착지"] or None,
                    source_file=original.source_file,
                )
            )
        except Exception:
            updated.append(original)
    st.session_state.receipts = _receipts_travel_date_sort(updated)
    receipts = st.session_state.receipts

    total = sum(r.amount for r in receipts)
    count = sum(r.quantity for r in receipts)

    m1, m2, m3 = st.columns(3)
    m1.metric("건수", f"{len(receipts)} 건")
    m2.metric("매수 합계", f"{count} 매")
    m3.metric("금액 합계", f"{total:,} 원")

    if st.session_state.file_urls:
        with st.expander("📎 원본 영수증 이미지 보기"):
            names = [
                n for n in st.session_state.file_urls if st.session_state.file_urls[n]
            ]
            if names:
                picked = st.selectbox("파일 선택", names, key="view_pick")
                st.image(st.session_state.file_urls[picked], caption=picked)

    st.divider()

    # ──────────────── 여비정산신청서 폼 입력 ────────────────
    trip_dates = _distinct_travel_dates(receipts)
    _init_per_date_form_keys(trip_dates)

    prof = PERSON_PROFILES.get(
        st.session_state.get("selected_person", "임영주"),
        PERSON_PROFILES["임영주"],
    )
    da = prof["dept_affiliation"]
    display_name = prof["name"]
    written = _date.today()

    st.subheader("📝 여비정산신청서 정보")
    with st.expander("출장·정산 정보 입력 (펼쳐서 편집)", expanded=False):
        st.caption(
            "탭은 표의 **운임일자**마다 자동 생성됩니다. "
            "각 일자별로 여비 등급·출장 종류·공무용 차량·정산 유형·출장 목적을 입력하세요. "
            "출장 기간·예산 과목은 양식에서 생략합니다."
        )
        tab_labels = [
            format_trip_period_kr(d) if d is not None else "일자 미상"
            for d in trip_dates
        ]
        for d, label, tab in zip(trip_dates, tab_labels, st.tabs(tab_labels)):
            suf = _date_suffix(d)
            with tab:
                st.markdown(f"**{label}** 일자용 입력")
                st.radio(
                    "여비 등급",
                    options=["제1호", "제2호"],
                    index=0,
                    horizontal=True,
                    key=f"f_grade_{suf}",
                )
                st.text_input(
                    "출장 종류",
                    key=f"f_trip_type_{suf}",
                    placeholder="예: 회의",
                )
                c_a, c_b = st.columns(2)
                with c_a:
                    st.radio(
                        "공무용 차량",
                        options=["이용안함", "이용"],
                        index=0,
                        horizontal=True,
                        key=f"f_vehicle_{suf}",
                    )
                with c_b:
                    st.radio(
                        "정산 유형",
                        options=["숙박비실비", "숙박비정액"],
                        index=1,
                        horizontal=True,
                        key=f"f_settlement_{suf}",
                    )
                st.text_area(
                    "출장 목적",
                    key=f"f_purpose_{suf}",
                    height=70,
                )
                _veh_tab = st.session_state.get(f"f_vehicle_{suf}", "이용안함")
                _pd_tab, _ml_tab = _per_diem_and_meal(_veh_tab)
                st.caption(
                    f"이 일자 시트의 일비·식비: 일비 {_pd_tab:,}원 · 식비 {_ml_tab:,}원 "
                    f"(해당 일자 영수증만 운임 합계 반영)"
                )

    entries: list[tuple[str, FormMeta, list[Receipt]]] = []
    for d in trip_dates:
        suf = _date_suffix(d)
        veh_raw = st.session_state.get(f"f_vehicle_{suf}", "이용안함")
        veh: VehicleChoice = (
            veh_raw if veh_raw in ("이용안함", "이용") else "이용안함"  # type: ignore[assignment]
        )
        set_raw = st.session_state.get(f"f_settlement_{suf}", "숙박비정액")
        settlement: SettlementChoice = (
            set_raw if set_raw in ("숙박비실비", "숙박비정액") else "숙박비정액"
        )
        per_diem_cash, meal_cash = _per_diem_and_meal(veh)
        period_str = format_trip_period_kr(d) if d is not None else ""
        g_raw = st.session_state.get(f"f_grade_{suf}", "제1호")
        travel_grade = g_raw if g_raw in ("제1호", "제2호") else "제1호"
        recs_for_day = [
            r
            for r in receipts
            if (r.travel_date == d if d is not None else r.travel_date is None)
        ]
        meta = FormMeta(
            written_date=written,
            department=da,
            trip_period=period_str,
            trip_type=st.session_state.get(f"f_trip_type_{suf}", ""),
            budget_item="",
            remarks="",
            purpose=st.session_state.get(f"f_purpose_{suf}", ""),
            vehicle=veh,
            settlement=settlement,
            traveler_name=display_name,
            affiliation=da,
            position=prof["position"],
            travel_grade=travel_grade,
            individual_period=period_str,
            per_diem_cash=per_diem_cash,
            meal_cash=meal_cash,
            lodging_fixed_cash=0,
            lodging_actual_cash=0,
            prepaid_cash=0,
        )
        entries.append((sheet_title_for_travel_date(d), meta, recs_for_day))

    form_bytes = build_settlement_workbook(
        entries, st.session_state.get("uploaded_files") or {}
    )
    _month_seg = _filename_month_part(receipts)
    settlement_filename = f"{display_name} {_month_seg} 여비정산신청서.xlsx"
    yeobi_grand_total = sum(
        yeobi_sheet_total(meta, recs) for _, meta, recs in entries
    )

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    d1, d2, d3 = st.columns(3)
    with d1:
        xlsx_bytes = build_xlsx(receipts, st.session_state.uploaded_files)
        st.download_button(
            "⬇️ 정산신청 리스트 (xlsx)",
            data=xlsx_bytes,
            file_name=f"정산신청리스트_{stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with d2:
        st.download_button(
            "⬇️ 여비정산신청서 (xlsx)",
            data=form_bytes,
            file_name=settlement_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
    with d3:
        if st.session_state.uploaded_files:
            pdf_bytes = bundle_receipts_pdf(
                list(st.session_state.uploaded_files.items())
            )
            st.download_button(
                "⬇️ 증빙 영수증 묶음 (PDF)",
                data=pdf_bytes,
                file_name=f"영수증증빙_{stamp}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

    breakdown_lines: list[str] = []
    for _, meta, recs in entries:
        day_lbl = (meta.trip_period or "").strip() or "일자 미상"
        amt = yeobi_sheet_total(meta, recs)
        breakdown_lines.append(f"- **{day_lbl}** : **{amt:,}**원")
    breakdown_lines.append(f"- **총합**: **{yeobi_grand_total:,}**원")
    st.markdown("**여비계**")
    st.markdown("\n".join(breakdown_lines))
