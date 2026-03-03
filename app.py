"""
app.py — 월별 활동계획서 자동입력 Streamlit 앱
달력형 계획서(.xls) → 이용자별 엑셀 활동계획서(.xlsx) 자동 입력
"""

import time

import streamlit as st
from core import parse_calendar, detect_users, count_available_rows, fill_sheets

# ── 페이지 설정 ────────────────────────────────────────────
st.set_page_config(
    page_title="월별 활동계획서 자동입력",
    page_icon="📅",
    layout="wide",
)


# ══════════════════════════════════════════════════════════
# 스플래시 로딩 화면
# ══════════════════════════════════════════════════════════
def _show_splash():
    st.markdown("""
    <style>
    [data-testid="stHeader"],
    [data-testid="stToolbar"],
    [data-testid="stDecoration"],
    [data-testid="stStatusWidget"] { display: none !important; }
    .splash-page {
      position: fixed; inset: 0;
      background: #ffffff;
      display: flex; flex-direction: column;
      align-items: center; justify-content: center;
      gap: 24px; z-index: 99999;
    }
    .splash-title {
      font-size: 2.2rem; font-weight: 700; color: #333;
    }
    .splash-sub {
      font-size: 1.1rem; color: #888; margin-top: -12px;
    }
    .splash-dots { display: flex; gap: 10px; }
    .dot {
      display: inline-block;
      width: 13px; height: 13px;
      border-radius: 50%;
      background: #4a90d9;
      animation: dot-wave 1.4s ease-in-out infinite;
    }
    .dot:nth-child(1) { animation-delay: 0s;    }
    .dot:nth-child(2) { animation-delay: 0.22s; }
    .dot:nth-child(3) { animation-delay: 0.44s; }
    @keyframes dot-wave {
      0%,60%,100% { transform: translateY(0);    opacity: 0.35; }
      30%          { transform: translateY(-12px); opacity: 1;    }
    }
    </style>
    <div class="splash-page">
      <div class="splash-title">월별 활동계획서 자동입력</div>
      <div class="splash-sub">성인발달장애인 주간활동센터</div>
      <div class="splash-dots">
        <span class="dot"></span>
        <span class="dot"></span>
        <span class="dot"></span>
      </div>
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# 메인 앱 화면
# ══════════════════════════════════════════════════════════
def main_app():
    # ── 스플래시 (최초 1회) ─────────────────────────────────
    if "splash_shown" not in st.session_state:
        st.session_state["splash_shown"] = False

    if not st.session_state["splash_shown"]:
        _show_splash()
        time.sleep(3)
        st.session_state["splash_shown"] = True
        st.rerun()
        return

    # ── 헤더 ──────────────────────────────────────────────
    st.title("월별 활동계획서 자동입력")
    st.caption("성인발달장애인 주간활동센터 — 달력 계획서 → 이용자별 활동계획서 자동 입력 도구")

    # ── 파일 업로드 영역 한글화 ──────────────────────────────
    st.markdown("""
    <style>
    [data-testid="stFileUploaderDropzoneInstructions"] div span { display: none; }
    [data-testid="stFileUploaderDropzoneInstructions"] div::before {
        content: "여기에 파일을 끌어다 놓으세요";
        font-size: 14px;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] div small { display: none; }
    [data-testid="stFileUploaderDropzone"] button span { display: none; }
    [data-testid="stFileUploaderDropzone"] button::before {
        content: "파일 선택";
    }
    </style>
    """, unsafe_allow_html=True)

    st.divider()

    # ── 세션 상태 초기화 ───────────────────────────────────
    if "detected_users" not in st.session_state:
        st.session_state["detected_users"] = []
    if "results" not in st.session_state:
        st.session_state["results"] = None
    if "cal_summary" not in st.session_state:
        st.session_state["cal_summary"] = None

    # ══════════════════════════════════════════════════════
    # ① 파일 업로드
    # ══════════════════════════════════════════════════════
    st.subheader("① 파일 업로드")

    col_cal, col_tpl = st.columns(2)
    with col_cal:
        cal_file = st.file_uploader(
            "달력 파일 (.xls)",
            type=["xls"],
            help="주간활동계획서 달력 파일을 업로드하세요.",
        )
    with col_tpl:
        tpl_file = st.file_uploader(
            "계획서 템플릿 (.xlsx)",
            type=["xlsx"],
            help="이용자별 시트가 포함된 활동계획서 엑셀 파일을 업로드하세요.",
        )

    # 달력 파일 업로드 시 파싱 요약
    if cal_file:
        cal_bytes = cal_file.read()
        cal_file.seek(0)
        try:
            activities, holidays, month, year = parse_calendar(cal_bytes)
            working_days = sorted([d for d in activities.keys() if d not in holidays])
            st.session_state["cal_summary"] = {
                "activities": activities,
                "holidays": holidays,
                "month": month,
                "year": year,
                "working_days": working_days,
                "cal_bytes": cal_bytes,
            }
            holiday_str = f", 공휴일: {sorted(holidays)}" if holidays else ""
            st.success(f"{year}년 {month}월 달력 파싱 완료 — 활동일 {len(working_days)}일{holiday_str}")
        except Exception as e:
            st.error(f"달력 파싱 오류: {e}")
            st.session_state["cal_summary"] = None

    # 템플릿 파일 업로드 시 이용자 감지
    if tpl_file:
        tpl_bytes = tpl_file.read()
        tpl_file.seek(0)
        users = detect_users(tpl_bytes)
        if users:
            st.session_state["detected_users"] = users
            st.info(f"이용자 자동 감지: {', '.join(users)}")
        else:
            st.warning("시트명에서 이용자 이름을 감지하지 못했습니다. 시트명 끝에 이용자 이름(한글 2~3자)이 포함되어야 합니다.")

    st.divider()

    # ══════════════════════════════════════════════════════
    # ② 설정
    # ══════════════════════════════════════════════════════
    st.subheader("② 설정")

    provider = st.text_input(
        "담임(제공인력) 이름",
        value="천만석",
        placeholder="예: 천만석",
    )

    detected_users = st.session_state.get("detected_users", [])

    if detected_users:
        st.write("**이용자별 설정**")

        # 헤더
        hcols = st.columns([1.2, 1.5, 0.6, 0.6, 2, 0.6, 0.6, 2])
        hcols[0].markdown("**이용자**")
        hcols[1].markdown("**수급시간**")
        hcols[3].markdown("**오전송영**")
        hcols[4].markdown("**오전송영시간**")
        hcols[6].markdown("**오후송영**")
        hcols[7].markdown("**오후송영시간**")

        for user in detected_users:
            cols = st.columns([1.2, 1.5, 0.6, 0.6, 2, 0.6, 0.6, 2])
            # [이름, 수급시간, 공백, 오전체크, 오전시간, 공백, 오후체크, 오후시간]
            with cols[0]:
                st.markdown(f"<div style='padding-top:8px'>{user}</div>", unsafe_allow_html=True)
            with cols[1]:
                st.selectbox(
                    "수급시간", options=[132, 176],
                    format_func=lambda x: f"{x}시간({'기본형' if x == 132 else '확장형'})",
                    key=f"service_hours_{user}",
                    label_visibility="collapsed",
                )
            with cols[3]:
                st.checkbox(
                    "오전송영", value=True, key=f"am_shuttle_{user}",
                    label_visibility="collapsed",
                )
            with cols[4]:
                if st.session_state.get(f"am_shuttle_{user}", True):
                    st.text_input(
                        "오전송영시간", value="08:30~09:00 송영",
                        key=f"am_shuttle_time_{user}",
                        label_visibility="collapsed",
                    )
                else:
                    st.empty()
            with cols[6]:
                st.checkbox(
                    "오후송영", value=True, key=f"pm_shuttle_{user}",
                    label_visibility="collapsed",
                )
            with cols[7]:
                if st.session_state.get(f"pm_shuttle_{user}", True):
                    st.text_input(
                        "오후송영시간", value="16:00~16:30 송영",
                        key=f"pm_shuttle_time_{user}",
                        label_visibility="collapsed",
                    )
                else:
                    st.empty()
    else:
        st.info("계획서 템플릿(.xlsx)을 업로드하면 이용자별 설정이 표시됩니다.")

    st.divider()

    # ══════════════════════════════════════════════════════
    # ③ 검증 및 처리
    # ══════════════════════════════════════════════════════
    st.subheader("③ 검증 및 처리")

    has_users = bool(st.session_state.get("detected_users"))
    ready = bool(cal_file and tpl_file and provider and st.session_state.get("cal_summary") and has_users)
    if not ready:
        if cal_file and tpl_file and provider and not has_users:
            st.error("시트명에서 이용자 이름을 감지하지 못했습니다. 시트명 끝에 이용자 이름(한글 2~3자)이 포함되어야 합니다. (예: '03월 활동계획서-홍길동', '활동계획서 홍길동')")
        elif not (cal_file and tpl_file and provider):
            st.warning("달력 파일, 계획서 템플릿, 담임 이름을 모두 입력해야 처리할 수 있습니다.")

    # 활동일 수 vs 엑셀 행 수 안내
    if ready and tpl_file and st.session_state.get("cal_summary"):
        cal_info = st.session_state["cal_summary"]
        num_working_days = len(cal_info["working_days"])
        tpl_file.seek(0)
        available_rows = count_available_rows(tpl_file.read())
        tpl_file.seek(0)

        if num_working_days > available_rows:
            diff = num_working_days - available_rows
            st.info(
                f"{cal_info['month']}월 활동일 {num_working_days}일, "
                f"엑셀 행 {available_rows}개 — {diff}행을 자동으로 추가합니다."
            )
        elif num_working_days < available_rows:
            excess = available_rows - num_working_days
            st.info(
                f"{cal_info['month']}월 활동일 {num_working_days}일, "
                f"엑셀 행 {available_rows}개 — 불필요한 {excess}행을 자동으로 삭제합니다."
            )

    if st.button("처리하기", disabled=not ready, type="primary",
                 use_container_width=True):
        cal_info = st.session_state["cal_summary"]
        progress = st.progress(0, text="달력 데이터 준비 중...")

        # user_config 구성
        user_config = {}
        for user in detected_users:
            has_am = st.session_state.get(f"am_shuttle_{user}", False)
            am_time = st.session_state.get(f"am_shuttle_time_{user}", "08:30~09:00 송영")
            has_pm = st.session_state.get(f"pm_shuttle_{user}", False)
            pm_time = st.session_state.get(f"pm_shuttle_time_{user}", "16:00~16:30 송영")
            svc_hours = st.session_state.get(f"service_hours_{user}", 132)
            user_config[user] = {
                "오전송영": has_am,
                "오전송영시간": am_time,
                "오후송영": has_pm,
                "오후송영시간": pm_time,
                "수급시간": svc_hours,
            }

        progress.progress(0.3, text="활동계획서 입력 중...")

        try:
            tpl_file.seek(0)
            tpl_bytes = tpl_file.read()
            output_bytes, results, working_days, formulas_ok = fill_sheets(
                template_bytes=tpl_bytes,
                activities=cal_info["activities"],
                holidays=cal_info["holidays"],
                user_config=user_config,
                provider=provider,
                month=cal_info["month"],
                year=cal_info["year"],
            )

            progress.progress(1.0, text="완료!")
            st.session_state["results"] = {
                "bytes": output_bytes,
                "filename": tpl_file.name,
                "user_results": results,
                "working_days": working_days,
                "formulas_ok": formulas_ok,
                "holidays": cal_info["holidays"],
                "month": cal_info["month"],
            }
        except Exception as e:
            progress.empty()
            st.error(f"처리 중 오류 발생: {e}")
            import traceback
            st.code(traceback.format_exc())

    # 결과 표시 및 다운로드
    if st.session_state.get("results"):
        r = st.session_state["results"]
        st.success(f"처리 완료! {r['month']}월 활동일 {len(r['working_days'])}일")

        for ur in r["user_results"]:
            shuttle_parts = []
            if ur["오전송영"]:
                shuttle_parts.append("오전송영")
            if ur["오후송영"]:
                shuttle_parts.append("오후송영")
            shuttle_text = ", ".join(shuttle_parts) if shuttle_parts else "송영 없음"
            st.write(f"  - {ur['name']}: {ur['수급시간']}시간, {shuttle_text}, {ur['days']}일 입력")

        if r["holidays"]:
            st.info(f"공휴일 제외: {sorted(r['holidays'])}")
        if not r["formulas_ok"]:
            st.warning("수식 보존 확인이 필요합니다. 엑셀에서 L~O열 30행의 SUM 수식을 확인해주세요.")

        st.markdown("---")
        st.markdown(
            f"""<div style="background: #e8f5e9; border: 2px solid #4caf50; border-radius: 10px;
            padding: 20px; text-align: center; margin: 10px 0;">
            <p style="font-size: 1.1em; font-weight: bold; color: #2e7d32; margin-bottom: 5px;">
            파일이 준비되었습니다</p>
            <p style="font-size: 0.95em; color: #555; margin: 0;">
            {r['filename']}</p></div>""",
            unsafe_allow_html=True,
        )
        st.download_button(
            label="다운로드",
            data=r["bytes"],
            file_name=r["filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

    st.divider()
    st.info(
        "**처리 완료 후 반드시 확인해주세요**\n\n"
        "- 엑셀 파일을 열어 전체 내용을 검토하세요.\n"
        "- 공휴일, 대체공휴일이 올바르게 제외되었는지 확인하세요.\n"
        "- (협) 빨간색 서식이 정상적으로 적용되었는지 확인하세요.\n"
        "- 수식(합계)이 정상적으로 보존되었는지 확인하세요."
    )


# ══════════════════════════════════════════════════════════
# 접근 코드 게이트
# ══════════════════════════════════════════════════════════
def _access_gate():
    """접근 코드 인증 화면. 인증 완료 시 True 반환."""
    if "access_granted" not in st.session_state:
        st.session_state["access_granted"] = False

    if st.session_state["access_granted"]:
        if "ac" in st.query_params:
            del st.query_params["ac"]
        return True

    # URL ?ac= 파라미터 자동 인증
    try:
        secret_code = st.secrets["ACCESS_CODE"]
    except Exception:
        secret_code = "2026"

    ac_param = st.query_params.get("ac", "")
    if ac_param and ac_param == secret_code:
        st.session_state["access_granted"] = True
        st.rerun()

    # 수동 입력 화면
    st.markdown("""
    <style>
    [data-testid="stHeader"], [data-testid="stToolbar"],
    [data-testid="stDecoration"], [data-testid="stStatusWidget"] {
        display: none !important;
    }
    </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        st.markdown("## 월별 활동계획서 자동입력")
        st.caption("접근 코드를 입력해주세요.")
        st.markdown("")
        code_input = st.text_input(
            "접근 코드",
            type="password",
            placeholder="접근 코드를 입력하세요",
            label_visibility="collapsed",
        )
        if st.button("입장하기", type="primary", use_container_width=True):
            if code_input == secret_code:
                st.session_state["access_granted"] = True
                st.rerun()
            else:
                st.error("접근 코드가 올바르지 않습니다.")
    return False


# ══════════════════════════════════════════════════════════
# 진입점
# ══════════════════════════════════════════════════════════
if _access_gate():
    main_app()
