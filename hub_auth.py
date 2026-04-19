"""
복서방 허브 JWT 인증 모듈 (Streamlit용).

복서방에서 발급한 30분짜리 JWT를 검증하고 세션에 저장한다.
- URL `?token=...` 쿼리에서 토큰을 읽어 검증 후 즉시 제거
- 검증 성공 시 st.session_state 에 사용자 정보 저장
- 만료 시 재접속 안내
"""
from __future__ import annotations

import os
import time
from typing import Optional

import jwt
import streamlit as st

HUB_URL = "https://www.bokji-ai.co.kr/"
APP_KEY = "plan2"  # 이 앱의 서브앱 식별자 (복서방 ALLOWED_APPS 와 일치)
ISSUER = "bokseobang-hub"


def _get_secret() -> Optional[str]:
    # Streamlit Cloud → st.secrets, 로컬 → 환경변수
    try:
        return st.secrets["HUB_JWT_SECRET"]
    except Exception:
        return os.environ.get("HUB_JWT_SECRET")


def _verify_token(token: str):
    secret = _get_secret()
    if not secret:
        return None, "서버 설정 오류: HUB_JWT_SECRET 미설정"
    try:
        payload = jwt.decode(
            token,
            secret,
            algorithms=["HS256"],
            options={"require": ["exp", "iat"]},
        )
        if payload.get("app") != APP_KEY:
            return None, "이 앱에 대한 접근 권한이 없는 토큰입니다."
        if payload.get("iss") != ISSUER:
            return None, "신뢰할 수 없는 발급자입니다."
        return payload, None
    except jwt.ExpiredSignatureError:
        return None, "expired"
    except jwt.InvalidTokenError as e:
        return None, f"유효하지 않은 토큰: {e}"


def _render_denied(reason: str = ""):
    st.markdown(
        """
        <style>
        [data-testid="stHeader"], [data-testid="stToolbar"],
        [data-testid="stDecoration"], [data-testid="stStatusWidget"] { display: none !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        if reason == "expired":
            st.markdown("## ⏰ 세션이 만료되었어요")
            st.write("30분 허브 세션이 지났습니다. 복서방 허브에서 다시 접속해주세요.")
        else:
            st.markdown("## 🔒 복서방 로그인이 필요해요")
            st.write(
                "계획서해방은 복서방에 로그인한 선생님만 사용하실 수 있습니다. "
                "아래 버튼을 눌러 허브에서 접속해주세요."
            )
            if reason and reason not in ("no-token",):
                st.caption(reason)
        st.markdown(
            f'<a href="{HUB_URL}" target="_self" '
            'style="display:inline-block;margin-top:16px;padding:12px 24px;'
            'background:#4a90d9;color:#fff;border-radius:10px;'
            'text-decoration:none;font-weight:600;">'
            "복서방 허브로 이동 →</a>",
            unsafe_allow_html=True,
        )


def hub_gate() -> bool:
    """복서방 허브 JWT 게이트. 인증 통과 시 True 반환."""
    # 이미 인증된 세션이 있으면 만료만 체크
    if st.session_state.get("hub_auth"):
        auth = st.session_state["hub_auth"]
        if auth.get("exp", 0) > time.time():
            # URL에 token이 아직 남아 있으면 조용히 제거
            if "token" in st.query_params:
                try:
                    del st.query_params["token"]
                except Exception:
                    pass
            return True
        # 만료
        st.session_state.pop("hub_auth", None)
        _render_denied("expired")
        return False

    # URL 토큰 파라미터 읽기
    token = st.query_params.get("token", "")
    if isinstance(token, list):
        token = token[0] if token else ""

    if not token:
        _render_denied("no-token")
        return False

    # JWT 검증
    payload, err = _verify_token(token)
    if err == "expired":
        _render_denied("expired")
        return False
    if err:
        _render_denied(err)
        return False

    # 세션에 저장 후 명시적 rerun
    st.session_state["hub_auth"] = {
        "user_id": payload.get("sub"),
        "name": payload.get("name"),
        "org": payload.get("org"),
        "role": payload.get("role"),
        "plan_end": payload.get("plan_end"),
        "exp": payload.get("exp"),
    }
    st.rerun()
    return False  # rerun 후 위의 hub_auth 분기에서 True 반환


def render_session_bar():
    """인증된 사용자 정보 상단 바."""
    auth = st.session_state.get("hub_auth")
    if not auth:
        return
    st.markdown(
        f"""
        <div style="background:#eef5ff;border-bottom:1px solid #cfe0f7;
                    padding:6px 14px;font-size:13px;color:#1a4a8a;text-align:center;">
          <b>{auth.get("name", "")}</b> 선생님 · {auth.get("org") or "복서방"} 허브 세션 (30분)
        </div>
        """,
        unsafe_allow_html=True,
    )
