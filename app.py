import base64
import html
import json
import os
import re
from datetime import date
from pathlib import Path

import pandas as pd
import resend
import streamlit as st
from dotenv import load_dotenv

from core.data_model import ReasonStatementData, TripReportData
from core.file_utils import (
    date_range,
    extract_text_from_file,
    iso_date,
    korean_date,
    save_uploaded_files,
    slugify,
)
from core.llm import LLMClient
from core.prompts import (
    REASON_STATEMENT_SYSTEM,
    TRIP_PURPOSE_SYSTEM,
    TRIP_RESEARCH_RELEVANCE_SYSTEM,
    TRIP_SCHEDULE_SYSTEM,
    TRIP_EXPECTED_EFFECT_SYSTEM,
)
from core.report_builder import create_zip, generate_reason_pdf_on_template, generate_trip_docx, generate_trip_pdf
from core.validation import check_missing_documents, checklist_rows


load_dotenv()

APP_ROOT = Path(__file__).parent
UPLOAD_ROOT = APP_ROOT / "data" / "uploads"
OUTPUT_ROOT = APP_ROOT / "data" / "outputs"
UPLOAD_ROOT.mkdir(parents=True, exist_ok=True)
OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

DEFAULT_EMAIL_RECIPIENT = "pung@khu.ac.kr"


st.set_page_config(
    page_title="KHU Trip Report System",
    layout="wide",
    initial_sidebar_state="expanded",
)


def apply_ui_style() -> None:
    st.markdown(
        """
        <style>
        :root {
            --primary: #1d4ed8;
            --primary-dark: #1e3a8a;
            --primary-soft: #eff6ff;
            --border: #e5e7eb;
            --muted: #64748b;
            --surface: #ffffff;
            --soft-surface: #f8fafc;
            --success-soft: #ecfdf5;
            --warning-soft: #fffbeb;
            --danger-soft: #fef2f2;
        }

        .main .block-container {
            padding-top: 1.35rem;
            padding-bottom: 3rem;
            max-width: 1320px;
        }

        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #f8fafc 0%, #eef2ff 100%);
            border-right: 1px solid var(--border);
        }

        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3 {
            color: #0f172a;
        }

        .app-hero {
            position: relative;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1.3rem;
            padding: 1.35rem 1.45rem;
            border-radius: 20px;
            background:
                radial-gradient(circle at 8% 20%, rgba(37, 99, 235, 0.10), transparent 30%),
                linear-gradient(135deg, #ffffff 0%, #f8fbff 62%, #eef6ff 100%);
            border: 1px solid #dbeafe;
            color: #0f172a;
            box-shadow: 0 12px 30px rgba(15, 23, 42, 0.08);
            margin-bottom: 1.25rem;
            overflow: hidden;
        }

        .app-hero::before {
            content: "";
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 6px;
            background: linear-gradient(180deg, #2563eb 0%, #38bdf8 100%);
        }

        .hero-kicker {
            display: inline-flex;
            align-items: center;
            padding: 0.28rem 0.65rem;
            border-radius: 999px;
            background: #eff6ff;
            border: 1px solid #bfdbfe;
            color: #1d4ed8;
            font-size: 0.78rem;
            font-weight: 800;
            letter-spacing: 0.01em;
            margin-bottom: 0.55rem;
        }

        .app-hero h1 {
            margin: 0;
            font-size: 2.05rem;
            line-height: 1.18;
            letter-spacing: -0.035em;
            color: #0f172a;
        }

        .app-hero p {
            margin: 0.6rem 0 0;
            color: #475569;
            font-size: 0.96rem;
            line-height: 1.65;
        }

        .hero-panel {
            min-width: 210px;
            padding: 1rem 1.1rem;
            border-radius: 18px;
            background: linear-gradient(135deg, #1d4ed8 0%, #2563eb 100%);
            color: #ffffff;
            box-shadow: 0 12px 24px rgba(37, 99, 235, 0.20);
            text-align: left;
        }

        .hero-panel .panel-label {
            font-size: 0.74rem;
            opacity: 0.82;
            font-weight: 700;
            margin-bottom: 0.45rem;
        }

        .hero-panel .panel-title {
            font-size: 1.08rem;
            font-weight: 900;
            line-height: 1.35;
        }

        .hero-panel .panel-hint {
            margin-top: 0.35rem;
            font-size: 0.78rem;
            opacity: 0.82;
        }

        .hero-panel.issue-panel {
            cursor: pointer;
        }

        .st-key-hero_wrap {
            position: relative;
            margin-bottom: 1.25rem;
        }

        .st-key-hero_wrap .app-hero {
            margin-bottom: 0;
        }

        .st-key-hero_issue_click_layer {
            position: absolute !important;
            top: 1.35rem;
            right: 1.45rem;
            width: 210px;
            min-width: 210px;
            height: 108px;
            z-index: 50;
            margin: 0 !important;
            padding: 0 !important;
        }

        .st-key-hero_issue_click_layer div[data-testid="stButton"] {
            width: 100%;
            height: 100%;
            margin: 0 !important;
            padding: 0 !important;
        }

        .st-key-hero_issue_click_layer div[data-testid="stButton"] > button {
            width: 100%;
            height: 100%;
            min-height: 108px;
            padding: 0 !important;
            margin: 0 !important;
            border: 0 !important;
            border-radius: 18px !important;
            background: transparent !important;
            color: transparent !important;
            box-shadow: none !important;
            opacity: 0.01;
            cursor: pointer;
        }

        .st-key-hero_issue_click_layer div[data-testid="stButton"] > button:hover {
            transform: none !important;
            box-shadow: none !important;
            background: rgba(255, 255, 255, 0.01) !important;
        }

        @media (max-width: 900px) {
            .app-hero {
                flex-direction: column;
                align-items: flex-start;
            }

            .hero-panel {
                width: 100%;
                min-width: 0;
            }
        }

        .status-card {
            border: 1px solid var(--border);
            background: var(--surface);
            border-radius: 18px;
            padding: 1rem 1.05rem;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
            min-height: 92px;
        }

        .status-card .label {
            color: var(--muted);
            font-size: 0.82rem;
            margin-bottom: 0.35rem;
            font-weight: 600;
        }

        .status-card .value {
            color: #0f172a;
            font-size: 1.05rem;
            font-weight: 800;
            margin-bottom: 0.15rem;
        }

        .status-card .hint {
            color: #64748b;
            font-size: 0.8rem;
        }

        .sidebar-card {
            border: 1px solid rgba(148, 163, 184, 0.25);
            background: rgba(255, 255, 255, 0.72);
            border-radius: 16px;
            padding: 0.95rem 1rem;
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.05);
            margin: 0.6rem 0 1rem;
        }

        .sidebar-card-title {
            font-weight: 800;
            color: #0f172a;
            margin-bottom: 0.35rem;
        }

        .sidebar-card-text {
            color: #475569;
            font-size: 0.86rem;
            line-height: 1.55;
        }

        div[data-testid="stExpander"] {
            border: 1px solid var(--border);
            border-radius: 18px;
            background: var(--surface);
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.045);
            margin-bottom: 1rem;
            overflow: hidden;
        }

        div[data-testid="stExpander"] details summary {
            background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
            border-radius: 18px;
            font-weight: 800;
            color: #0f172a;
            padding-top: 0.85rem !important;
            padding-bottom: 0.85rem !important;
        }

        div[data-testid="stTextInput"] input,
        div[data-testid="stTextArea"] textarea,
        div[data-testid="stDateInput"] input {
            border-radius: 12px;
            border-color: #dbe3ef;
        }

        div[data-testid="stFileUploader"] section {
            border-radius: 16px;
            border: 1px dashed #93c5fd;
            background: #f8fbff;
        }

        div.stButton > button,
        div[data-testid="stDownloadButton"] > button {
            border-radius: 12px;
            font-weight: 700;
            border: 1px solid #bfdbfe;
            transition: all 0.18s ease;
        }

        div.stButton > button:hover,
        div[data-testid="stDownloadButton"] > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 8px 18px rgba(37, 99, 235, 0.14);
        }

        div.stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #1d4ed8 0%, #2563eb 100%);
            border: 0;
            color: white;
        }

        .stDataFrame {
            border-radius: 16px;
            overflow: hidden;
        }

        hr {
            margin: 1.15rem 0;
        }

        .block-note {
            padding: 0.85rem 1rem;
            border-radius: 14px;
            background: #ffffff;
            border: 1px solid #e5e7eb;
            color: #475569;
            font-size: 0.88rem;
            line-height: 1.55;
            margin-top: 1.25rem;
            margin-bottom: 1rem;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.035);
        }

        .final-action-box {
            margin-top: 1.25rem;
            padding: 1.1rem 1.2rem;
            border-radius: 18px;
            border: 1px solid #dbeafe;
            background: linear-gradient(135deg, #ffffff 0%, #f8fbff 100%);
            box-shadow: 0 10px 26px rgba(15, 23, 42, 0.06);
        }

        .final-action-title {
            color: #0f172a;
            font-size: 1.1rem;
            font-weight: 900;
            margin-bottom: 0.25rem;
        }

        .final-action-caption {
            color: #64748b;
            font-size: 0.88rem;
            line-height: 1.55;
            margin-bottom: 0.85rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def open_issue_report_from_hero() -> None:
    st.session_state["show_issue_report_dialog"] = True


def render_hero() -> None:
    with st.container(key="hero_wrap"):
        st.markdown(
            """
            <div class="app-hero">
                <div class="hero-text">
                    <div class="hero-kicker">Kyung Hee University · Research Administration</div>
                    <h1>KHU Trip Report System</h1>
                    <p>Overview 추출, 출장목적 및 세부일정 자동생성, 증빙자료 정리, ZIP 패키지 생성 및 이메일 발송까지 한 번에 처리합니다.</p>
                </div>
                <div class="hero-panel issue-panel">
                    <div class="panel-label">Support</div>
                    <div class="panel-title">Report an Issue</div>
                    <div class="panel-hint">Click here if the system fails or returns incorrect results.</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.container(key="hero_issue_click_layer"):
            st.button(
                " ",
                key="hero_issue_click_btn",
                help="Report an Issue",
                use_container_width=True,
                on_click=open_issue_report_from_hero,
            )


def render_status_cards() -> None:
    openai_ready = bool(get_config_value("OPENAI_API_KEY", ""))
    resend_ready = bool(get_config_value("RESEND_API_KEY", ""))
    zip_path = st.session_state.get("last_generated_zip", "")
    zip_ready = bool(zip_path and Path(zip_path).exists())
    checklist_count = len(st.session_state.get("last_checklist", []))
    missing_count = len(st.session_state.get("last_missing", []))

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown(
            f"""
            <div class="status-card">
                <div class="label">OpenAI 설정</div>
                <div class="value">{'완료' if openai_ready else '확인 필요'}</div>
                <div class="hint">자동생성 기능 사용 가능 여부</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col2:
        st.markdown(
            f"""
            <div class="status-card">
                <div class="label">Resend 이메일</div>
                <div class="value">{'완료' if resend_ready else '확인 필요'}</div>
                <div class="hint">서류 발송 설정</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col3:
        st.markdown(
            f"""
            <div class="status-card">
                <div class="label">생성 파일</div>
                <div class="value">{'준비됨' if zip_ready else '미생성'}</div>
                <div class="hint">DOCX, PDF, ZIP 패키지</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col4:
        st.markdown(
            f"""
            <div class="status-card">
                <div class="label">누락 검사</div>
                <div class="value">{missing_count}건</div>
                <div class="hint">체크리스트 {checklist_count}개 항목 기준</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def init_state() -> None:
    defaults = {
        "overview_text": "",
        "purpose_text": "",
        "research_relatedness": "",
        "expected_effect_text": "",
        "daily_schedule": [],
        "date_signature": "",
        "reason_title": "",
        "reason_content": "",
        "last_missing": [],
        "last_checklist": [],
        "last_generated_docx": "",
        "last_generated_pdf": "",
        "last_generated_reason_pdf": "",
        "last_generated_zip": "",
        "last_generation_notice": "",
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def get_config_value(name: str, default: str = "") -> str:
    value = os.getenv(name)
    if value:
        return value

    try:
        value = st.secrets.get(name, "")
        if value:
            return str(value)
    except Exception:
        pass

    return default


def get_float_config_value(name: str, default: float) -> float:
    raw_value = get_config_value(name, str(default)).strip()

    try:
        return float(raw_value)
    except ValueError:
        return default



ISSUE_EMAIL_FROM = "Error Issue <onboarding@resend.dev>"


def get_issue_recipient() -> str:
    return (
        get_config_value("ISSUE_REPORT_TO_EMAIL", "")
        or get_config_value("RESEND_ISSUE_TO_EMAIL", "")
        or get_config_value("RESEND_TO_EMAIL", "")
        or DEFAULT_EMAIL_RECIPIENT
    )


def build_issue_email_body(
    issue_type: str,
    severity: str,
    reporter_name: str,
    reporter_email: str,
    description: str,
    include_context: bool,
) -> tuple[str, str]:
    context = {
        "app": "KHU Trip Report System",
        "conference_name": st.session_state.get("conference_name", ""),
        "trip_type": st.session_state.get("trip_type", ""),
        "model": st.session_state.get("model", get_config_value("OPENAI_MODEL", "")),
        "has_generated_zip": bool(st.session_state.get("last_generated_zip", "")),
        "timestamp": datetime.now().isoformat(timespec="seconds"),
    }

    context_text = "\n".join([f"{key}: {value}" for key, value in context.items()]) if include_context else "Not included"

    plain_text = f"""KHU Trip Report System Issue Report

Issue type: {issue_type}
Severity: {severity}
Reporter name: {reporter_name or "Not provided"}
Reporter email: {reporter_email or "Not provided"}

Description:
{description}

Runtime context:
{context_text}
"""

    html_body = f"""
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #0f172a;">
        <h2 style="margin-bottom: 8px;">KHU Trip Report System Issue Report</h2>
        <p style="margin-top: 0; color: #475569;">A user submitted an issue report from the Streamlit app.</p>

        <table style="border-collapse: collapse; width: 100%; margin: 16px 0;">
            <tr>
                <td style="border: 1px solid #e5e7eb; padding: 8px; font-weight: bold;">Issue type</td>
                <td style="border: 1px solid #e5e7eb; padding: 8px;">{html.escape(issue_type)}</td>
            </tr>
            <tr>
                <td style="border: 1px solid #e5e7eb; padding: 8px; font-weight: bold;">Severity</td>
                <td style="border: 1px solid #e5e7eb; padding: 8px;">{html.escape(severity)}</td>
            </tr>
            <tr>
                <td style="border: 1px solid #e5e7eb; padding: 8px; font-weight: bold;">Reporter name</td>
                <td style="border: 1px solid #e5e7eb; padding: 8px;">{html.escape(reporter_name or "Not provided")}</td>
            </tr>
            <tr>
                <td style="border: 1px solid #e5e7eb; padding: 8px; font-weight: bold;">Reporter email</td>
                <td style="border: 1px solid #e5e7eb; padding: 8px;">{html.escape(reporter_email or "Not provided")}</td>
            </tr>
        </table>

        <h3>Description</h3>
        <div style="white-space: pre-wrap; padding: 12px; border-radius: 10px; background: #f8fafc; border: 1px solid #e5e7eb;">
            {html.escape(description)}
        </div>

        <h3>Runtime context</h3>
        <pre style="white-space: pre-wrap; padding: 12px; border-radius: 10px; background: #f8fafc; border: 1px solid #e5e7eb;">{html.escape(context_text)}</pre>
    </div>
    """

    return plain_text, html_body


def send_issue_report_email(
    issue_type: str,
    severity: str,
    reporter_name: str,
    reporter_email: str,
    description: str,
    include_context: bool,
) -> dict:
    resend_api_key = get_config_value("RESEND_API_KEY", "")
    if not resend_api_key:
        raise RuntimeError("RESEND_API_KEY is not configured.")

    recipient = get_issue_recipient()
    if not recipient:
        raise RuntimeError("Issue report recipient email is not configured.")

    text_body, html_body = build_issue_email_body(
        issue_type=issue_type,
        severity=severity,
        reporter_name=reporter_name,
        reporter_email=reporter_email,
        description=description,
        include_context=include_context,
    )

    resend.api_key = resend_api_key

    payload = {
        "from": ISSUE_EMAIL_FROM,
        "to": [recipient],
        "subject": f"[KHU Trip Report] {issue_type} Issue Report",
        "html": html_body,
        "text": text_body,
    }

    if reporter_email.strip():
        payload["reply_to"] = reporter_email.strip()

    try:
        response = resend.Emails.send(payload)
        if isinstance(response, dict):
            return response
        return {"response": str(response)}
    except Exception as exc:
        raise RuntimeError(f"Resend API issue report failed: {exc}") from exc


@st.dialog("Report an Issue", dismissible=True, width="large")
def open_issue_report_dialog() -> None:
    st.caption("Use this form to report system errors, incorrect results, generation failures, or email delivery issues.")

    with st.form("issue_report_form", clear_on_submit=False):
        issue_type = st.selectbox(
            "Issue type",
            [
                "System error",
                "Incorrect output",
                "Report generation failure",
                "File download problem",
                "Email sending problem",
                "Other",
            ],
            index=0,
        )
        severity = st.selectbox(
            "Severity",
            ["Low", "Medium", "High", "Critical"],
            index=1,
        )

        col1, col2 = st.columns(2)
        with col1:
            reporter_name = st.text_input("Your name", placeholder="Optional")
        with col2:
            reporter_email = st.text_input("Your email", placeholder="Optional")

        description = st.text_area(
            "Issue description *",
            height=180,
            placeholder="Please describe what happened, which step failed, and any error message shown on the screen.",
        )
        include_context = st.checkbox(
            "Include basic runtime context",
            value=True,
            help="Includes app name, conference name, trip type, model, generated ZIP status, and timestamp.",
        )

        submitted = st.form_submit_button(
            "Send Issue Report",
            type="primary",
            use_container_width=True,
        )

    if submitted:
        if not description.strip():
            st.warning("Please enter an issue description before sending.")
        else:
            try:
                response = send_issue_report_email(
                    issue_type=issue_type,
                    severity=severity,
                    reporter_name=reporter_name,
                    reporter_email=reporter_email,
                    description=description,
                    include_context=include_context,
                )
                message_id = response.get("id", "") if isinstance(response, dict) else ""
                if message_id:
                    st.success(f"Issue report sent successfully. Resend ID: {message_id}")
                else:
                    st.success("Issue report sent successfully.")
            except Exception as exc:
                st.error(f"Failed to send issue report: {exc}")

    if st.button("Close", use_container_width=True, key="close_issue_report_dialog_btn"):
        st.session_state["show_issue_report_dialog"] = False
        st.rerun()


def get_llm() -> LLMClient:
    return LLMClient(
        api_key=st.session_state.get("api_key") or get_config_value("OPENAI_API_KEY", ""),
        model=st.session_state.get("model") or get_config_value("OPENAI_MODEL", "gpt-5.5"),
    )


def get_default_expected_effect() -> str:
    return (
        f"이번 {conference_name or '학회'} 참여를 통해 파악한 최신 연구 동향과 주요 세션 내용은 "
        f"향후 본 연구과제의 연구 방향을 구체화하고 관련 분야의 학술적 네트워크를 확대하는 데 "
        f"기여할 것으로 기대된다."
    )


def generate_expected_effect_text() -> str:
    prompt = f"""
학회명: {conference_name}
출장 구분: {trip_type}
출장기간: {korean_date(start_date)} ~ {korean_date(end_date)}
본인 연구 주제: {research_theme}

출장목적:
{st.session_state.get("purpose_text", "")}

본 연구와 관련성 및 주요 세션 요약:
{st.session_state.get("research_relatedness", "")}

학회 Overview:
{st.session_state.get("overview_text", "")}
"""

    output = get_llm().generate(
        TRIP_EXPECTED_EFFECT_SYSTEM,
        prompt,
        temperature=0.25,
        max_output_tokens=900,
    )

    cleaned = str(output).strip()
    if not cleaned:
        return get_default_expected_effect()

    return cleaned


def parse_schedule_json(text: str, dates: list[date]) -> list[dict]:
    raw = text.strip()
    raw = re.sub(r"^```(?:json)?", "", raw).strip()
    raw = re.sub(r"```$", "", raw).strip()

    match = re.search(r"\[[\s\S]*\]", raw)
    if match:
        raw = match.group(0)

    try:
        data = json.loads(raw)
        output = []

        for i, d in enumerate(dates):
            if i < len(data):
                item = data[i]
                output.append(
                    {
                        "day": int(item.get("day", i + 1)),
                        "date": item.get("date", iso_date(d)),
                        "content": str(item.get("content", "")).strip()[:120],
                    }
                )
            else:
                output.append({"day": i + 1, "date": iso_date(d), "content": ""})

        return output
    except Exception:
        output = []
        lines = [line.strip("-• 0123456789.일차") for line in raw.splitlines() if line.strip()]

        for i, d in enumerate(dates):
            content = lines[i] if i < len(lines) else "학회 일정 확인 및 관련 세션 참석"
            output.append({"day": i + 1, "date": iso_date(d), "content": content[:120]})

        return output


def ensure_schedule_for_dates(dates: list[date]) -> None:
    signature = "|".join([iso_date(d) for d in dates])

    if st.session_state["date_signature"] != signature:
        st.session_state["date_signature"] = signature
        st.session_state["daily_schedule"] = [
            {"day": i + 1, "date": iso_date(d), "content": ""}
            for i, d in enumerate(dates)
        ]

        for i, d in enumerate(dates):
            st.session_state[f"schedule_content_{i}"] = ""


def sync_schedule_from_widgets(dates: list[date]) -> list[dict]:
    schedule = []

    for i, d in enumerate(dates):
        content = st.session_state.get(f"schedule_content_{i}", "")
        schedule.append({"day": i + 1, "date": iso_date(d), "content": content})

    st.session_state["daily_schedule"] = schedule
    return schedule


def apply_schedule_to_widgets(schedule: list[dict]) -> None:
    st.session_state["daily_schedule"] = schedule

    for i, row in enumerate(schedule):
        st.session_state[f"schedule_content_{i}"] = row.get("content", "")


def style_checklist(df: pd.DataFrame):
    def color_status(value):
        if value == "누락":
            return "background-color: #f8d7da; color: #842029; font-weight: bold;"
        if value == "완료":
            return "background-color: #d1e7dd; color: #0f5132; font-weight: bold;"
        if value == "선택":
            return "background-color: #e2e3e5; color: #41464b;"
        return ""

    return df.style.map(color_status, subset=["상태"])


def download_button(path: str, label: str, mime: str) -> None:
    p = Path(path)

    if p.exists():
        with open(p, "rb") as f:
            st.download_button(
                label=label,
                data=f.read(),
                file_name=p.name,
                mime=mime,
                use_container_width=True,
            )


def get_resend_config() -> dict:
    api_key = get_config_value("RESEND_API_KEY", "")
    sender = get_config_value("RESEND_FROM_EMAIL", "")
    default_recipient = get_config_value("RESEND_TO_EMAIL", DEFAULT_EMAIL_RECIPIENT)
    subject_prefix = get_config_value("RESEND_SUBJECT_PREFIX", "[출장보고서]")
    max_attachment_mb = get_float_config_value("RESEND_MAX_ATTACHMENT_MB", 20.0)

    if not api_key:
        raise ValueError("RESEND_API_KEY가 설정되지 않았습니다. Streamlit Cloud의 App settings > Secrets에 추가하세요.")

    if not sender:
        raise ValueError("RESEND_FROM_EMAIL이 설정되지 않았습니다. Resend에서 사용 가능한 발신자 이메일을 설정하세요.")

    return {
        "api_key": api_key,
        "sender": sender,
        "default_recipient": default_recipient,
        "subject_prefix": subject_prefix,
        "max_attachment_bytes": int(max_attachment_mb * 1024 * 1024),
    }


def normalize_recipients(recipient: str) -> list[str]:
    recipients = [email.strip() for email in re.split(r"[,;]", recipient or "") if email.strip()]

    if not recipients:
        raise ValueError("수신자 이메일을 입력하세요.")

    return recipients


def add_subject_prefix(subject: str, prefix: str) -> str:
    cleaned_subject = (subject or "").strip()
    cleaned_prefix = (prefix or "").strip()

    if cleaned_prefix and not cleaned_subject.startswith(cleaned_prefix):
        return f"{cleaned_prefix} {cleaned_subject}"

    return cleaned_subject


def send_email_with_attachment(
    recipient: str,
    subject: str,
    body: str,
    attachment_path: str | Path,
) -> dict:
    attachment = Path(attachment_path)

    if not attachment.exists():
        raise FileNotFoundError(f"첨부파일을 찾을 수 없습니다: {attachment}")

    resend_config = get_resend_config()

    if attachment.stat().st_size > resend_config["max_attachment_bytes"]:
        max_mb = resend_config["max_attachment_bytes"] / 1024 / 1024
        raise ValueError(
            f"첨부 ZIP 파일이 {max_mb:.0f}MB를 초과합니다. 파일을 분할하거나 다운로드 방식으로 제출하세요."
        )

    encoded_attachment = base64.b64encode(attachment.read_bytes()).decode("utf-8")
    safe_html_body = html.escape(body or "").replace("\n", "<br>")

    resend.api_key = resend_config["api_key"]

    payload = {
        "from": resend_config["sender"],
        "to": normalize_recipients(recipient),
        "subject": add_subject_prefix(subject, resend_config["subject_prefix"]),
        "text": body or "",
        "html": f"<p>{safe_html_body}</p>",
        "attachments": [
            {
                "filename": attachment.name,
                "content": encoded_attachment,
            }
        ],
    }

    try:
        response = resend.Emails.send(payload)
        if isinstance(response, dict):
            return response
        return {"response": str(response)}
    except Exception as exc:
        raise RuntimeError(f"Resend API 발송 실패: {exc}") from exc


def render_generated_downloads() -> None:
    generated_docx = st.session_state.get("last_generated_docx", "")
    generated_pdf = st.session_state.get("last_generated_pdf", "")
    generated_reason_pdf = st.session_state.get("last_generated_reason_pdf", "")
    generated_zip = st.session_state.get("last_generated_zip", "")

    has_docx = bool(generated_docx and Path(generated_docx).exists())
    has_pdf = bool(generated_pdf and Path(generated_pdf).exists())
    has_reason = bool(generated_reason_pdf and Path(generated_reason_pdf).exists())
    has_zip = bool(generated_zip and Path(generated_zip).exists())

    if not any([has_docx, has_pdf, has_reason, has_zip]):
        return

    st.markdown("#### 생성된 파일 다운로드")

    if has_reason:
        download_cols = st.columns(4)

        with download_cols[0]:
            if has_docx:
                download_button(
                    generated_docx,
                    "DOCX 다운로드",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

        with download_cols[1]:
            if has_pdf:
                download_button(
                    generated_pdf,
                    "PDF 다운로드",
                    "application/pdf",
                )

        with download_cols[2]:
            download_button(
                generated_reason_pdf,
                "사유서 PDF",
                "application/pdf",
            )

        with download_cols[3]:
            if has_zip:
                download_button(
                    generated_zip,
                    "ZIP 다운로드",
                    "application/zip",
                )
    else:
        download_cols = st.columns(3)

        with download_cols[0]:
            if has_docx:
                download_button(
                    generated_docx,
                    "DOCX 다운로드",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

        with download_cols[1]:
            if has_pdf:
                download_button(
                    generated_pdf,
                    "PDF 다운로드",
                    "application/pdf",
                )

        with download_cols[2]:
            if has_zip:
                download_button(
                    generated_zip,
                    "ZIP 다운로드",
                    "application/zip",
                )


@st.dialog("서류 발송")
def open_zip_email_dialog() -> None:
    last_zip_path = st.session_state.get("last_generated_zip", "")

    if not last_zip_path or not Path(last_zip_path).exists():
        st.warning("먼저 출장보고서를 생성하면 이메일로 서류 발송이 가능합니다.")
        return

    st.caption(f"발송 대상 ZIP: {Path(last_zip_path).name}")

    email_recipient = st.text_input(
        "수신자 이메일",
        value=get_config_value("RESEND_TO_EMAIL", DEFAULT_EMAIL_RECIPIENT),
        key="email_recipient_dialog",
    )
    email_subject = st.text_input(
        "메일 제목",
        value=f"{conference_name or '학회'} 출장결과보고서 ZIP 패키지 송부",
        key="email_subject_dialog",
    )
    email_body = st.text_area(
        "메일 본문",
        value=(
            "안녕하세요.\n\n"
            f"{conference_name or '학회'} 출장결과보고서 ZIP 패키지 파일을 첨부드립니다.\n\n"
            "감사합니다."
        ),
        height=160,
        key="email_body_dialog",
    )

    action_cols = st.columns(2)

    with action_cols[0]:
        send_clicked = st.button(
            "이메일 발송",
            type="primary",
            key="send_zip_email_dialog_btn",
            use_container_width=True,
        )

    with action_cols[1]:
        close_clicked = st.button(
            "닫기",
            key="close_zip_email_dialog_btn",
            use_container_width=True,
        )

    if send_clicked:
        try:
            resend_response = send_email_with_attachment(
                recipient=email_recipient,
                subject=email_subject,
                body=email_body,
                attachment_path=last_zip_path,
            )
            message_id = resend_response.get("id", "")

            if message_id:
                st.session_state["email_dialog_result"] = (
                    f"전체 ZIP 파일을 {email_recipient}로 발송했습니다. Resend ID: {message_id}"
                )
            else:
                st.session_state["email_dialog_result"] = (
                    f"전체 ZIP 파일을 {email_recipient}로 발송했습니다."
                )

        except Exception as exc:
            st.session_state["email_dialog_error"] = f"이메일 발송 실패: {exc}"

    if close_clicked:
        st.session_state.pop("email_dialog_result", None)
        st.session_state.pop("email_dialog_error", None)
        st.rerun()

    if st.session_state.get("email_dialog_result"):
        st.success(st.session_state["email_dialog_result"])

    if st.session_state.get("email_dialog_error"):
        st.error(st.session_state["email_dialog_error"])


init_state()
apply_ui_style()

with st.sidebar:

    st.subheader("기본 설정")
    st.session_state["api_key"] = st.text_input(
        "OpenAI API Key",
        value=get_config_value("OPENAI_API_KEY", ""),
        type="password",
        help="Streamlit Secrets 또는 환경변수에 저장된 값을 기본으로 불러옵니다.",
    )
    st.session_state["model"] = st.text_input(
        "Model",
        value=get_config_value("OPENAI_MODEL", "gpt-5.5"),
        help="자동생성 기능에 사용할 모델명입니다.",
    )

    st.divider()
    st.subheader("이메일 발송 설정")

    if get_config_value("RESEND_API_KEY", ""):
        st.success("Resend API Key 설정이 완료되었습니다.")
    else:
        st.warning("RESEND_API_KEY가 설정되지 않았습니다.")

    st.text_input(
        "기본 수신자",
        value=get_config_value("RESEND_TO_EMAIL", DEFAULT_EMAIL_RECIPIENT),
        disabled=True,
    )
    st.text_input(
        "보내는 이메일",
        value=get_config_value("RESEND_FROM_EMAIL", "미설정"),
        disabled=True,
    )
    st.caption("이메일 발송은 Resend API를 통해 처리됩니다. 문의 사항은 관리자(pung@khu.ac.kr)에게 연락해 주시기 바랍니다.")


render_hero()
if st.session_state.get("show_issue_report_dialog", False):
    open_issue_report_dialog()
render_status_cards()

st.markdown(
    """
    <div class="block-note">
        필수 항목은 별표(*)로 표시되어 있습니다. 자동생성 버튼을 사용할 수 있으며, 생성된 내용은 사용자가 직접 수정할 수 있습니다.
    </div>
    """,
    unsafe_allow_html=True,
)


with st.expander("Step 1. 표지 정보", expanded=True):
    st.caption("보고서 표지와 기본 제목을 구성하는 정보입니다.")

    col1, col2, col3 = st.columns([2, 1, 1])

    with col1:
        conference_name = st.text_input("학회명 *", key="conference_name")

    with col2:
        trip_type = st.radio("출장 구분 *", ["국내", "국외"], horizontal=True, key="trip_type")

    with col3:
        report_date = st.date_input("작성일", value=date.today(), key="report_date")

    title_preview = f"{conference_name} 참석 {trip_type} 출장 결과보고서" if conference_name else "학회명 참석 출장 결과보고서"
    st.text_input("표지 제목 자동 생성", value=title_preview, disabled=True)


with st.expander("Step 2. 학회 Overview 및 출장목적", expanded=True):
    st.caption("학회 개요 자료를 업로드하고 출장목적을 자동생성하거나 직접 작성합니다.")

    overview_files = st.file_uploader(
        "학회 Overview / Program Overview 업로드 *",
        type=["pdf", "docx", "txt", "md", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="overview_files",
    )

    col1, col2 = st.columns(2)

    with col1:
        start_date = st.date_input("출장기간 시작일 *", value=date.today(), key="start_date")

    with col2:
        end_date = st.date_input("출장기간 종료일 *", value=date.today(), key="end_date")

    trip_dates = date_range(start_date, end_date)
    ensure_schedule_for_dates(trip_dates)

    st.info(f"출장기간: {korean_date(start_date)} ~ {korean_date(end_date)} / 총 {len(trip_dates)}일")

    if st.button("Overview 텍스트 추출", key="extract_overview_btn"):
        run_id = f"{date.today().isoformat()}_{slugify(conference_name)}"
        paths = save_uploaded_files(overview_files, UPLOAD_ROOT / run_id / "overview", "overview")
        texts = [extract_text_from_file(path) for path in paths]
        extracted = "\n\n".join([t for t in texts if t.strip()])
        st.session_state["overview_paths_saved"] = paths

        if extracted:
            st.session_state["overview_text"] = extracted
            st.success("Overview 텍스트를 추출했습니다.")
        else:
            st.warning("텍스트 추출 결과가 없습니다. 직접 입력할 수 있습니다.")

        st.rerun()

    st.text_area(
        "추출된 Overview 텍스트 또는 직접 입력",
        height=180,
        key="overview_text",
        placeholder="여기에 직접 입력하거나, 위에서 Overview 파일을 업로드한 뒤 텍스트 추출을 누르세요.",
    )

    if st.button("출장목적 자동생성", key="generate_purpose_btn"):
        prompt = f"""
학회명: {conference_name}
출장 구분: {trip_type}
출장기간: {korean_date(start_date)} ~ {korean_date(end_date)}

학회 Overview:
{st.session_state.get("overview_text", "")}
"""
        try:
            output = get_llm().generate(TRIP_PURPOSE_SYSTEM, prompt, temperature=0.25, max_output_tokens=1000)
            st.session_state["purpose_text"] = output
            st.success("출장목적을 자동생성했습니다.")
            st.rerun()
        except Exception as exc:
            st.error(f"출장목적 자동생성 실패: {exc}")

    st.text_area(
        "출장목적 *",
        height=140,
        key="purpose_text",
        placeholder="자동생성하거나 직접 입력하세요.",
    )


with st.expander("Step 3. 출장자 / 출장지 / 학회장소 / 세부일정", expanded=True):
    st.caption("출장자, 장소, 일자별 세부일정을 입력합니다.")

    col1, col2, col3 = st.columns(3)

    with col1:
        traveler_name = st.text_input("출장자 성명 *", key="traveler_name")

    with col2:
        destination = st.text_input("출장지 *", key="destination", placeholder="예: 이탈리아 로마")

    with col3:
        venue = st.text_input("학회장소 *", key="venue", placeholder="예: Rome Convention Center")

    if st.button("세부일정 전체 자동생성", key="generate_schedule_btn"):
        prompt = f"""
학회명: {conference_name}
출장 구분: {trip_type}
출장기간: {korean_date(start_date)} ~ {korean_date(end_date)}
날짜 목록: {[iso_date(d) for d in trip_dates]}

학회 Overview:
{st.session_state.get("overview_text", "")}
"""
        try:
            output = get_llm().generate(TRIP_SCHEDULE_SYSTEM, prompt, temperature=0.2, max_output_tokens=1600)
            schedule = parse_schedule_json(output, trip_dates)
            apply_schedule_to_widgets(schedule)
            st.success("세부일정을 자동생성했습니다.")
            st.rerun()
        except Exception as exc:
            st.error(f"세부일정 자동생성 실패: {exc}")

    st.markdown("#### 세부일정")

    for i, d in enumerate(trip_dates):
        col_day, col_content = st.columns([1, 4])

        with col_day:
            st.text_input(
                "일차 / 날짜",
                value=f"{i + 1}일차 / {korean_date(d)}",
                disabled=True,
                key=f"day_label_{i}",
            )

        with col_content:
            st.text_area(
                "출장내용",
                height=70,
                key=f"schedule_content_{i}",
                placeholder="직접 입력하거나 자동생성하세요. 하루 100자 이내 권장.",
            )


with st.expander("Step 4. 본 연구와 관련성 및 주요 세션 요약", expanded=True):
    st.caption("연구과제와 학회 내용의 관련성을 정리합니다.")

    research_theme = st.text_input("연구과제명", key="research_theme")

    st.text_area(
        "직접 작성 내용 선택 입력",
        height=100,
        key="user_research_summary",
        placeholder="직접 쓴 요약이 있으면 여기에 입력하세요. 자동생성 시 반영됩니다.",
    )

    if st.button("본 연구와 관련성 및 주요 세션 요약 자동생성", key="generate_relevance_btn"):
        prompt = f"""
학회명: {conference_name}
출장 구분: {trip_type}
본인 연구 주제: {research_theme}

학회 Overview:
{st.session_state.get("overview_text", "")}

사용자 직접 입력 내용:
{st.session_state.get("user_research_summary", "")}
"""
        try:
            output = get_llm().generate(TRIP_RESEARCH_RELEVANCE_SYSTEM, prompt, temperature=0.25, max_output_tokens=2000)
            st.session_state["research_relatedness"] = output
            st.success("본 연구와 관련성 및 주요 세션 요약을 자동생성했습니다.")
            st.rerun()
        except Exception as exc:
            st.error(f"요약 자동생성 실패: {exc}")

    st.text_area(
        "본 연구와 관련성 및 주요 세션 요약 *",
        height=260,
        key="research_relatedness",
        placeholder="자동생성하거나 직접 입력하세요.",
    )


with st.expander("Step 5. 항공권 / 전자티켓 / 탑승권 관련 자료", expanded=False):
    st.caption("항공권, 탑승권, 영수증, 초청장 등 이동 관련 증빙을 업로드합니다.")

    e_ticket_files = st.file_uploader(
        "전자티켓 업로드 선택",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="e_ticket_files",
    )
    boarding_pass_files = st.file_uploader(
        "탑승권 업로드 *",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="boarding_pass_files",
    )
    ticket_receipt_files = st.file_uploader(
        "티켓 영수증 업로드 *",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="ticket_receipt_files",
    )
    acceptance_letter_files = st.file_uploader(
        "초청장 / Acceptance Letter 업로드 선택",
        type=["pdf", "docx", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="acceptance_letter_files",
    )

    use_research_card = st.checkbox("연구비카드에서 지출함", value=True, key="use_research_card")
    personal_card_reason = ""

    if not use_research_card:
        personal_card_reason = st.text_area("개인카드 사용 사유", key="personal_card_reason", height=100)


with st.expander("Step 6. 일자별 영수증 업로드", expanded=False):
    st.caption("출장기간 각 날짜별 영수증을 업로드합니다.")

    receipt_uploads = {}

    for i, d in enumerate(trip_dates):
        key = iso_date(d)
        receipt_uploads[key] = st.file_uploader(
            f"{i + 1}일차 / {korean_date(d)} 영수증 *",
            type=["pdf", "png", "jpg", "jpeg"],
            accept_multiple_files=True,
            key=f"receipt_files_{key}",
        )


with st.expander("Step 7. 일자별 출장 사진 업로드", expanded=False):
    st.caption("일자별 현장 방문 또는 학회 참석 사진을 업로드합니다.")

    skip_first_photo = st.checkbox(
        "국외 출장의 학회 시작 전날 이동일 사진 생략 적용",
        value=False,
        disabled=(trip_type != "국외"),
        key="skip_first_photo",
    )

    photo_uploads = {}

    for i, d in enumerate(trip_dates):
        key = iso_date(d)
        st.markdown(f"#### {i + 1}일차 / {korean_date(d)}")
        st.caption(f"출장내용: {st.session_state.get(f'schedule_content_{i}', '') or '미입력'}")

        photo_uploads[key] = st.file_uploader(
            f"{i + 1}일차 사진 *",
            type=["png", "jpg", "jpeg"],
            accept_multiple_files=True,
            key=f"photo_files_{key}",
        )


with st.expander("Step 8. 숙박확인서 / 숙박 영수증", expanded=False):
    st.caption("숙박확인서와 숙박 영수증을 업로드합니다.")

    lodging_confirmation_files = st.file_uploader(
        "숙박확인서 업로드 *",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="lodging_confirmation_files",
    )
    lodging_receipt_files = st.file_uploader(
        "숙박 영수증 업로드 *",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="lodging_receipt_files",
    )


reason_enabled = False

with st.expander("Step 9. 추가 제출서류 및 사유서", expanded=False):
    st.caption("추가 제출서류와 필요 시 사유서를 생성합니다.")

    conference_intro_files = st.file_uploader(
        "학회 소개자료 업로드 선택",
        type=["pdf", "docx", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="conference_intro_files",
    )
    registration_statement_files = st.file_uploader(
        "등록비 비용 명세 업로드 선택",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="registration_statement_files",
    )
    registration_invoice_files = st.file_uploader(
        "청구서 및 인보이스/영수증 업로드 선택",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="registration_invoice_files",
    )
    acknowledgement_paper_files = st.file_uploader(
        "사사 acknowledgement 기재 논문 원본 업로드 선택",
        type=["pdf", "docx"],
        accept_multiple_files=True,
        key="acknowledgement_paper_files",
    )

    st.divider()

    reason_enabled = st.checkbox("사유서 생성", value=(not use_research_card), key="reason_enabled")

    if reason_enabled:
        col1, col2 = st.columns(2)

        with col1:
            principal_affiliation = st.text_input("연구책임자 소속", key="principal_affiliation")
            funding_agency = st.text_input("지원기관", key="funding_agency")
            project_title = st.text_input("연구과제명", key="project_title")

        with col2:
            principal_name = st.text_input("연구책임자 성명", key="principal_name")
            research_period = st.text_input(
                "당해연도 연구기간",
                key="research_period",
                placeholder="예: 2026.03.01 ~ 2027.02.28",
            )

        st.text_area(
            "사유서 작성 참고 내용",
            key="reason_context",
            height=100,
            placeholder="예: 개인카드 사용 사유, 증빙자료 누락/대체 제출 사유 등",
        )

        if st.button("사유서 제목/내용 자동생성", key="generate_reason_btn"):
            prompt = f"""
학회명: {conference_name}
출장 구분: {trip_type}
출장기간: {korean_date(start_date)} ~ {korean_date(end_date)}
연구비카드 사용 여부: {use_research_card}
개인카드 사용 사유: {personal_card_reason}
사용자 참고 내용: {st.session_state.get("reason_context", "")}
"""
            try:
                output = get_llm().generate(REASON_STATEMENT_SYSTEM, prompt, temperature=0.25, max_output_tokens=1200)
                parsed = json.loads(output[output.find("{"): output.rfind("}") + 1])
                st.session_state["reason_title"] = parsed.get("title", "")
                st.session_state["reason_content"] = parsed.get("content", "")
                st.success("사유서 제목/내용을 자동생성했습니다.")
                st.rerun()
            except Exception as exc:
                st.error(f"사유서 자동생성 실패: {exc}")

        st.text_input("제 목", key="reason_title")
        st.text_area("내 용", key="reason_content", height=180)


def build_current_data(save_files: bool = False) -> TripReportData:
    schedule = sync_schedule_from_widgets(trip_dates)

    run_id = f"{date.today().isoformat()}_{slugify(conference_name)}"
    root = UPLOAD_ROOT / run_id

    if save_files:
        overview_paths = save_uploaded_files(overview_files, root / "overview", "overview")

        if not overview_paths:
            overview_paths = st.session_state.get("overview_paths_saved", [])

        transport_files = {
            "e_ticket": save_uploaded_files(e_ticket_files, root / "transport", "e_ticket"),
            "boarding_pass": save_uploaded_files(boarding_pass_files, root / "transport", "boarding_pass"),
            "ticket_receipt": save_uploaded_files(ticket_receipt_files, root / "transport", "ticket_receipt"),
            "acceptance_letter": save_uploaded_files(acceptance_letter_files, root / "transport", "acceptance_letter"),
        }
        daily_receipts = {
            k: save_uploaded_files(v, root / "daily_receipts" / k, f"receipt_{k}")
            for k, v in receipt_uploads.items()
        }
        daily_photos = {
            k: save_uploaded_files(v, root / "daily_photos" / k, f"photo_{k}")
            for k, v in photo_uploads.items()
        }
        lodging_files = {
            "lodging_confirmation": save_uploaded_files(
                lodging_confirmation_files,
                root / "lodging",
                "lodging_confirmation",
            ),
            "lodging_receipt": save_uploaded_files(
                lodging_receipt_files,
                root / "lodging",
                "lodging_receipt",
            ),
        }
        extra_files = {
            "conference_intro": save_uploaded_files(conference_intro_files, root / "extra", "conference_intro"),
            "registration_statement": save_uploaded_files(
                registration_statement_files,
                root / "extra",
                "registration_statement",
            ),
            "registration_invoice": save_uploaded_files(
                registration_invoice_files,
                root / "extra",
                "registration_invoice",
            ),
            "acknowledgement_paper": save_uploaded_files(
                acknowledgement_paper_files,
                root / "extra",
                "acknowledgement_paper",
            ),
        }
    else:
        overview_paths = st.session_state.get("overview_paths_saved", []) or (["uploaded"] if overview_files else [])
        transport_files = {
            "e_ticket": ["uploaded"] if e_ticket_files else [],
            "boarding_pass": ["uploaded"] if boarding_pass_files else [],
            "ticket_receipt": ["uploaded"] if ticket_receipt_files else [],
            "acceptance_letter": ["uploaded"] if acceptance_letter_files else [],
        }
        daily_receipts = {k: (["uploaded"] if v else []) for k, v in receipt_uploads.items()}
        daily_photos = {k: (["uploaded"] if v else []) for k, v in photo_uploads.items()}
        lodging_files = {
            "lodging_confirmation": ["uploaded"] if lodging_confirmation_files else [],
            "lodging_receipt": ["uploaded"] if lodging_receipt_files else [],
        }
        extra_files = {
            "conference_intro": ["uploaded"] if conference_intro_files else [],
            "registration_statement": ["uploaded"] if registration_statement_files else [],
            "registration_invoice": ["uploaded"] if registration_invoice_files else [],
            "acknowledgement_paper": ["uploaded"] if acknowledgement_paper_files else [],
        }

    reason = ReasonStatementData(
        enabled=reason_enabled,
        principal_affiliation=st.session_state.get("principal_affiliation", ""),
        principal_name=st.session_state.get("principal_name", ""),
        funding_agency=st.session_state.get("funding_agency", ""),
        research_period=st.session_state.get("research_period", ""),
        project_title=st.session_state.get("project_title", ""),
        reason_context=st.session_state.get("reason_context", ""),
        generated_title=st.session_state.get("reason_title", ""),
        generated_content=st.session_state.get("reason_content", ""),
    )

    return TripReportData(
        conference_name=conference_name,
        trip_type=trip_type,
        report_date=iso_date(report_date),
        start_date=iso_date(start_date),
        end_date=iso_date(end_date),
        overview_paths=overview_paths,
        overview_text=st.session_state.get("overview_text", ""),
        purpose_text=st.session_state.get("purpose_text", ""),
        traveler_name=traveler_name,
        destination=destination,
        venue=venue,
        daily_schedule=schedule,
        research_theme=research_theme,
        research_relatedness=st.session_state.get("research_relatedness", ""),
        expected_effect_text=st.session_state.get("expected_effect_text", ""),
        transport_files=transport_files,
        daily_receipts=daily_receipts,
        daily_photos=daily_photos,
        lodging_files=lodging_files,
        extra_files=extra_files,
        use_research_card=use_research_card,
        personal_card_reason=personal_card_reason,
        skip_first_photo_for_international=skip_first_photo,
        reason_statement=reason,
    )

notice = st.session_state.get("last_generation_notice", "")
if notice:
    if "누락" in notice or "확인 필요" in notice:
        st.warning(notice)
    else:
        st.success(notice)
    st.session_state["last_generation_notice"] = ""

last_zip_path = st.session_state.get("last_generated_zip", "")
email_enabled = bool(last_zip_path and Path(last_zip_path).exists())

action_cols = st.columns(2)

with action_cols[0]:
    generate_clicked = st.button(
        "출장보고서 생성",
        type="primary",
        key="generate_report_btn",
        use_container_width=True,
    )

with action_cols[1]:
    email_clicked = st.button(
        "서류 발송",
        key="open_zip_email_dialog_btn",
        disabled=not email_enabled,
        use_container_width=True,
    )

if not email_enabled:
    st.caption("보고서 생성 전에는 서류 발송 버튼이 비활성화됩니다.")

if email_clicked:
    open_zip_email_dialog()

if generate_clicked:
    try:
        with st.spinner("기대효과 자동생성 중입니다."):
            try:
                st.session_state["expected_effect_text"] = generate_expected_effect_text()
            except Exception:
                st.session_state["expected_effect_text"] = get_default_expected_effect()

        final_data = build_current_data(save_files=True)
        missing = check_missing_documents(final_data)
        st.session_state["last_missing"] = missing
        st.session_state["last_checklist"] = checklist_rows(final_data)

        run_id = f"{date.today().isoformat()}_{slugify(conference_name)}"
        out_dir = OUTPUT_ROOT / run_id
        out_dir.mkdir(parents=True, exist_ok=True)

        base_name = slugify(conference_name or "Trip_Report")
        docx_path = out_dir / f"{base_name}_출장결과보고서.docx"
        pdf_path = out_dir / f"{base_name}_출장결과보고서.pdf"
        zip_path = out_dir / f"{base_name}_출장결과보고서_패키지.zip"

        generated_docx = generate_trip_docx(final_data, docx_path)
        generated_pdf = generate_trip_pdf(final_data, pdf_path)

        reason_path = None

        if final_data.reason_statement.enabled:
            template_path = APP_ROOT / "templates" / "사유서_template.pdf"
            reason_path = generate_reason_pdf_on_template(
                final_data.reason_statement,
                out_dir / f"{base_name}_사유서.pdf",
                template_path,
                final_data.report_date,
            )

        generated_zip = create_zip(final_data, zip_path, generated_docx, generated_pdf, reason_path)

        st.session_state["last_generated_docx"] = str(generated_docx)
        st.session_state["last_generated_pdf"] = str(generated_pdf)
        st.session_state["last_generated_reason_pdf"] = str(reason_path) if reason_path else ""
        st.session_state["last_generated_zip"] = str(generated_zip)

        if missing:
            st.session_state["last_generation_notice"] = "보고서는 생성되었지만 누락 또는 확인 필요 항목이 있습니다. 아래 체크리스트를 확인하세요."
        else:
            st.session_state["last_generation_notice"] = "출장보고서 생성이 완료되었습니다."

        st.rerun()

    except Exception as exc:
        st.error(f"보고서 생성 실패: {exc}")

render_generated_downloads()

if st.session_state.get("last_checklist"):
    st.markdown("#### 제출서류 체크리스트")
    df = pd.DataFrame(st.session_state["last_checklist"])
    st.dataframe(style_checklist(df), use_container_width=True)

if st.session_state.get("last_missing"):
    st.error("누락 또는 확인 필요 항목")
    for item in st.session_state["last_missing"]:
        st.write(f"- {item}")
