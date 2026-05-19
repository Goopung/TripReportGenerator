from core.data_model import TripReportData


def check_missing_documents(data: TripReportData) -> list[str]:
    missing: list[str] = []

    if not data.conference_name:
        missing.append("학회명")
    if not data.trip_type:
        missing.append("출장 구분")
    if not data.overview_paths and not data.overview_text:
        missing.append("학회 Overview / Program Overview")
    if not data.start_date or not data.end_date:
        missing.append("출장기간")
    if not data.purpose_text:
        missing.append("출장목적")
    if not data.traveler_name:
        missing.append("출장자 성명")
    if not data.destination:
        missing.append("출장지")
    if not data.venue:
        missing.append("학회장소")
    if not data.research_relatedness:
        missing.append("본 연구와 관련성 및 주요 세션 요약")

    if not data.transport_files.get("boarding_pass"):
        missing.append("탑승권")
    if not data.transport_files.get("ticket_receipt"):
        missing.append("티켓 영수증")

    if not data.use_research_card and not data.personal_card_reason and not data.reason_statement.enabled:
        missing.append("개인카드 사용 사유서")

    for idx, row in enumerate(data.daily_schedule):
        date_key = row.get("date", "")
        day_label = f"{row.get('day', idx + 1)}일차 {date_key}"

        if not data.daily_receipts.get(date_key):
            missing.append(f"{day_label} 영수증")

        skip_photo = (
            data.trip_type == "국외"
            and data.skip_first_photo_for_international
            and idx == 0
        )
        if not skip_photo and not data.daily_photos.get(date_key):
            missing.append(f"{day_label} 출장 사진")

    if not data.lodging_files.get("lodging_confirmation"):
        missing.append("숙박확인서")
    if not data.lodging_files.get("lodging_receipt"):
        missing.append("숙박 영수증")

    if data.reason_statement.enabled:
        rs = data.reason_statement
        if not rs.principal_affiliation:
            missing.append("사유서: 연구책임자 소속")
        if not rs.principal_name:
            missing.append("사유서: 연구책임자 성명")
        if not rs.funding_agency:
            missing.append("사유서: 지원기관")
        if not rs.research_period:
            missing.append("사유서: 당해연도 연구기간")
        if not rs.project_title:
            missing.append("사유서: 연구과제명")
        if not rs.generated_title or not rs.generated_content:
            missing.append("사유서: 제 목 / 내 용")

    return missing


def checklist_rows(data: TripReportData) -> list[dict[str, str]]:
    missing = set(check_missing_documents(data))

    def done(condition: bool, required: bool = True) -> str:
        if condition:
            return "완료"
        return "누락" if required else "선택"

    rows = [
        {"구분": "표지", "서류명": "학회명", "상태": done(bool(data.conference_name))},
        {"구분": "표지", "서류명": "출장 구분", "상태": done(bool(data.trip_type))},
        {"구분": "본문", "서류명": "학회 Overview / Program Overview", "상태": done(bool(data.overview_paths or data.overview_text))},
        {"구분": "본문", "서류명": "출장목적", "상태": done(bool(data.purpose_text))},
        {"구분": "본문", "서류명": "본 연구와 관련성 및 주요 세션 요약", "상태": done(bool(data.research_relatedness))},
        {"구분": "항공권", "서류명": "티켓 영수증", "상태": done(bool(data.transport_files.get("ticket_receipt")))},
        {"구분": "항공권", "서류명": "전자 티켓", "상태": done(bool(data.transport_files.get("e_ticket")), required=False)},
        {"구분": "항공권", "서류명": "초청장 / Acceptance Letter", "상태": done(bool(data.transport_files.get("acceptance_letter")), required=False)},
        {"구분": "항공권", "서류명": "탑승권", "상태": done(bool(data.transport_files.get("boarding_pass")))},
        {"구분": "학회 등록비", "서류명": "학회 소개자료", "상태": done(bool(data.extra_files.get("conference_intro")), required=False)},
        {"구분": "학회 등록비", "서류명": "등록비 비용 명세", "상태": done(bool(data.extra_files.get("registration_statement")), required=False)},
        {"구분": "학회 등록비", "서류명": "청구서 및 인보이스/영수증", "상태": done(bool(data.extra_files.get("registration_invoice")), required=False)},
        {"구분": "여비 청구", "서류명": "각 일자별 영수증", "상태": "누락" if any("영수증" in x and "일차" in x for x in missing) else "완료"},
        {"구분": "여비 청구", "서류명": "각 일자별 사진", "상태": "누락" if any("출장 사진" in x for x in missing) else "완료"},
        {"구분": "여비 청구", "서류명": "숙박확인서", "상태": done(bool(data.lodging_files.get("lodging_confirmation")))},
        {"구분": "여비 청구", "서류명": "숙박 영수증", "상태": done(bool(data.lodging_files.get("lodging_receipt")))},
        {"구분": "선택 자료", "서류명": "사사 acknowledgement 기재 논문 원본", "상태": done(bool(data.extra_files.get("acknowledgement_paper")), required=False)},
    ]

    if data.reason_statement.enabled or not data.use_research_card:
        rows.append(
            {
                "구분": "사유서",
                "서류명": "개인카드 사용 또는 증빙 관련 사유서",
                "상태": "완료" if data.reason_statement.generated_title and data.reason_statement.generated_content else "누락",
            }
        )

    return rows
