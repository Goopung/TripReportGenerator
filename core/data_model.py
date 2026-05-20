from dataclasses import dataclass, field


@dataclass
class ReasonStatementData:
    enabled: bool = False
    principal_affiliation: str = ""
    principal_name: str = ""
    funding_agency: str = ""
    research_period: str = ""
    project_title: str = ""
    reason_context: str = ""
    generated_title: str = ""
    generated_content: str = ""


@dataclass
class TripReportData:
    conference_name: str
    trip_type: str
    report_date: str
    start_date: str
    end_date: str
    overview_paths: list[str] = field(default_factory=list)
    overview_text: str = ""
    purpose_text: str = ""
    traveler_name: str = ""
    destination: str = ""
    venue: str = ""
    daily_schedule: list[dict] = field(default_factory=list)
    research_theme: str = ""
    research_relatedness: str = ""
    expected_effect_text: str = ""
    transport_files: dict[str, list[str]] = field(default_factory=dict)
    daily_receipts: dict[str, list[str]] = field(default_factory=dict)
    daily_photos: dict[str, list[str]] = field(default_factory=dict)
    lodging_files: dict[str, list[str]] = field(default_factory=dict)
    extra_files: dict[str, list[str]] = field(default_factory=dict)
    use_research_card: bool = True
    personal_card_reason: str = ""
    skip_first_photo_for_international: bool = False
    reason_statement: ReasonStatementData = field(default_factory=ReasonStatementData)

    def to_dict(self) -> dict:
        return {
            "conference_name": self.conference_name,
            "trip_type": self.trip_type,
            "report_date": self.report_date,
            "start_date": self.start_date,
            "end_date": self.end_date,
            "overview_paths": self.overview_paths,
            "overview_text": self.overview_text,
            "purpose_text": self.purpose_text,
            "traveler_name": self.traveler_name,
            "destination": self.destination,
            "venue": self.venue,
            "daily_schedule": self.daily_schedule,
            "research_theme": self.research_theme,
            "research_relatedness": self.research_relatedness,
            "expected_effect_text": self.expected_effect_text,
            "transport_files": self.transport_files,
            "daily_receipts": self.daily_receipts,
            "daily_photos": self.daily_photos,
            "lodging_files": self.lodging_files,
            "extra_files": self.extra_files,
            "use_research_card": self.use_research_card,
            "personal_card_reason": self.personal_card_reason,
            "skip_first_photo_for_international": self.skip_first_photo_for_international,
            "reason_statement": self.reason_statement.__dict__,
        }
