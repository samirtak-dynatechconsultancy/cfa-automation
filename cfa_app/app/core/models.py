"""Data holders and detection configuration for the CFA transfer engine."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class DetectionConfig:
    """Header-anchor detection settings. Mirrors the CONFIG block of the original script.

    These are defaults; the admin page can override them (persisted in settings.xlsx).
    """

    source_tab: str = "G"
    form_header: str = "FORM"
    line_header: str = "LINE"
    check_header: str = "EQUAL TO"
    diff_header: str = "Absolute difference"   # source column with the |difference| per line
    diff_threshold: float = 5000.0             # |difference| above this -> highlight the written cell
    header_scan_rows: int = 20
    entity_row_scan: int = 20
    max_data_rows: int = 400
    target_name_token: str = "verification"

    @classmethod
    def from_mapping(cls, data: dict | None) -> "DetectionConfig":
        """Build a config from a (possibly partial) mapping, ignoring unknown keys."""
        if not data:
            return cls()
        known = {f for f in cls.__dataclass_fields__}  # type: ignore[attr-defined]
        clean = {}
        for k, v in data.items():
            key = str(k).strip().lower()
            if key in known and v is not None and str(v).strip() != "":
                clean[key] = v
        # coerce ints
        for int_key in ("header_scan_rows", "entity_row_scan", "max_data_rows"):
            if int_key in clean:
                try:
                    clean[int_key] = int(clean[int_key])
                except (TypeError, ValueError):
                    del clean[int_key]
        if "diff_threshold" in clean:                       # coerce the numeric threshold
            try:
                clean["diff_threshold"] = float(clean["diff_threshold"])
            except (TypeError, ValueError):
                del clean["diff_threshold"]
        return cls(**clean)

    def to_mapping(self) -> dict:
        return {f: getattr(self, f) for f in self.__dataclass_fields__}  # type: ignore[attr-defined]


# Values that legitimately appear in the check column (informational only; copied verbatim).
KNOWN_CHECK_VALUES = {"OK", "NOT EQUAL", "-----"}


@dataclass
class SourceData:
    filename: str
    entity: str
    month: int
    year: int
    sheet_name: str
    header_row: int          # 1-based
    form_col: int            # 1-based
    line_col: int            # 1-based
    check_col: int           # 1-based
    rows: list = field(default_factory=list)  # list of (form, line, check_value, abs_diff), data order
    currency: str = "USD"    # from the filename; drives which period sheet the values land in
    company: str = ""        # 'COMPANY:' name from the source (written under a new entity column)
    source_url: str = ""     # SharePoint webUrl of the source (hyperlinked on the entity number)


@dataclass
class FileResult:
    filename: str
    entity: str | None = None
    period: str | None = None
    currency: str = "USD"
    path: str = ""               # SharePoint folder path the source came from
    sheet: str | None = None     # target sheet the values were written to
    status: str = "skipped"      # "done" | "skipped" | "error"
    written: int = 0
    skipped_lines: int = 0
    entity_added: bool = False   # True if a new column had to be created for this entity
    messages: list = field(default_factory=list)


@dataclass
class VerifyResult:
    checked: int = 0
    mismatches: int = 0
    samples: list = field(default_factory=list)   # up to N "SHEET!CELL expected/got" strings

    @property
    def passed(self) -> bool:
        return self.mismatches == 0


@dataclass
class RunResult:
    run_id: str
    status: str = "queued"        # queued | running | done | error
    message: str = ""
    output_name: str | None = None
    output_url: str | None = None
    files: list = field(default_factory=list)     # list[FileResult]
    verify: VerifyResult | None = None
    started_at: str | None = None
    finished_at: str | None = None
    stage: str = ""                               # current plain-language activity (replaced live)
    processed: int = 0                            # entity files done so far
    total: int = 0                                # entity files to do (0 until known)
    logs: list = field(default_factory=list)      # technical detail lines (collapsible)
    notices: list = field(default_factory=list)   # notable events (e.g. entity column added)

    def as_dict(self) -> dict:
        return {
            "run_id": self.run_id,
            "status": self.status,
            "message": self.message,
            "output_name": self.output_name,
            "output_url": self.output_url,
            "started_at": self.started_at,
            "finished_at": self.finished_at,
            "stage": self.stage,
            "processed": self.processed,
            "total": self.total,
            "logs": self.logs,
            "notices": self.notices,
            "verify": None
            if self.verify is None
            else {
                "passed": self.verify.passed,
                "checked": self.verify.checked,
                "mismatches": self.verify.mismatches,
                "samples": self.verify.samples,
            },
            "files": [
                {
                    "filename": f.filename,
                    "entity": f.entity,
                    "period": f.period,
                    "currency": f.currency,
                    "path": f.path,
                    "sheet": f.sheet,
                    "status": f.status,
                    "written": f.written,
                    "skipped_lines": f.skipped_lines,
                    "messages": f.messages,
                }
                for f in self.files
            ],
        }
