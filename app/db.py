from __future__ import annotations

import json
import os
import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = Path(os.getenv("ACCOUNTING_DATA_DIR", BASE_DIR / "data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA_DIR / "accounting_review.db"


def connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


@contextmanager
def db() -> Iterator[sqlite3.Connection]:
    conn = connect()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _ensure_column(conn: sqlite3.Connection, table: str, column: str, ddl: str) -> None:
    cols = {r[1] for r in conn.execute(f"PRAGMA table_info({table})")}
    if column not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {ddl}")


DEFAULT_SETTINGS = {
    "program_name": "회계자료 검토 및 보고 보조 프로그램",
    "report_company_name": "회사명 미설정",
    "ai_mode": "local",
}

DEFAULT_MENUS = [
    ("dashboard", "/", "검토 현황", 1, 10),
    ("upload", "/upload", "자료 업로드", 1, 20),
    ("transactions", "/transactions", "회계자료 검토", 1, 30),
    ("issues", "/issues", "확인·보완 업무", 1, 35),
    ("documents", "/documents", "증빙·문서 관리", 1, 40),
    ("checklist", "/checklist", "제출 전 체크리스트", 1, 50),
    ("reports", "/reports", "월별 보고", 1, 60),
    ("ai", "/ai-assist", "AI 보고 보조", 1, 70),
    ("settings", "/settings", "관리자 설정", 1, 80),
]

DEFAULT_UI_TEXTS = {
    # 공통/사이드바
    "brand_subtitle": "Accounting Review Assistant",
    "sidebar_mode_title": "로컬 안전 모드",
    "sidebar_mode_desc": "계산·검증은 규칙 기반으로 수행하며, 외부 AI로 자료를 자동 전송하지 않습니다.",
    # 대시보드
    "dashboard_description": "업로드된 회계자료의 품질, 검토 필요 항목, 증빙 누락과 처리대상을 한눈에 확인합니다.",
    "dashboard_upload_button": "+ 회계자료 업로드",
    "dashboard_excel_button": "검토결과 Excel",
    "metric_total": "전체 자료",
    "metric_total_sub": "업로드 {count}회",
    "metric_review": "검토 필요",
    "metric_review_sub": "정상 {count}건",
    "metric_evidence": "증빙 누락",
    "metric_evidence_sub": "증빙번호·문서 연결 기준",
    "metric_duplicate": "중복 의심",
    "metric_duplicate_sub": "활성 중복 Rule 기준",
    "metric_unassigned": "담당자 미지정",
    "metric_overdue": "처리기한 경과",
    "dashboard_quality_title": "데이터 품질",
    "dashboard_issue_title": "미완료 이슈 구성",
    "dashboard_recent_tx_title": "최근 회계자료",
    "dashboard_recent_import_title": "최근 업로드",
    "dashboard_empty_issues": "미완료 이슈가 없습니다.",
    "dashboard_empty_tx": "아직 업로드된 자료가 없습니다.",
    "dashboard_empty_import": "업로드 이력이 없습니다.",
    "dashboard_footer": "검토결과는 회계·세무·감사인의 전문적 판단을 대체하지 않습니다. 프로그램은 누락·중복·형식 오류와 확인 필요 항목을 좁히는 보조 도구입니다.",
    # 업로드
    "upload_title": "회계자료 업로드",
    "upload_description": "CSV/XLSX를 읽고 열 이름을 자동 인식한 뒤, 업로드 전 사용자가 매핑을 확인합니다.",
    "upload_sample_button": "표준양식 다운로드",
    "upload_demo_button": "시연용 데모자료",
    "upload_file_label": "회계자료 파일",
    "upload_file_help": "권장: 첫 행에 열 제목이 있는 표 형식. 최대 30MB.",
    "upload_preview_button": "1. 열 자동인식 및 미리보기",
    "upload_process_title": "정확성을 위해 2단계로 업로드합니다",
    "upload_process_1": "파일을 읽어 열 이름을 자동 인식합니다.",
    "upload_process_2": "지출일자·금액·사용목적 등 실제 연결 열을 사용자가 직접 확인합니다.",
    "upload_process_3": "확정 후 데이터베이스에 저장하고 활성 검토 Rule을 실행합니다.",
    # 거래 검토
    "transactions_description": "원본 파일과 행번호를 보존한 상태로 검토 결과와 처리상태를 확인합니다.",
    "transactions_upload_button": "+ 자료 업로드",
    "transactions_search_placeholder": "거래처·목적·전표번호 검색",
    "transactions_search_button": "조회",
    # 확인·보완 업무
    "issues_description": "검토에서 발견된 확인·보완 사항을 담당자, 기한, 상태와 함께 실제 처리업무로 관리합니다.",
    "issues_filter_button": "업무 조회",
    "issues_empty": "조건에 해당하는 확인·보완 업무가 없습니다.",
    "detail_title": "거래 검토 상세",
    "detail_back_button": "목록으로",
    "detail_data_title": "회계자료",
    "detail_documents_title": "연결된 증빙문서",
    "detail_issues_title": "자동 검토 결과",
    "detail_request_summary": "확인 요청 문구",
    "detail_save_button": "처리 저장",
    # 확인 요청 문구 템플릿
    "confirm_greeting": "안녕하세요, {name}님.",
    "confirm_closing": "확인 감사합니다.",
    "confirm_evidence": "{date} {vendor} {amount} 지출 건의 증빙자료가 확인되지 않습니다. 영수증·세금계산서 등 증빙 제출 여부를 확인 부탁드립니다.",
    "confirm_duplicate": "{date} {vendor} {amount} 건이 중복 또는 분할 결제로 보이는 유사 거래와 함께 확인됩니다. 실제 별도 지출인지 확인 부탁드립니다.",
    "confirm_purpose": "{date} {vendor} {amount} 지출 건의 사용목적이 부족합니다. 어떤 업무와 관련된 비용인지 구체적인 사용목적을 회신 부탁드립니다.",
    "confirm_vat": "{date} {vendor} {amount} 건의 공급가액·부가세 합계와 총액이 일치하지 않습니다. 원 증빙과 입력금액을 다시 확인 부탁드립니다.",
    "confirm_default": "{date} {vendor} {amount} 지출 건에 확인사항이 있습니다: {message}. 확인 후 회신 부탁드립니다.",
    # 체크리스트
    "checklist_description": "제출 전에 사람이 최종 확인해야 할 항목을 자동·수동 체크리스트로 관리합니다.",
    "checklist_validate_button": "검토 규칙 다시 실행",
    "checklist_audit_button": "감사대응 패키지",
    "checklist_ready_title": "제출 준비 상태:",
    "checklist_ready_text": "핵심 체크리스트가 완료되었습니다. 최종 승인자는 원본 증빙과 예외처리 근거를 확인하세요.",
    "checklist_not_ready_title": "추가 검토 필요:",
    "checklist_not_ready_text": "제출 전 해결 또는 예외근거 기록이 필요한 항목이 남아 있습니다.",
    "checklist_auto_title": "제출 체크리스트",
    "checklist_status_title": "현황",
    "checklist_unresolved_title": "미완료 항목",
    # 보고
    "reports_description": "월별 회계자료를 수치 기반으로 요약하고, 전월 변화의 주요 기여 항목과 보고 초안을 만듭니다.",
    "reports_excel_button": "Excel 보고서",
    "reports_audit_button": "감사대응 패키지",
    "reports_draft_title": "월별 보고 초안",
    "reports_draft_note": "숫자는 프로그램 계산 결과이며, 문장은 검증된 집계값을 템플릿으로 정리한 것입니다.",
    "reports_check_title": "핵심 확인항목",
    "reports_account_title": "계정과목별 금액",
    "reports_vendor_title": "거래처별 금액",
    "reports_change_title": "전월 대비 주요 증감 기여",
    # 문서
    "documents_description": "증빙번호를 부여해 거래와 연결하고, 디지털 문서는 원문 텍스트를 발췌해 분류·요약합니다.",
    "documents_upload_title": "문서 업로드",
    "documents_principle_title": "자동화 원칙",
    "documents_principle_desc": "문서 요약은 원문 발췌 방식으로 처리하며, 증빙번호가 회계자료와 일치하면 증빙 누락 검토에 자동 반영됩니다.",
    "documents_image_notice": "이미지(JPG/PNG)는 현재 OCR 자동확정을 하지 않습니다. 이미지 문서의 본문은 원본 확인이 필요합니다.",
    # 매핑
    "mapping_title": "열 매핑 확인",
    "mapping_description": "업로드 파일의 열을 표준 회계 검토 항목과 연결합니다. 필수 항목은 반드시 확인하세요.",
    "mapping_warning": "자동인식 결과를 그대로 믿지 말고 실제 열 의미를 확인하세요. 잘못된 열 매핑은 이후 모든 검토 결과를 왜곡할 수 있습니다.",
    "mapping_confirm_button": "2. 매핑 확정 및 검토 실행",
    "mapping_preview_title": "원본 미리보기",
    # AI
    "ai_description": "검증된 집계값만 사용해 외부 생성형 AI에 전달할 보고 보조 프롬프트를 만듭니다.",
    "ai_privacy_notice": "프로그램이 외부 AI 서비스로 자료를 자동 전송하지 않습니다. 집계된 사실만 포함한 프롬프트를 사용자가 직접 확인해 활용합니다.",
    "ai_prompt_title": "AI 월간보고 보조 프롬프트",
    "ai_footer": "AI가 만든 문장은 반드시 사람이 원자료·증빙·검토이슈와 대조한 뒤 사용하세요. AI에게 숫자 계산, 회계처리 확정, 규정 위반 판정을 맡기지 않습니다.",
    # 설정
    "settings_description": "코드 수정 없이 화면 문구, 메뉴, 검토 Rule, 체크리스트와 운영기준을 변경합니다.",
    "settings_backup_button": "전체 데이터 백업",
}


# Streamlit portfolio view texts are also code-less. Existing v0.2 databases
# receive these automatically on the next init_db() run.
DEFAULT_UI_TEXTS.update({
    "st_dashboard_kicker": "01 / REVIEW OVERVIEW",
    "st_dashboard_title": "숫자보다 먼저, 근거를 봅니다.",
    "st_upload_kicker": "02 / DATA INTAKE",
    "st_upload_title": "자료를 먼저 깨끗하게.",
    "st_transactions_kicker": "03 / TRANSACTION REVIEW",
    "st_transactions_title": "원본 행까지 추적하는 검토.",
    "st_issues_kicker": "04 / ACTION QUEUE",
    "st_issues_title": "발견한 문제를 실제 업무로.",
    "st_documents_kicker": "05 / EVIDENCE",
    "st_documents_title": "증빙을 거래와 연결하기.",
    "st_checklist_kicker": "06 / PRE-SUBMISSION",
    "st_checklist_title": "제출 전 마지막 확인.",
    "st_reports_kicker": "07 / MONTHLY REPORT",
    "st_reports_title": "무엇이 변했는지부터.",
    "st_ai_kicker": "08 / AI ASSIST",
    "st_ai_title": "AI는 계산기가 아니라 문장 보조자.",
    "st_settings_kicker": "09 / NO-CODE ADMIN",
    "st_settings_title": "코드 없이 운영 기준을 바꾸기.",
    "st_help_kicker": "HOW TO USE",
    "st_help_title": "5분이면 흐름을 볼 수 있습니다.",
})

DEFAULT_RULES = [
    # code, name, enabled, rule_type, field, operator, compare, severity, category, message, order, system
    ("DATE_MISSING", "지출일자 누락", 1, "FIELD", "expense_date", "missing", "", "오류", "기본정보", "지출일자가 없거나 올바른 날짜 형식이 아닙니다.", 10, 1),
    ("DATE_FUTURE", "미래 일자 확인", 1, "DATE_FUTURE", "expense_date", "", "", "주의", "기본정보", "지출일자가 오늘 이후입니다. 예정 지출인지 확인하세요.", 20, 1),
    ("DATE_WEEKEND", "주말 지출 확인", 0, "DATE_WEEKEND", "expense_date", "", "", "확인", "기본정보", "주말 지출입니다. 회사 내부기준상 확인이 필요한지 검토하세요.", 25, 1),
    ("AMOUNT_MISSING", "금액 누락", 1, "FIELD", "amount", "missing", "", "오류", "금액", "금액이 없거나 숫자로 변환할 수 없습니다.", 30, 1),
    ("AMOUNT_ZERO", "0원 거래 확인", 1, "FIELD", "amount", "eq", "0", "주의", "금액", "금액이 0원입니다. 입력 오류 또는 취소 건인지 확인하세요.", 40, 1),
    ("AMOUNT_NEGATIVE", "음수 거래 확인", 1, "FIELD", "amount", "lt", "0", "주의", "금액", "음수 금액입니다. 환불·취소·수정전표 여부를 확인하세요.", 50, 1),
    ("AMOUNT_HIGH", "고액 지출 확인", 1, "FIELD", "amount", "gte", "1000000", "확인", "금액", "설정된 고액 기준({threshold}원) 이상입니다. 승인·증빙을 확인하세요.", 60, 1),
    ("PURPOSE_MISSING", "사용목적 누락", 1, "FIELD", "purpose", "missing", "", "오류", "사용목적", "사용목적이 비어 있습니다.", 70, 1),
    ("PURPOSE_SHORT", "사용목적 상세도 확인", 1, "FIELD", "purpose", "text_min_length", "4", "주의", "사용목적", "사용목적이 설정된 최소 글자수({threshold}자)보다 짧습니다. 구체적인 업무 목적을 확인하세요.", 80, 1),
    ("EVIDENCE_MISSING", "증빙 누락", 1, "EVIDENCE_MISSING", "evidence_status", "", "", "오류", "증빙", "증빙이 확인되지 않습니다. 증빙자료 또는 증빙번호를 확인하세요.", 90, 1),
    ("EVIDENCE_NO_MISSING", "증빙번호 누락", 1, "EVIDENCE_NO_MISSING", "evidence_no", "", "", "확인", "증빙", "증빙은 있다고 표시되어 있으나 증빙번호가 없습니다.", 100, 1),
    ("VAT_MISMATCH", "공급가액+부가세 불일치", 1, "VAT_MISMATCH", "amount", "", "1", "오류", "금액", "공급가액+부가세와 합계금액이 허용오차({threshold}원)를 초과해 일치하지 않습니다.", 110, 1),
    ("VENDOR_MISSING", "거래처 누락", 1, "FIELD", "vendor", "missing", "", "확인", "거래처", "거래처가 비어 있습니다. 증빙과 대조가 필요한지 확인하세요.", 120, 1),
    ("ACCOUNT_MISSING", "계정과목 누락", 1, "FIELD", "account_name", "missing", "", "확인", "계정과목", "계정과목이 비어 있습니다. 제출 목적상 필요한지 확인하세요.", 130, 1),
    ("DUPLICATE_EXACT", "정확 중복", 1, "DUPLICATE_EXACT", "", "", "", "오류", "중복", "동일 일자·금액·거래처·사용목적의 거래가 중복으로 확인됩니다.", 140, 1),
    ("DUPLICATE_PROBABLE", "중복·분할결제 의심", 1, "DUPLICATE_PROBABLE", "", "", "", "주의", "중복", "동일 일자·금액·거래처 거래가 반복됩니다. 분할·중복 결제 여부를 확인하세요.", 150, 1),
]

DEFAULT_CHECKLIST = [
    ("필수 열(지출일자·금액·사용목적) 누락 확인", "AUTO_RULES", "DATE_MISSING,AMOUNT_MISSING,PURPOSE_MISSING", "", 1, 10),
    ("증빙 누락 건 확인", "AUTO_RULES", "EVIDENCE_MISSING", "", 1, 20),
    ("중복·분할결제 의심 건 확인", "AUTO_RULES", "DUPLICATE_EXACT,DUPLICATE_PROBABLE", "", 1, 30),
    ("금액·부가세 불일치 확인", "AUTO_RULES", "VAT_MISMATCH", "", 1, 40),
    ("고액 지출의 승인·근거 확인", "AUTO_RULES", "AMOUNT_HIGH", "", 1, 50),
    ("모든 오류 등급 이슈 처리", "AUTO_SEVERITY", "", "오류", 1, 60),
    ("최종 승인자가 원본 증빙과 예외처리 근거를 확인", "MANUAL", "", "", 1, 70),
]


def init_db() -> None:
    schema = """
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS imports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT NOT NULL,
        file_hash TEXT,
        imported_at TEXT NOT NULL,
        row_count INTEGER NOT NULL DEFAULT 0,
        mapped_columns TEXT,
        notes TEXT
    );

    CREATE TABLE IF NOT EXISTS transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        import_id INTEGER,
        source_file TEXT NOT NULL,
        source_row INTEGER NOT NULL,
        transaction_id TEXT,
        expense_date TEXT,
        amount REAL,
        supply_amount REAL,
        tax_amount REAL,
        vendor TEXT,
        purpose TEXT,
        account_name TEXT,
        department TEXT,
        employee TEXT,
        evidence_no TEXT,
        evidence_status TEXT,
        payment_method TEXT,
        note TEXT,
        review_status TEXT NOT NULL DEFAULT '미검토',
        reviewer_note TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY(import_id) REFERENCES imports(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS issues (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        transaction_id INTEGER NOT NULL,
        rule_code TEXT NOT NULL,
        severity TEXT NOT NULL,
        category TEXT NOT NULL,
        message TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT '미확인',
        resolution_note TEXT,
        created_at TEXT NOT NULL,
        resolved_at TEXT,
        assignee TEXT,
        due_date TEXT,
        FOREIGN KEY(transaction_id) REFERENCES transactions(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS documents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT NOT NULL,
        stored_name TEXT NOT NULL,
        category TEXT NOT NULL,
        reference_no TEXT,
        summary TEXT,
        text_preview TEXT,
        uploaded_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        action TEXT NOT NULL,
        target_type TEXT,
        target_id TEXT,
        detail TEXT,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS menus (
        menu_key TEXT PRIMARY KEY,
        path TEXT NOT NULL,
        label TEXT NOT NULL,
        enabled INTEGER NOT NULL DEFAULT 1,
        sort_order INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS ui_texts (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL,
        group_name TEXT NOT NULL DEFAULT '공통',
        description TEXT,
        sort_order INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS review_rules (
        code TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        enabled INTEGER NOT NULL DEFAULT 1,
        rule_type TEXT NOT NULL,
        field_name TEXT,
        operator TEXT,
        compare_value TEXT,
        severity TEXT NOT NULL,
        category TEXT NOT NULL,
        message TEXT NOT NULL,
        sort_order INTEGER NOT NULL DEFAULT 0,
        is_system INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS checklist_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        label TEXT NOT NULL,
        item_type TEXT NOT NULL DEFAULT 'MANUAL',
        rule_codes TEXT,
        severity_filter TEXT,
        enabled INTEGER NOT NULL DEFAULT 1,
        sort_order INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS checklist_confirmations (
        item_id INTEGER NOT NULL,
        period TEXT NOT NULL,
        checked INTEGER NOT NULL DEFAULT 0,
        note TEXT,
        updated_at TEXT,
        PRIMARY KEY(item_id, period),
        FOREIGN KEY(item_id) REFERENCES checklist_items(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS mapping_profiles (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        header_signature TEXT NOT NULL UNIQUE,
        headers_json TEXT NOT NULL,
        mapping_json TEXT NOT NULL,
        created_at TEXT NOT NULL,
        last_used_at TEXT
    );

    CREATE UNIQUE INDEX IF NOT EXISTS idx_import_hash ON imports(file_hash) WHERE file_hash IS NOT NULL;
    CREATE INDEX IF NOT EXISTS idx_tx_date ON transactions(expense_date);
    CREATE INDEX IF NOT EXISTS idx_tx_evidence ON transactions(evidence_no);
    CREATE INDEX IF NOT EXISTS idx_issue_status ON issues(status);
    CREATE INDEX IF NOT EXISTS idx_issue_tx ON issues(transaction_id);
    CREATE INDEX IF NOT EXISTS idx_doc_ref ON documents(reference_no);
    """
    with db() as conn:
        conn.executescript(schema)
        # v0.1 DB migration safety
        _ensure_column(conn, "issues", "assignee", "TEXT")
        _ensure_column(conn, "issues", "due_date", "TEXT")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_issue_due ON issues(due_date)")

        # migrate old setting names into menu table on first run
        old = {r["key"]: r["value"] for r in conn.execute("SELECT key,value FROM settings")}
        menu_old_map = {
            "dashboard": "menu_dashboard", "transactions": "menu_transactions", "documents": "menu_documents",
            "checklist": "menu_checklist", "reports": "menu_reports", "ai": "menu_ai", "settings": "menu_settings",
        }
        for key, path, label, enabled, sort_order in DEFAULT_MENUS:
            label = old.get(menu_old_map.get(key, ""), label)
            conn.execute("INSERT OR IGNORE INTO menus(menu_key,path,label,enabled,sort_order) VALUES(?,?,?,?,?)", (key, path, label, enabled, sort_order))

        for key, value in DEFAULT_SETTINGS.items():
            conn.execute("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (key, old.get(key, value)))

        # UI texts grouped by key prefix for admin readability
        for order, (key, value) in enumerate(DEFAULT_UI_TEXTS.items(), start=1):
            group = key.split("_", 1)[0]
            conn.execute("INSERT OR IGNORE INTO ui_texts(key,value,group_name,sort_order) VALUES(?,?,?,?)", (key, value, group, order))

        # Seed rules, preserving v0.1 thresholds if present
        migrated_rules = []
        for row in DEFAULT_RULES:
            row = list(row)
            if row[0] == "AMOUNT_HIGH" and old.get("high_amount_threshold"):
                row[6] = old["high_amount_threshold"]
            elif row[0] == "PURPOSE_SHORT" and old.get("purpose_min_length"):
                row[6] = old["purpose_min_length"]
            elif row[0] == "VAT_MISMATCH" and old.get("vat_tolerance"):
                row[6] = old["vat_tolerance"]
            elif row[0] in {"EVIDENCE_MISSING", "EVIDENCE_NO_MISSING"} and old.get("evidence_required") == "0":
                row[2] = 0
            migrated_rules.append(tuple(row))
        conn.executemany(
            """INSERT OR IGNORE INTO review_rules(code,name,enabled,rule_type,field_name,operator,compare_value,severity,category,message,sort_order,is_system)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
            migrated_rules,
        )

        if conn.execute("SELECT COUNT(*) FROM checklist_items").fetchone()[0] == 0:
            conn.executemany(
                "INSERT INTO checklist_items(label,item_type,rule_codes,severity_filter,enabled,sort_order) VALUES(?,?,?,?,?,?)",
                DEFAULT_CHECKLIST,
            )


def get_settings() -> dict[str, str]:
    with db() as conn:
        return {row["key"]: row["value"] for row in conn.execute("SELECT key, value FROM settings")}


def set_settings(values: dict[str, str]) -> None:
    with db() as conn:
        for key, value in values.items():
            conn.execute(
                "INSERT INTO settings(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (key, str(value)),
            )
        # v0.1 compatibility: legacy threshold settings update the v0.2 Rule table too.
        if "high_amount_threshold" in values:
            conn.execute("UPDATE review_rules SET compare_value=? WHERE code='AMOUNT_HIGH'", (str(values["high_amount_threshold"]),))
        if "purpose_min_length" in values:
            conn.execute("UPDATE review_rules SET compare_value=? WHERE code='PURPOSE_SHORT'", (str(values["purpose_min_length"]),))
        if "vat_tolerance" in values:
            conn.execute("UPDATE review_rules SET compare_value=? WHERE code='VAT_MISMATCH'", (str(values["vat_tolerance"]),))
        if "evidence_required" in values:
            enabled = 1 if str(values["evidence_required"]) == "1" else 0
            conn.execute("UPDATE review_rules SET enabled=? WHERE code IN ('EVIDENCE_MISSING','EVIDENCE_NO_MISSING')", (enabled,))


def get_menus(include_disabled: bool = False) -> list[dict]:
    with db() as conn:
        where = "" if include_disabled else "WHERE enabled=1"
        return [dict(r) for r in conn.execute(f"SELECT * FROM menus {where} ORDER BY sort_order, menu_key")]


def get_ui_texts() -> dict[str, str]:
    with db() as conn:
        return {r["key"]: r["value"] for r in conn.execute("SELECT key,value FROM ui_texts")}


def list_ui_texts() -> list[dict]:
    with db() as conn:
        return [dict(r) for r in conn.execute("SELECT * FROM ui_texts ORDER BY group_name, sort_order, key")]


def set_ui_text(key: str, value: str) -> None:
    with db() as conn:
        conn.execute("UPDATE ui_texts SET value=? WHERE key=?", (value, key))


def get_rules(include_disabled: bool = True) -> list[dict]:
    with db() as conn:
        where = "" if include_disabled else "WHERE enabled=1"
        return [dict(r) for r in conn.execute(f"SELECT * FROM review_rules {where} ORDER BY sort_order, code")]
