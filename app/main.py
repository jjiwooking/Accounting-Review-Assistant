from __future__ import annotations

import json
import uuid
import hashlib
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from .db import (
    BASE_DIR, DATA_DIR, db, get_settings, init_db, set_settings,
    get_menus, get_ui_texts, list_ui_texts, get_rules,
)
from .documents import delete_document, get_document, list_documents, save_document, UPLOAD_DIR
from .exports import create_audit_package, create_demo_data, create_full_backup, create_review_workbook, create_sample_template
from .services import (
    DISPLAY_NAMES, ai_assist_prompt, checklist_data, confirmation_message,
    dashboard_summary, data_quality_score, detect_mapping, get_transaction_detail,
    get_transactions, import_transactions, monthly_report, months_available,
    read_table_bytes, resolve_issue, set_checklist_confirmation, update_transaction,
    validate_all,
)

app = FastAPI(title="회계자료 검토 및 보고 보조 프로그램")
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"
STAGING_DIR = DATA_DIR / "staging"
STAGING_DIR.mkdir(parents=True, exist_ok=True)
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


def ctx(request: Request, **kwargs):
    settings = get_settings()
    active_menus = get_menus()
    all_menus = get_menus(include_disabled=True)
    base = {
        "request": request,
        "settings": settings,
        "program_name": settings.get("program_name", "회계자료 검토 및 보고 보조 프로그램"),
        "menus": active_menus,
        "menu_map": {m["menu_key"]: m for m in all_menus},
        "texts": get_ui_texts(),
    }
    base.update(kwargs)
    return base


init_db()


def _validate_rule_value(rule_type: str, field_name: str, operator: str, compare_value: str) -> None:
    numeric_fields = {"amount", "supply_amount", "tax_amount"}
    if rule_type == "VAT_MISMATCH":
        try:
            if float(compare_value) < 0:
                raise ValueError
        except ValueError:
            raise HTTPException(400, "부가세 허용오차는 0 이상의 숫자여야 합니다.")
    if rule_type == "FIELD" and operator in {"gt","gte","lt","lte"} and field_name in numeric_fields:
        try:
            float(compare_value)
        except ValueError:
            raise HTTPException(400, "금액 비교 검토 기준의 기준값은 숫자여야 합니다.")
    if rule_type == "FIELD" and operator == "text_min_length":
        try:
            value = int(float(compare_value))
            if value < 0:
                raise ValueError
        except ValueError:
            raise HTTPException(400, "최소 글자수는 0 이상의 정수여야 합니다.")


def _validate_checklist_config(item_type: str, rule_codes: str, severity_filter: str) -> None:
    if item_type not in {"MANUAL", "AUTO_RULES", "AUTO_SEVERITY"}:
        raise HTTPException(400, "체크리스트 유형이 올바르지 않습니다.")
    if item_type == "AUTO_SEVERITY" and severity_filter not in {"오류", "주의", "확인"}:
        raise HTTPException(400, "AUTO_SEVERITY는 오류/주의/확인 중 하나를 지정해야 합니다.")
    if item_type == "AUTO_RULES":
        codes = [x.strip() for x in rule_codes.split(",") if x.strip()]
        if not codes:
            raise HTTPException(400, "자동 기준 연동은 하나 이상의 검토 코드를 지정해야 합니다.")
        with db() as conn:
            known = {r[0] for r in conn.execute("SELECT code FROM review_rules")}
        missing = [c for c in codes if c not in known]
        if missing:
            raise HTTPException(400, f"존재하지 않는 검토 코드가 있습니다: {', '.join(missing)}")


@app.get("/health")
def health():
    return {"ok": True, "version": "0.3"}


@app.get("/")
def dashboard(request: Request):
    summary = dashboard_summary()
    quality = data_quality_score()
    recent = get_transactions()[:8]
    with db() as conn:
        issue_summary = [dict(r) for r in conn.execute(
            """SELECT category, COUNT(*) count FROM issues
               WHERE status NOT IN ('확인완료','예외인정')
               GROUP BY category ORDER BY count DESC LIMIT 6"""
        )]
        imports = [dict(r) for r in conn.execute("SELECT * FROM imports ORDER BY id DESC LIMIT 5")]
    return templates.TemplateResponse(request=request, name="dashboard.html", context=ctx(
        request, summary=summary, quality=quality, recent=recent, issue_summary=issue_summary, imports=imports
    ))


@app.get("/upload")
def upload_page(request: Request):
    return templates.TemplateResponse(request=request, name="upload.html", context=ctx(request))


@app.post("/upload/preview")
async def upload_preview(request: Request, file: UploadFile = File(...)):
    content = await file.read()
    if not content:
        raise HTTPException(400, "빈 파일입니다.")
    if len(content) > 30 * 1024 * 1024:
        raise HTTPException(400, "30MB 이하 파일만 업로드할 수 있습니다.")
    try:
        headers, rows = read_table_bytes(file.filename or "upload.xlsx", content)
    except ValueError as e:
        return templates.TemplateResponse(request=request, name="upload.html", context=ctx(request, error=str(e)), status_code=400)
    auto_mapping = detect_mapping(headers)
    header_signature = hashlib.sha256(json.dumps(headers, ensure_ascii=False, separators=(",", ":")).encode("utf-8")).hexdigest()
    profile_name = ""
    with db() as conn:
        profile = conn.execute("SELECT * FROM mapping_profiles WHERE header_signature=?", (header_signature,)).fetchone()
        if profile:
            try:
                saved = json.loads(profile["mapping_json"])
                mapping = {k: v for k, v in saved.items() if v in headers}
                profile_name = profile["name"]
                conn.execute("UPDATE mapping_profiles SET last_used_at=datetime('now','localtime') WHERE id=?", (profile["id"],))
            except Exception:
                mapping = auto_mapping
        else:
            mapping = auto_mapping
    token = uuid.uuid4().hex
    bin_path = STAGING_DIR / f"{token}.bin"
    meta_path = STAGING_DIR / f"{token}.json"
    bin_path.write_bytes(content)
    meta_path.write_text(json.dumps({"filename": file.filename, "headers": headers, "header_signature": header_signature}, ensure_ascii=False), encoding="utf-8")
    fields = ["transaction_id", "expense_date", "amount", "supply_amount", "tax_amount", "vendor", "purpose", "account_name", "department", "employee", "evidence_no", "evidence_status", "payment_method", "note"]
    return templates.TemplateResponse(request=request, name="mapping.html", context=ctx(
        request, token=token, filename=file.filename, headers=headers, preview=rows[:5], mapping=mapping,
        fields=fields, display_names=DISPLAY_NAMES, profile_name=profile_name
    ))


@app.post("/upload/confirm")
def upload_confirm(
    request: Request, token: str = Form(...),
    transaction_id: str = Form(""), expense_date: str = Form(""), amount: str = Form(""),
    supply_amount: str = Form(""), tax_amount: str = Form(""), vendor: str = Form(""), purpose: str = Form(""),
    account_name: str = Form(""), department: str = Form(""), employee: str = Form(""), evidence_no: str = Form(""),
    evidence_status: str = Form(""), payment_method: str = Form(""), note: str = Form(""),
):
    bin_path = STAGING_DIR / f"{token}.bin"
    meta_path = STAGING_DIR / f"{token}.json"
    if not bin_path.exists() or not meta_path.exists():
        return templates.TemplateResponse(request=request, name="upload.html", context=ctx(request, error="업로드 임시파일이 만료되었거나 없습니다. 다시 업로드해 주세요."), status_code=400)
    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    mapping_override = {
        "transaction_id": transaction_id, "expense_date": expense_date, "amount": amount,
        "supply_amount": supply_amount, "tax_amount": tax_amount, "vendor": vendor, "purpose": purpose,
        "account_name": account_name, "department": department, "employee": employee,
        "evidence_no": evidence_no, "evidence_status": evidence_status, "payment_method": payment_method, "note": note,
    }
    try:
        result = import_transactions(meta["filename"], bin_path.read_bytes(), mapping_override)
        signature = meta.get("header_signature") or hashlib.sha256(json.dumps(meta.get("headers", []), ensure_ascii=False, separators=(",", ":")).encode("utf-8")).hexdigest()
        profile_name = f"자동 저장 · {Path(meta['filename']).stem[:50]}"
        with db() as conn:
            conn.execute(
                """INSERT INTO mapping_profiles(name,header_signature,headers_json,mapping_json,created_at,last_used_at)
                   VALUES(?,?,?,?,datetime('now','localtime'),datetime('now','localtime'))
                   ON CONFLICT(header_signature) DO UPDATE SET mapping_json=excluded.mapping_json,last_used_at=excluded.last_used_at""",
                (profile_name, signature, json.dumps(meta.get("headers", []), ensure_ascii=False), json.dumps(mapping_override, ensure_ascii=False)),
            )
    except ValueError as e:
        return templates.TemplateResponse(request=request, name="upload.html", context=ctx(request, error=str(e)), status_code=400)
    finally:
        bin_path.unlink(missing_ok=True); meta_path.unlink(missing_ok=True)
    return RedirectResponse(url=f"/transactions?imported={result['rows']}", status_code=303)


@app.get("/transactions")
def transactions_page(request: Request, q: str = "", status: str = "", month: str = "", imported: int | None = None):
    rows = get_transactions(q=q, status=status, month=month)
    return templates.TemplateResponse(request=request, name="transactions.html", context=ctx(
        request, rows=rows, q=q, status=status, month=month, months=months_available(), imported=imported
    ))


@app.get("/issues")
def issues_page(
    request: Request, status: str = "open", severity: str = "", assignee: str = "", due: str = "", rule: str = ""
):
    clauses = []; params: list[object] = []
    if status == "open":
        clauses.append("i.status NOT IN ('확인완료','예외인정')")
    elif status:
        clauses.append("i.status=?"); params.append(status)
    if severity:
        clauses.append("i.severity=?"); params.append(severity)
    if assignee == "__UNASSIGNED__":
        clauses.append("(i.assignee IS NULL OR TRIM(i.assignee)='')")
    elif assignee:
        clauses.append("COALESCE(i.assignee,'') LIKE ?"); params.append(f"%{assignee}%")
    if due == "overdue":
        clauses.append("i.status NOT IN ('확인완료','예외인정') AND i.due_date IS NOT NULL AND i.due_date < date('now','localtime')")
    if rule:
        if rule.endswith("_"):
            clauses.append("i.rule_code LIKE ?"); params.append(rule + "%")
        else:
            clauses.append("i.rule_code=?"); params.append(rule)
    where = "WHERE " + " AND ".join(clauses) if clauses else ""
    with db() as conn:
        rows = [dict(r) for r in conn.execute(f"""
            SELECT i.*, t.source_file,t.source_row,t.expense_date,t.amount,t.vendor,t.purpose,t.employee,t.department
            FROM issues i JOIN transactions t ON t.id=i.transaction_id
            {where}
            ORDER BY CASE i.severity WHEN '오류' THEN 1 WHEN '주의' THEN 2 ELSE 3 END,
                     CASE WHEN i.due_date IS NULL THEN 1 ELSE 0 END, i.due_date, i.id DESC
            LIMIT 500
        """, params)]
        assignees = [r[0] for r in conn.execute("SELECT DISTINCT assignee FROM issues WHERE assignee IS NOT NULL AND TRIM(assignee)<>'' ORDER BY assignee")]
        rule_codes = [r[0] for r in conn.execute("SELECT DISTINCT rule_code FROM issues ORDER BY rule_code")]
    return templates.TemplateResponse(request=request, name="issues.html", context=ctx(
        request, rows=rows, status=status, severity=severity, assignee=assignee, due=due, rule=rule,
        assignees=assignees, rule_codes=rule_codes
    ))


@app.post("/issues/{issue_id}/update")
def issue_queue_update(
    issue_id: int, status: str = Form(...), note: str = Form(""), assignee: str = Form(""), due_date: str = Form("")
):
    resolve_issue(issue_id, status, note, assignee, due_date)
    return RedirectResponse(url="/issues", status_code=303)


@app.get("/transactions/{tx_id}")
def transaction_detail(request: Request, tx_id: int):
    tx = get_transaction_detail(tx_id)
    if not tx:
        raise HTTPException(404, "거래를 찾을 수 없습니다.")
    messages = {issue["id"]: confirmation_message(tx, issue) for issue in tx["issues"]}
    return templates.TemplateResponse(request=request, name="transaction_detail.html", context=ctx(request, tx=tx, messages=messages))


@app.post("/issues/{issue_id}/resolve")
def issue_resolve(
    issue_id: int, tx_id: int = Form(...), status: str = Form(...), note: str = Form(""),
    assignee: str = Form(""), due_date: str = Form(""),
):
    resolve_issue(issue_id, status, note, assignee, due_date)
    return RedirectResponse(url=f"/transactions/{tx_id}", status_code=303)


@app.post("/transactions/{tx_id}/memo")
def transaction_memo(tx_id: int, reviewer_note: str = Form("")):
    update_transaction(tx_id, reviewer_note)
    return RedirectResponse(url=f"/transactions/{tx_id}", status_code=303)


@app.get("/documents")
def documents_page(request: Request):
    return templates.TemplateResponse(request=request, name="documents.html", context=ctx(request, documents=list_documents()))


@app.post("/documents/upload")
async def document_upload(request: Request, file: UploadFile = File(...), reference_no: str = Form(""), category: str = Form("")):
    content = await file.read()
    try:
        save_document(file.filename or "document", content, reference_no=reference_no, category_override=category)
    except ValueError as e:
        return templates.TemplateResponse(request=request, name="documents.html", context=ctx(request, documents=list_documents(), error=str(e)), status_code=400)
    return RedirectResponse(url="/documents", status_code=303)


@app.get("/documents/{doc_id}/download")
def document_download(doc_id: int):
    doc = get_document(doc_id)
    if not doc:
        raise HTTPException(404, "문서를 찾을 수 없습니다.")
    path = UPLOAD_DIR / doc["stored_name"]
    if not path.exists():
        raise HTTPException(404, "저장 파일이 없습니다.")
    return FileResponse(path, filename=doc["filename"])


@app.post("/documents/{doc_id}/delete")
def document_delete(doc_id: int):
    delete_document(doc_id)
    return RedirectResponse(url="/documents", status_code=303)


@app.get("/checklist")
def checklist_page(request: Request, month: str = ""):
    data = checklist_data(month or None)
    with db() as conn:
        params = []
        period_clause = ""
        if month:
            period_clause = " AND SUBSTR(t.expense_date,1,7)=?"
            params.append(month)
        unresolved = [dict(r) for r in conn.execute(f"""
            SELECT i.*, t.expense_date, t.amount, t.vendor, t.purpose, t.source_file, t.source_row
            FROM issues i JOIN transactions t ON t.id=i.transaction_id
            WHERE i.status NOT IN ('확인완료','예외인정') {period_clause}
            ORDER BY CASE i.severity WHEN '오류' THEN 1 WHEN '주의' THEN 2 ELSE 3 END, i.id DESC LIMIT 150
        """, params)]
    return templates.TemplateResponse(request=request, name="checklist.html", context=ctx(
        request, data=data, unresolved=unresolved, months=months_available(), selected_month=month
    ))


@app.post("/checklist/{item_id}/confirm")
def checklist_confirm(item_id: int, period: str = Form("전체"), checked: str = Form("0"), note: str = Form("")):
    set_checklist_confirmation(item_id, period, checked == "1", note)
    q = "" if period == "전체" else f"?month={period}"
    return RedirectResponse(url=f"/checklist{q}", status_code=303)


@app.get("/reports")
def reports_page(request: Request, month: str = ""):
    report = monthly_report(month or None)
    return templates.TemplateResponse(request=request, name="reports.html", context=ctx(request, report=report))


@app.get("/ai-assist")
def ai_assist_page(request: Request, month: str = ""):
    pack = ai_assist_prompt(month or None)
    return templates.TemplateResponse(request=request, name="ai_assist.html", context=ctx(request, pack=pack, months=months_available()))


@app.get("/settings")
def settings_page(request: Request, saved: int | None = None):
    with db() as conn:
        imports = [dict(r) for r in conn.execute("SELECT * FROM imports ORDER BY id DESC")]
        logs = [dict(r) for r in conn.execute("SELECT * FROM audit_log ORDER BY id DESC LIMIT 80")]
        checklist_items = [dict(r) for r in conn.execute("SELECT * FROM checklist_items ORDER BY sort_order,id")]
        mapping_profiles = [dict(r) for r in conn.execute("SELECT * FROM mapping_profiles ORDER BY COALESCE(last_used_at,created_at) DESC,id DESC")]
    return templates.TemplateResponse(request=request, name="settings.html", context=ctx(
        request, imports=imports, logs=logs, all_menus=get_menus(include_disabled=True),
        ui_items=list_ui_texts(), rules=get_rules(include_disabled=True), checklist_items=checklist_items,
        mapping_profiles=mapping_profiles, saved=saved,
    ))


@app.post("/settings/basic")
def settings_basic(program_name: str = Form(...), report_company_name: str = Form(...)):
    set_settings({"program_name": program_name.strip(), "report_company_name": report_company_name.strip()})
    return RedirectResponse(url="/settings?saved=1#basic", status_code=303)


@app.post("/settings/menu/{menu_key}")
def settings_menu(menu_key: str, label: str = Form(...), enabled: str = Form("0"), sort_order: int = Form(0)):
    with db() as conn:
        row = conn.execute("SELECT 1 FROM menus WHERE menu_key=?", (menu_key,)).fetchone()
        if not row:
            raise HTTPException(404, "메뉴를 찾을 수 없습니다.")
        conn.execute("UPDATE menus SET label=?,enabled=?,sort_order=? WHERE menu_key=?", (label.strip(), 1 if enabled == "1" else 0, sort_order, menu_key))
        conn.execute("INSERT INTO audit_log(action,target_type,target_id,detail,created_at) VALUES(?,?,?,?,datetime('now','localtime'))", ("메뉴 설정 변경", "menu", menu_key, f"{label} / 표시={enabled} / 순서={sort_order}"))
    return RedirectResponse(url="/settings?saved=1#menus", status_code=303)


@app.post("/settings/text/{text_key}")
def settings_text(text_key: str, value: str = Form(...)):
    with db() as conn:
        if not conn.execute("SELECT 1 FROM ui_texts WHERE key=?", (text_key,)).fetchone():
            raise HTTPException(404, "화면 문구를 찾을 수 없습니다.")
        conn.execute("UPDATE ui_texts SET value=? WHERE key=?", (value.strip(), text_key))
        conn.execute("INSERT INTO audit_log(action,target_type,target_id,detail,created_at) VALUES(?,?,?,?,datetime('now','localtime'))", ("화면 문구 변경", "ui_text", text_key, value[:200]))
    return RedirectResponse(url="/settings?saved=1#texts", status_code=303)


@app.post("/settings/rule/{code}")
def settings_rule_update(
    code: str, name: str = Form(...), enabled: str = Form("0"), severity: str = Form(...), category: str = Form(...),
    compare_value: str = Form(""), message: str = Form(...), sort_order: int = Form(0),
):
    with db() as conn:
        existing = conn.execute("SELECT rule_type,field_name,operator FROM review_rules WHERE code=?", (code,)).fetchone()
        if not existing:
            raise HTTPException(404, "검토 기준을 찾을 수 없습니다.")
        _validate_rule_value(existing["rule_type"], existing["field_name"] or "", existing["operator"] or "", compare_value.strip())
        conn.execute("""UPDATE review_rules SET name=?,enabled=?,severity=?,category=?,compare_value=?,message=?,sort_order=? WHERE code=?""",
                     (name.strip(), 1 if enabled == "1" else 0, severity, category.strip(), compare_value.strip(), message.strip(), sort_order, code))
        conn.execute("INSERT INTO audit_log(action,target_type,target_id,detail,created_at) VALUES(?,?,?,?,datetime('now','localtime'))", ("검토 기준 변경", "rule", code, f"{name} / {severity} / 활성={enabled}"))
    validate_all()
    return RedirectResponse(url="/settings?saved=1#rules", status_code=303)


@app.post("/settings/rule-add")
def settings_rule_add(
    name: str = Form(...), field_name: str = Form(...), operator: str = Form(...), compare_value: str = Form(""),
    severity: str = Form("확인"), category: str = Form("사용자 기준"), message: str = Form(...),
):
    allowed_fields = {"expense_date","amount","supply_amount","tax_amount","vendor","purpose","account_name","department","employee","evidence_no","evidence_status","payment_method","note"}
    allowed_ops = {"missing","not_missing","eq","neq","gt","gte","lt","lte","contains","not_contains","text_min_length"}
    if field_name not in allowed_fields or operator not in allowed_ops:
        raise HTTPException(400, "지원하지 않는 필드 또는 조건입니다.")
    _validate_rule_value("FIELD", field_name, operator, compare_value.strip())
    code = f"CUSTOM_{uuid.uuid4().hex[:10].upper()}"
    with db() as conn:
        order = conn.execute("SELECT COALESCE(MAX(sort_order),0)+10 FROM review_rules").fetchone()[0]
        conn.execute("""INSERT INTO review_rules(code,name,enabled,rule_type,field_name,operator,compare_value,severity,category,message,sort_order,is_system)
                        VALUES(?,?,1,'FIELD',?,?,?,?,?,?,?,0)""",
                     (code, name.strip(), field_name, operator, compare_value.strip(), severity, category.strip(), message.strip(), order))
        conn.execute("INSERT INTO audit_log(action,target_type,target_id,detail,created_at) VALUES(?,?,?,?,datetime('now','localtime'))", ("사용자 기준 추가", "rule", code, name.strip()))
    validate_all()
    return RedirectResponse(url="/settings?saved=1#rules", status_code=303)


@app.post("/settings/rule/{code}/delete")
def settings_rule_delete(code: str):
    with db() as conn:
        row = conn.execute("SELECT is_system,name FROM review_rules WHERE code=?", (code,)).fetchone()
        if not row:
            raise HTTPException(404, "검토 기준을 찾을 수 없습니다.")
        if row["is_system"]:
            raise HTTPException(400, "기본 검토 기준은 삭제 대신 사용 안 함으로 설정하세요.")
        refs = [dict(r) for r in conn.execute("SELECT id,label,rule_codes FROM checklist_items WHERE item_type='AUTO_RULES'")]
        used_by = [r["label"] for r in refs if code in [x.strip() for x in (r.get("rule_codes") or "").split(",") if x.strip()]]
        if used_by:
            raise HTTPException(400, f"체크리스트에서 사용 중인 검토 기준입니다. 먼저 체크리스트 연결을 해제하세요: {', '.join(used_by)}")
        conn.execute("DELETE FROM review_rules WHERE code=?", (code,))
        conn.execute("INSERT INTO audit_log(action,target_type,target_id,detail,created_at) VALUES(?,?,?,?,datetime('now','localtime'))", ("사용자 기준 삭제", "rule", code, row["name"]))
    validate_all()
    return RedirectResponse(url="/settings?saved=1#rules", status_code=303)


@app.post("/settings/checklist-add")
def checklist_item_add(label: str = Form(...), item_type: str = Form("MANUAL"), rule_codes: str = Form(""), severity_filter: str = Form("")):
    _validate_checklist_config(item_type, rule_codes.strip(), severity_filter.strip())
    with db() as conn:
        order = conn.execute("SELECT COALESCE(MAX(sort_order),0)+10 FROM checklist_items").fetchone()[0]
        conn.execute("INSERT INTO checklist_items(label,item_type,rule_codes,severity_filter,enabled,sort_order) VALUES(?,?,?,?,1,?)", (label.strip(), item_type, rule_codes.strip(), severity_filter.strip(), order))
    return RedirectResponse(url="/settings?saved=1#checklist", status_code=303)


@app.post("/settings/checklist/{item_id}")
def checklist_item_update(item_id: int, label: str = Form(...), item_type: str = Form(...), rule_codes: str = Form(""), severity_filter: str = Form(""), enabled: str = Form("0"), sort_order: int = Form(0)):
    _validate_checklist_config(item_type, rule_codes.strip(), severity_filter.strip())
    with db() as conn:
        conn.execute("UPDATE checklist_items SET label=?,item_type=?,rule_codes=?,severity_filter=?,enabled=?,sort_order=? WHERE id=?", (label.strip(), item_type, rule_codes.strip(), severity_filter.strip(), 1 if enabled == "1" else 0, sort_order, item_id))
    return RedirectResponse(url="/settings?saved=1#checklist", status_code=303)


@app.post("/settings/checklist/{item_id}/delete")
def checklist_item_delete(item_id: int):
    with db() as conn:
        conn.execute("DELETE FROM checklist_items WHERE id=?", (item_id,))
    return RedirectResponse(url="/settings?saved=1#checklist", status_code=303)


@app.post("/settings/mapping-profile/{profile_id}/delete")
def mapping_profile_delete(profile_id: int):
    with db() as conn:
        row = conn.execute("SELECT name FROM mapping_profiles WHERE id=?", (profile_id,)).fetchone()
        if row:
            conn.execute("DELETE FROM mapping_profiles WHERE id=?", (profile_id,))
            conn.execute("INSERT INTO audit_log(action,target_type,target_id,detail,created_at) VALUES(?,?,?,?,datetime('now','localtime'))", ("매핑 프로필 삭제", "mapping_profile", str(profile_id), row["name"]))
    return RedirectResponse(url="/settings#mapping-profiles", status_code=303)


@app.post("/validate")
def validate_route():
    validate_all()
    return RedirectResponse(url="/checklist", status_code=303)


@app.post("/imports/{import_id}/delete")
def import_delete(import_id: int):
    with db() as conn:
        row = conn.execute("SELECT filename FROM imports WHERE id=?", (import_id,)).fetchone()
        if row:
            conn.execute("DELETE FROM imports WHERE id=?", (import_id,))
            conn.execute("INSERT INTO audit_log(action,target_type,target_id,detail,created_at) VALUES(?,?,?,?,datetime('now','localtime'))", ("업로드 자료 삭제", "import", str(import_id), row["filename"]))
    validate_all()
    return RedirectResponse(url="/settings#data", status_code=303)


@app.get("/export/sample")
def export_sample():
    path = create_sample_template(); return FileResponse(path, filename=path.name)

@app.get("/export/demo-data")
def export_demo_data():
    path = create_demo_data(); return FileResponse(path, filename=path.name)

@app.get("/export/review")
def export_review(month: str = ""):
    path = create_review_workbook(month or None); return FileResponse(path, filename=path.name)

@app.get("/export/audit-package")
def export_audit_package(month: str = ""):
    path = create_audit_package(month or None); return FileResponse(path, filename=path.name)

@app.get("/export/backup")
def export_backup():
    path = create_full_backup(); return FileResponse(path, filename=path.name)
