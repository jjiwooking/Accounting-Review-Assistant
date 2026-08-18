from __future__ import annotations

from datetime import date
from typing import Any


def _num(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _is_missing(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def field_rule_matches(tx: dict[str, Any], field: str, operator: str, compare: str | None) -> bool:
    value = tx.get(field)
    op = (operator or "").lower()
    if op == "missing":
        return _is_missing(value)
    if op == "not_missing":
        return not _is_missing(value)
    if op == "text_min_length":
        if _is_missing(value):
            return False
        try:
            return len(str(value).strip()) < int(float(compare or 0))
        except ValueError:
            return False
    if op in {"eq", "neq", "gt", "gte", "lt", "lte"}:
        v_num, c_num = _num(value), _num(compare)
        if v_num is not None and c_num is not None:
            return {
                "eq": v_num == c_num, "neq": v_num != c_num,
                "gt": v_num > c_num, "gte": v_num >= c_num,
                "lt": v_num < c_num, "lte": v_num <= c_num,
            }[op]
        v = "" if value is None else str(value).strip().lower()
        c = "" if compare is None else str(compare).strip().lower()
        return (v == c) if op == "eq" else (v != c if op == "neq" else False)
    if op == "contains":
        return str(compare or "").lower() in str(value or "").lower()
    if op == "not_contains":
        return str(compare or "").lower() not in str(value or "").lower()
    return False


def rule_matches(tx: dict[str, Any], rule: dict[str, Any], evidence_refs: set[str]) -> bool:
    kind = rule.get("rule_type")
    if kind == "FIELD":
        return field_rule_matches(tx, rule.get("field_name") or "", rule.get("operator") or "", rule.get("compare_value"))
    if kind == "DATE_FUTURE":
        value = tx.get("expense_date")
        if not value:
            return False
        try:
            return date.fromisoformat(str(value)) > date.today()
        except ValueError:
            return False
    if kind == "DATE_WEEKEND":
        value = tx.get("expense_date")
        if not value:
            return False
        try:
            return date.fromisoformat(str(value)).weekday() >= 5
        except ValueError:
            return False
    if kind == "EVIDENCE_MISSING":
        status = (tx.get("evidence_status") or "").strip()
        no = (tx.get("evidence_no") or "").strip()
        matched = bool(no and no.lower() in evidence_refs)
        return status != "있음" and not matched
    if kind == "EVIDENCE_NO_MISSING":
        status = (tx.get("evidence_status") or "").strip()
        no = (tx.get("evidence_no") or "").strip()
        matched = bool(no and no.lower() in evidence_refs)
        return status == "있음" and not no and not matched
    if kind == "VAT_MISMATCH":
        amount, supply, tax = tx.get("amount"), tx.get("supply_amount"), tx.get("tax_amount")
        if amount is None or supply is None or tax is None:
            return False
        tol = _num(rule.get("compare_value")) or 0
        return abs((float(supply) + float(tax)) - float(amount)) > tol
    return False


def render_message(tx: dict[str, Any], rule: dict[str, Any]) -> str:
    message = rule.get("message") or rule.get("name") or rule.get("code") or "확인 필요"
    compare = rule.get("compare_value") or ""
    replacements = {
        "{threshold}": compare,
        "{value}": "" if tx.get(rule.get("field_name") or "") is None else str(tx.get(rule.get("field_name") or "")),
        "{vendor}": tx.get("vendor") or "",
        "{date}": tx.get("expense_date") or "",
    }
    for k, v in replacements.items():
        message = message.replace(k, v)
    return message
