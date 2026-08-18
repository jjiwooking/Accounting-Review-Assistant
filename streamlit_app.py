from __future__ import annotations

import hashlib
import json
import os
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from app.db import db, get_menus, get_rules, get_settings, get_ui_texts, init_db, list_ui_texts, set_settings, set_ui_text
from app.documents import list_documents, save_document
from app.exports import create_audit_package, create_demo_data, create_review_workbook, create_sample_template
from app.services import (
    DISPLAY_NAMES,
    ai_assist_prompt,
    checklist_data,
    confirmation_message,
    dashboard_summary,
    data_quality_score,
    detect_mapping,
    get_transaction_detail,
    get_transactions,
    import_transactions,
    monthly_report,
    months_available,
    read_table_bytes,
    resolve_issue,
    set_checklist_confirmation,
    update_transaction,
    validate_all,
)

st.set_page_config(
    page_title="회계자료 검토 및 보고 보조",
    page_icon="◼",
    layout="wide",
    initial_sidebar_state="expanded",
)

init_db()
BASE_DIR = Path(__file__).resolve().parent

# -----------------------------------------------------------------------------
# Visual system: original accounting UI inspired only by the referenced mood
# (bold geometry, generous whitespace, monochrome base + primary color accents).
# -----------------------------------------------------------------------------
CSS = r"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@500;600;700;800&family=Noto+Sans+KR:wght@400;500;600;700;800&display=swap');
:root{
  --paper:#f4f3ef; --ink:#111111; --muted:#6d6d68; --line:#d9d7cf;
  --blue:#2447ff; --blue2:#162ccb; --red:#ff3b30; --yellow:#ffd900; --white:#ffffff;
  --shadow:0 12px 35px rgba(17,17,17,.06);
}
html, body, [class*="css"]{font-family:'Noto Sans KR','Inter',sans-serif;}
.stApp{background:var(--paper); color:var(--ink);}
[data-testid="stSidebar"]{background:#111111;border-right:0;}
[data-testid="stSidebar"] *{color:#f7f7f4;}
[data-testid="stSidebar"] .stRadio label{padding:.36rem .15rem;font-weight:650;}
[data-testid="stSidebar"] .stButton button{border:1px solid #3b3b3b;background:#1b1b1b;color:#fff;border-radius:0;}
[data-testid="stSidebar"] .stButton button:hover{border-color:var(--yellow);color:var(--yellow);}
.block-container{padding-top:2.1rem;padding-bottom:4rem;max-width:1450px;}
#MainMenu{visibility:hidden;} footer{visibility:hidden;}

.ds-kicker{font:800 12px/1 'Inter',sans-serif;letter-spacing:.18em;text-transform:uppercase;color:var(--blue);margin-bottom:14px;}
.ds-title{font-size:clamp(32px,4vw,58px);line-height:1.02;letter-spacing:-.055em;font-weight:800;margin:0;max-width:900px;}
.ds-sub{font-size:15px;line-height:1.75;color:var(--muted);max-width:820px;margin-top:18px;}
.ds-hero{position:relative;background:var(--white);border:1px solid var(--line);padding:34px 36px 30px;overflow:hidden;box-shadow:var(--shadow);margin-bottom:24px;}
.ds-hero:before{content:"";position:absolute;width:160px;height:42px;background:var(--red);right:58px;top:22px;transform:rotate(-7deg);}
.ds-hero:after{content:"";position:absolute;width:74px;height:74px;border-radius:50%;background:var(--yellow);right:24px;bottom:-22px;}
.ds-slash{position:absolute;width:35px;height:120px;background:var(--blue);right:188px;bottom:-28px;transform:skew(-17deg) rotate(4deg);}

.ds-grid{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:12px;margin:16px 0 26px;}
.ds-card{background:var(--white);border:1px solid var(--line);padding:18px 18px 16px;min-height:116px;position:relative;box-shadow:0 4px 14px rgba(17,17,17,.025);}
.ds-card .label{font-size:12px;color:var(--muted);font-weight:700;margin-bottom:18px;}
.ds-card .value{font-family:'Inter','Noto Sans KR',sans-serif;font-size:30px;line-height:1;font-weight:800;letter-spacing:-.04em;}
.ds-card .foot{font-size:11px;color:#8b8a84;margin-top:10px;}
.ds-card.blue{background:var(--blue);border-color:var(--blue);color:white}.ds-card.blue .label,.ds-card.blue .foot{color:#dce2ff}
.ds-card.dark{background:#111;border-color:#111;color:#fff}.ds-card.dark .label,.ds-card.dark .foot{color:#aaa}
.ds-card.alert:before{content:"";position:absolute;width:12px;height:12px;background:var(--red);right:14px;top:14px;border-radius:50%;}

.ds-section{display:flex;justify-content:space-between;align-items:flex-end;gap:20px;margin:34px 0 13px;border-top:1px solid var(--ink);padding-top:13px;}
.ds-section h2{font-size:24px;letter-spacing:-.035em;margin:0;font-weight:800}.ds-section p{margin:0;color:var(--muted);font-size:12px;}
.ds-note{border-left:5px solid var(--blue);background:#fff;padding:15px 17px;margin:12px 0;color:#343434;font-size:13px;line-height:1.65;}
.ds-warning{border-left-color:var(--red);background:#fff7f6;}
.ds-success{border-left-color:#22a06b;background:#f4fff8;}
.ds-chip{display:inline-block;border:1px solid var(--ink);padding:4px 8px;margin-right:5px;font:700 10px/1.2 'Inter','Noto Sans KR';background:#fff;}
.ds-chip.red{background:var(--red);border-color:var(--red);color:#fff}.ds-chip.yellow{background:var(--yellow);border-color:var(--yellow)}.ds-chip.blue{background:var(--blue);border-color:var(--blue);color:#fff}

div[data-testid="stMetric"]{background:#fff;border:1px solid var(--line);padding:14px 16px;}
div[data-testid="stMetric"] label{font-size:12px!important;color:var(--muted)!important;font-weight:700!important;}
div[data-testid="stMetric"] [data-testid="stMetricValue"]{font-size:26px!important;font-weight:800!important;letter-spacing:-.035em;}
.stButton>button,.stDownloadButton>button{border-radius:0!important;border:1px solid #111!important;background:#111!important;color:#fff!important;font-weight:750!important;min-height:42px;}
.stButton>button:hover,.stDownloadButton>button:hover{background:var(--blue)!important;border-color:var(--blue)!important;color:#fff!important;}
.stTextInput input,.stNumberInput input,.stDateInput input,.stSelectbox div[data-baseweb="select"]>div,.stTextArea textarea{border-radius:0!important;background:#fff!important;}
[data-testid="stFileUploaderDropzone"]{border-radius:0!important;background:#fff!important;border:1px dashed #777!important;}
[data-testid="stDataFrame"]{background:#fff;border:1px solid var(--line);}
.stTabs [data-baseweb="tab-list"]{gap:4px}.stTabs [data-baseweb="tab"]{border-radius:0;background:#fff;border:1px solid var(--line);padding:10px 15px}.stTabs [aria-selected="true"]{background:#111!important;color:#fff!important;}
hr{border-color:var(--line)!important;}
@media(max-width:1000px){.ds-grid{grid-template-columns:repeat(2,minmax(0,1fr));}.ds-hero:before,.ds-hero:after,.ds-slash{opacity:.25}}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def won(value: float | int | None) -> str:
    return f"₩{float(value or 0):,.0f}"


def section(title: str, note: str = "") -> None:
    st.markdown(f'<div class="ds-section"><h2>{title}</h2><p>{note}</p></div>', unsafe_allow_html=True)


def hero(kicker: str, title: str, desc: str) -> None:
    st.markdown(
        f'''<div class="ds-hero"><div class="ds-slash"></div><div class="ds-kicker">{kicker}</div>
        <h1 class="ds-title">{title}</h1><div class="ds-sub">{desc}</div></div>''',
        unsafe_allow_html=True,
    )


def metric_cards(summary: dict) -> None:
    cards = [
        ("전체 거래", f"{summary['total']:,}", f"총 {won(summary['total_amount'])}", ""),
        ("검토 필요", f"{summary['review_required']:,}", f"정상 {summary['normal_count']:,}건", "blue"),
        ("증빙 누락", f"{summary['missing_evidence']:,}", "증빙번호·문서 연결 기준", "alert"),
        ("중복 의심", f"{summary['duplicates']:,}", "활성 중복 Rule 기준", ""),
        ("기한 경과", f"{summary['overdue']:,}", f"담당자 미지정 {summary['unassigned']:,}", "dark"),
    ]
    html = ['<div class="ds-grid">']
    for label, value, foot, cls in cards:
        html.append(f'<div class="ds-card {cls}"><div class="label">{label}</div><div class="value">{value}</div><div class="foot">{foot}</div></div>')
    html.append('</div>')
    st.markdown(''.join(html), unsafe_allow_html=True)


def issue_rows(open_only: bool = True) -> list[dict]:
    where = "WHERE i.status NOT IN ('확인완료','예외인정')" if open_only else ""
    with db() as conn:
        return [dict(r) for r in conn.execute(f"""
            SELECT i.*, t.expense_date, t.amount, t.vendor, t.purpose, t.employee, t.source_file, t.source_row
            FROM issues i JOIN transactions t ON t.id=i.transaction_id
            {where}
            ORDER BY CASE i.severity WHEN '오류' THEN 1 WHEN '주의' THEN 2 ELSE 3 END,
                     CASE WHEN i.due_date IS NULL THEN 1 ELSE 0 END, i.due_date, i.id DESC
        """)]


def menu_label(key: str, fallback: str) -> str:
    menu = {m['menu_key']: m for m in get_menus(include_disabled=True)}.get(key)
    return menu['label'] if menu else fallback


def reset_demo() -> None:
    with db() as conn:
        for table in ("checklist_confirmations", "issues", "documents", "transactions", "imports", "audit_log", "mapping_profiles"):
            conn.execute(f"DELETE FROM {table}")
    demo = create_demo_data()
    import_transactions(demo.name, demo.read_bytes())


def ensure_demo_seed() -> None:
    # Public portfolio view should be understandable on first open.
    # This seeds only an empty Streamlit DB; the original FastAPI app behavior is unchanged.
    with db() as conn:
        count = conn.execute("SELECT COUNT(*) FROM transactions").fetchone()[0]
    if count == 0:
        demo = create_demo_data()
        try:
            import_transactions(demo.name, demo.read_bytes())
        except ValueError:
            pass


ensure_demo_seed()

# --- Sidebar -----------------------------------------------------------------
settings = get_settings()
with st.sidebar:
    st.markdown("### ◼ LEDGER CHECK")
    st.caption("Accounting Review Portfolio")
    st.markdown("---")
    menus = [m for m in get_menus() if m['menu_key'] != 'settings']
    nav_options = [m['label'] for m in menus] + [menu_label('settings', '관리자 설정'), "사용 방법"]
    current = st.radio("메뉴", nav_options, label_visibility="collapsed")
    st.markdown("---")
    st.markdown("**PUBLIC DEMO**")
    st.caption("Streamlit Community Cloud의 로컬 DB는 영구 저장이 보장되지 않습니다. 실제 운영용은 외부 DB 연결이 필요합니다.")
    if st.button("데모 데이터 초기화", use_container_width=True):
        reset_demo()
        st.success("데모 데이터를 다시 만들었습니다.")
        st.rerun()

label_to_key = {m['label']: m['menu_key'] for m in get_menus()}
page = label_to_key.get(current, 'settings' if current == menu_label('settings','관리자 설정') else 'help')
texts = get_ui_texts()

# --- Dashboard ---------------------------------------------------------------
if page == 'dashboard':
    hero(texts.get("st_dashboard_kicker","01 / REVIEW OVERVIEW"), texts.get("st_dashboard_title", settings.get('program_name', '회계자료 검토 및 보고 보조')), texts.get("dashboard_description","업로드된 회계자료의 품질과 처리대상을 한눈에 확인합니다."))
    summary = dashboard_summary(); quality = data_quality_score()
    metric_cards(summary)
    c1, c2 = st.columns([1.15, 1])
    with c1:
        section("지금 처리해야 할 것", "미완료 이슈 우선순위")
        rows = issue_rows(True)[:10]
        if rows:
            df = pd.DataFrame(rows)[['severity','category','vendor','amount','message','assignee','due_date','status']]
            df['amount'] = df['amount'].map(lambda x: f"{float(x or 0):,.0f}")
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.markdown('<div class="ds-note ds-success">미완료 검토 이슈가 없습니다.</div>', unsafe_allow_html=True)
    with c2:
        section("데이터 품질", "규칙 기반 내부 지표")
        st.metric("품질 점수", f"{quality['score']} / 100", quality['label'])
        st.caption(quality['note'])
        st.markdown(f"<span class='ds-chip red'>오류 {summary['errors']}</span><span class='ds-chip yellow'>주의 {summary['warnings']}</span><span class='ds-chip blue'>확인 {summary['checks']}</span>", unsafe_allow_html=True)
        section("빠른 실행", "검토 결과 내보내기")
        col_a, col_b = st.columns(2)
        with col_a:
            p = create_review_workbook()
            st.download_button("검토결과 Excel", p.read_bytes(), p.name, use_container_width=True)
        with col_b:
            p = create_audit_package()
            st.download_button("감사대응 ZIP", p.read_bytes(), p.name, use_container_width=True)

# --- Upload ------------------------------------------------------------------
elif page == 'upload':
    hero(texts.get("st_upload_kicker","02 / DATA INTAKE"), texts.get("st_upload_title","자료를 먼저 깨끗하게."), texts.get("upload_description","CSV/XLSX의 열을 자동 인식하고 사람이 확인합니다."))
    sample = create_sample_template()
    st.download_button("표준 업로드 양식 받기", sample.read_bytes(), sample.name)
    file = st.file_uploader("회계자료 파일", type=["csv","xlsx","xlsm"], help="첫 행에 열 제목이 있는 표 형식을 권장합니다.")
    if file:
        content = file.getvalue()
        try:
            headers, rows = read_table_bytes(file.name, content)
            auto = detect_mapping(headers)
            sig = hashlib.sha256(json.dumps(headers, ensure_ascii=False, separators=(",", ":")).encode()).hexdigest()
            with db() as conn:
                prof = conn.execute("SELECT * FROM mapping_profiles WHERE header_signature=?", (sig,)).fetchone()
            mapping = auto.copy()
            if prof:
                try:
                    saved = json.loads(prof['mapping_json'])
                    mapping.update({k:v for k,v in saved.items() if v in headers})
                    st.info(f"이전에 확정한 동일 양식 매핑을 불러왔습니다: {prof['name']}")
                except Exception:
                    pass
            section("열 매핑 확인", "자동 인식 결과를 반드시 확인하세요")
            fields = list(DISPLAY_NAMES)
            chosen = {}
            opts = ["(사용 안 함)"] + headers
            cols = st.columns(2)
            for idx, f in enumerate(fields):
                default = mapping.get(f, "(사용 안 함)")
                chosen[f] = cols[idx%2].selectbox(DISPLAY_NAMES[f], opts, index=opts.index(default) if default in opts else 0, key=f"map_{f}")
            section("원본 미리보기", f"{len(rows):,}행 중 앞 5행")
            st.dataframe(pd.DataFrame(rows[:5]), use_container_width=True, hide_index=True)
            if st.button("매핑 확정 및 검토 실행", type="primary", use_container_width=True):
                override = {k:("" if v == "(사용 안 함)" else v) for k,v in chosen.items()}
                try:
                    result = import_transactions(file.name, content, override)
                    with db() as conn:
                        conn.execute("""INSERT INTO mapping_profiles(name,header_signature,headers_json,mapping_json,created_at,last_used_at)
                                      VALUES(?,?,?,?,datetime('now','localtime'),datetime('now','localtime'))
                                      ON CONFLICT(header_signature) DO UPDATE SET mapping_json=excluded.mapping_json,last_used_at=excluded.last_used_at""",
                                     (f"자동 저장 · {Path(file.name).stem[:50]}", sig, json.dumps(headers,ensure_ascii=False), json.dumps(override,ensure_ascii=False)))
                    st.success(f"{result['rows']:,}건을 저장하고 검토 Rule을 실행했습니다.")
                except Exception as e:
                    st.error(str(e))
        except Exception as e:
            st.error(str(e))

# --- Transactions ------------------------------------------------------------
elif page == 'transactions':
    hero(texts.get("st_transactions_kicker","03 / TRANSACTION REVIEW"), texts.get("st_transactions_title","원본 행까지 추적하는 검토."), texts.get("transactions_description","원본 파일과 행번호를 보존한 상태로 검토합니다."))
    f1,f2,f3 = st.columns([2,1,1])
    q = f1.text_input("검색", placeholder="거래처·사용목적·전표번호·사용자")
    status = f2.selectbox("상태", ["","검토필요","정상","미검토"])
    month_opts = [""] + months_available()
    month = f3.selectbox("월", month_opts)
    txs = get_transactions(q, status, month)
    if txs:
        view = pd.DataFrame(txs)
        cols = ['id','expense_date','vendor','purpose','amount','account_name','employee','evidence_status','review_status','issue_count','source_file','source_row']
        st.dataframe(view[cols], use_container_width=True, hide_index=True, column_config={'amount':st.column_config.NumberColumn('금액',format="₩ %.0f")})
        tx_id = st.selectbox("상세 검토할 거래 ID", [t['id'] for t in txs], format_func=lambda x: next(f"#{t['id']} · {t.get('vendor') or '-'} · {won(t.get('amount'))}" for t in txs if t['id']==x))
        detail = get_transaction_detail(int(tx_id))
        if detail:
            with st.expander("거래 상세 / 이슈 처리", expanded=True):
                a,b,c,d = st.columns(4)
                a.metric("일자", detail.get('expense_date') or '-')
                b.metric("금액", won(detail.get('amount')))
                c.metric("거래처", detail.get('vendor') or '-')
                d.metric("원본", f"{detail.get('source_file')} : {detail.get('source_row')}")
                st.write(detail.get('purpose') or '사용목적 없음')
                note = st.text_area("검토 메모", value=detail.get('reviewer_note') or '', key=f"txn_{tx_id}")
                if st.button("검토 메모 저장", key=f"txn_save_{tx_id}"):
                    update_transaction(int(tx_id), note)
                    st.success("저장했습니다.")
                for issue in detail.get('issues',[]):
                    st.markdown(f"**[{issue['severity']}] {issue['category']} · {issue['rule_code']}**  \n{issue['message']}")
                    st.caption(f"상태 {issue['status']} · 담당 {issue.get('assignee') or '-'} · 기한 {issue.get('due_date') or '-'}")
    else:
        st.info("조건에 맞는 거래가 없습니다.")

# --- Issues ------------------------------------------------------------------
elif page == 'issues':
    hero(texts.get("st_issues_kicker","04 / ACTION QUEUE"), texts.get("st_issues_title","발견한 문제를 실제 업무로."), texts.get("issues_description","확인·보완 사항을 실제 처리업무로 관리합니다."))
    rows = issue_rows(True)
    a,b,c = st.columns(3)
    sev = a.selectbox("등급", ["전체","오류","주의","확인"])
    assignee_filter = b.text_input("담당자")
    state = c.selectbox("처리상태", ["전체","미확인","담당자지정","보완요청","재검토"])
    filt = [r for r in rows if (sev=='전체' or r['severity']==sev) and (not assignee_filter or assignee_filter in (r.get('assignee') or '')) and (state=='전체' or r['status']==state)]
    if filt:
        st.dataframe(pd.DataFrame(filt)[['id','severity','category','vendor','amount','message','assignee','due_date','status']], use_container_width=True, hide_index=True)
        issue_id = st.selectbox("처리할 이슈", [r['id'] for r in filt], format_func=lambda x: next(f"#{r['id']} [{r['severity']}] {r.get('vendor') or '-'} · {r['message'][:42]}" for r in filt if r['id']==x))
        current_issue = next(r for r in filt if r['id']==issue_id)
        with st.form("issue_form"):
            c1,c2 = st.columns(2)
            new_status = c1.selectbox("처리상태", ["미확인","담당자지정","보완요청","재검토","확인완료","예외인정"], index=["미확인","담당자지정","보완요청","재검토","확인완료","예외인정"].index(current_issue['status']) if current_issue['status'] in ["미확인","담당자지정","보완요청","재검토","확인완료","예외인정"] else 0)
            assignee = c2.text_input("담당자", value=current_issue.get('assignee') or '')
            due_text = c1.text_input("처리기한 (YYYY-MM-DD)", value=current_issue.get('due_date') or '')
            note = st.text_area("처리근거 / 보완요청 내용", value=current_issue.get('resolution_note') or '')
            if st.form_submit_button("업무 저장", use_container_width=True):
                
                if due_text:
                    try:
                        date.fromisoformat(due_text)
                    except ValueError:
                        st.error("처리기한은 YYYY-MM-DD 형식이어야 합니다.")
                        st.stop()
                resolve_issue(int(issue_id), new_status, note, assignee, due_text)
                st.success("처리상태를 저장했습니다.")
                st.rerun()
        tx = get_transaction_detail(current_issue['transaction_id'])
        if tx:
            section("확인 요청 문구", "코드리스 문구 템플릿 기반")
            st.code(confirmation_message(tx, current_issue), language=None)
    else:
        st.markdown('<div class="ds-note ds-success">조건에 해당하는 미완료 업무가 없습니다.</div>', unsafe_allow_html=True)

# --- Documents ---------------------------------------------------------------
elif page == 'documents':
    hero(texts.get("st_documents_kicker","05 / EVIDENCE"), texts.get("st_documents_title","증빙을 거래와 연결하기."), texts.get("documents_description","증빙번호를 부여해 거래와 연결합니다."))
    u = st.file_uploader("증빙·문서 업로드", type=['pdf','docx','txt','jpg','jpeg','png'])
    c1,c2 = st.columns(2)
    ref = c1.text_input("증빙번호", placeholder="예: AUG-RC-009")
    cat = c2.selectbox("문서분류", ["자동","세금계산서","영수증","거래명세서","계약서","견적서","지출결의","기타"])
    if u and st.button("문서 저장", type="primary"):
        try:
            save_document(u.name, u.getvalue(), ref, "" if cat=="자동" else cat)
            st.success("문서를 저장하고 증빙 연결 Rule을 다시 실행했습니다.")
            st.rerun()
        except Exception as e:
            st.error(str(e))
    docs = list_documents()
    section("등록 문서", f"총 {len(docs)}건")
    if docs:
        st.dataframe(pd.DataFrame(docs)[['id','filename','category','reference_no','summary','uploaded_at']], use_container_width=True, hide_index=True)

# --- Checklist ---------------------------------------------------------------
elif page == 'checklist':
    hero(texts.get("st_checklist_kicker","06 / PRE-SUBMISSION"), texts.get("st_checklist_title","제출 전 마지막 확인."), texts.get("checklist_description","제출 전에 사람이 최종 확인해야 할 항목을 관리합니다."))
    opts = ["전체"] + months_available()
    period = st.selectbox("검토 기간", opts)
    data = checklist_data("" if period=="전체" else period)
    st.markdown(f"<div class='ds-note {'ds-success' if data['ready'] else 'ds-warning'}'><b>{'제출 준비 가능' if data['ready'] else '추가 검토 필요'}</b><br>미완료 이슈 {data['unresolved']:,}건 · 오류 {data['errors']:,}건 · 주의 {data['warnings']:,}건</div>", unsafe_allow_html=True)
    for item in data['items']:
        with st.container(border=True):
            cols = st.columns([0.7,3.5,1.1])
            cols[0].markdown("✅" if item['ok'] else "⬜")
            cols[1].markdown(f"**{item['label']}**  \n{item.get('note') or ''}")
            if item['item_type']=='MANUAL':
                checked = cols[2].checkbox("완료", value=bool(item.get('checked')), key=f"ck_{item['id']}_{period}")
                note = st.text_input("확인 메모", value=item.get('confirmation_note') or '', key=f"cknote_{item['id']}_{period}")
                if st.button("확인 저장", key=f"cksave_{item['id']}_{period}"):
                    set_checklist_confirmation(item['id'], period, checked, note)
                    st.rerun()
    p = create_audit_package("" if period=="전체" else period)
    st.download_button("감사대응 패키지 다운로드", p.read_bytes(), p.name)

# --- Reports -----------------------------------------------------------------
elif page == 'reports':
    hero(texts.get("st_reports_kicker","07 / MONTHLY REPORT"), texts.get("st_reports_title","무엇이 변했는지부터."), texts.get("reports_description","월별 회계자료를 수치 기반으로 요약합니다."))
    months = months_available()
    month = st.selectbox("보고 월", months) if months else ""
    rpt = monthly_report(month)
    if rpt.get('month'):
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("자료 건수", f"{rpt['count']:,}")
        c2.metric("합계금액", won(rpt['amount']), f"{rpt['amount_change_pct']:+.1f}%" if rpt['amount_change_pct'] is not None else None)
        c3.metric("미확인 이슈", f"{rpt['issue_count']:,}")
        c4.metric("오류 등급", f"{rpt['error_count']:,}")
        st.markdown(f"<div class='ds-note'>{rpt['draft']}</div>", unsafe_allow_html=True)
        t1,t2 = st.tabs(["계정과목", "거래처"])
        with t1:
            if rpt['categories']:
                df = pd.DataFrame(rpt['categories']).set_index('name')
                st.bar_chart(df['amount'])
            st.dataframe(pd.DataFrame(rpt['account_changes']), use_container_width=True, hide_index=True)
        with t2:
            if rpt['vendors']:
                df = pd.DataFrame(rpt['vendors']).set_index('name')
                st.bar_chart(df['amount'])
            st.dataframe(pd.DataFrame(rpt['vendor_changes']), use_container_width=True, hide_index=True)
        p = create_review_workbook(month)
        st.download_button("월별 검토보고서 Excel", p.read_bytes(), p.name)
    else:
        st.info("분석할 월 데이터가 없습니다.")

# --- AI ----------------------------------------------------------------------
elif page == 'ai':
    hero(texts.get("st_ai_kicker","08 / AI ASSIST"), texts.get("st_ai_title","AI는 계산기가 아니라 문장 보조자."), texts.get("ai_description","검증된 집계값만 사용해 AI 보고 보조 프롬프트를 만듭니다."))
    months = months_available(); month = st.selectbox("대상 월", months) if months else ""
    info = ai_assist_prompt(month)
    st.markdown('<div class="ds-note">프로그램이 외부 AI로 자료를 자동 전송하지 않습니다. 아래 프롬프트를 사람이 확인한 뒤 활용하는 구조입니다.</div>', unsafe_allow_html=True)
    st.code(info['prompt'], language=None)

# --- Settings ----------------------------------------------------------------
elif page == 'settings':
    hero(texts.get("st_settings_kicker","09 / NO-CODE ADMIN"), texts.get("st_settings_title","코드 없이 운영 기준을 바꾸기."), texts.get("settings_description","코드 수정 없이 화면 문구와 검토 기준을 변경합니다."))
    tabs = st.tabs(["기본 설정","메뉴","화면 문구","검토 Rule","체크리스트"])
    with tabs[0]:
        with st.form("basic_settings"):
            program_name = st.text_input("프로그램명", value=settings.get('program_name',''))
            company_name = st.text_input("보고 회사명", value=settings.get('report_company_name',''))
            if st.form_submit_button("기본 설정 저장"):
                set_settings({'program_name':program_name, 'report_company_name':company_name})
                st.success("저장했습니다."); st.rerun()
    with tabs[1]:
        menus_all = get_menus(include_disabled=True)
        for m in menus_all:
            with st.expander(f"{m['sort_order']} · {m['label']}"):
                c1,c2,c3 = st.columns([2,1,1])
                label = c1.text_input("메뉴명", value=m['label'], key=f"ml_{m['menu_key']}")
                enabled = c2.checkbox("표시", value=bool(m['enabled']), key=f"me_{m['menu_key']}")
                order = c3.number_input("순서", value=int(m['sort_order']), step=1, key=f"mo_{m['menu_key']}")
                if st.button("저장", key=f"ms_{m['menu_key']}"):
                    with db() as conn:
                        conn.execute("UPDATE menus SET label=?,enabled=?,sort_order=? WHERE menu_key=?", (label,1 if enabled else 0,int(order),m['menu_key']))
                    st.success("메뉴를 저장했습니다."); st.rerun()
    with tabs[2]:
        rows = list_ui_texts()
        labels = [f"{r['group_name']} · {r['key']}" for r in rows]
        pick = st.selectbox("수정할 문구", range(len(rows)), format_func=lambda i: labels[i])
        row = rows[pick]
        val = st.text_area("표시 문구", value=row['value'], height=130)
        if st.button("문구 저장"):
            set_ui_text(row['key'], val); st.success("저장했습니다."); st.rerun()
    with tabs[3]:
        rules = get_rules(True)
        rule_pick = st.selectbox("Rule", range(len(rules)), format_func=lambda i: f"{rules[i]['code']} · {rules[i]['name']}")
        r = rules[rule_pick]
        c1,c2,c3 = st.columns(3)
        enabled = c1.checkbox("사용", value=bool(r['enabled']))
        severity = c2.selectbox("등급", ["오류","주의","확인"], index=["오류","주의","확인"].index(r['severity']))
        compare = c3.text_input("기준값", value=r.get('compare_value') or '')
        name = st.text_input("Rule 이름", value=r['name'])
        message = st.text_area("표시 문구", value=r['message'])
        if st.button("Rule 저장", type="primary"):
            if r['rule_type']=='VAT_MISMATCH' or (r['rule_type']=='FIELD' and r.get('operator') in {'gt','gte','lt','lte'} and r.get('field_name') in {'amount','supply_amount','tax_amount'}):
                try: float(compare)
                except ValueError: st.error("이 Rule의 기준값은 숫자여야 합니다."); st.stop()
            with db() as conn:
                conn.execute("UPDATE review_rules SET name=?,enabled=?,severity=?,compare_value=?,message=? WHERE code=?", (name,1 if enabled else 0,severity,compare,message,r['code']))
            validate_all(); st.success("Rule을 저장하고 전체 자료를 재검토했습니다."); st.rerun()
    with tabs[4]:
        with db() as conn:
            items = [dict(x) for x in conn.execute("SELECT * FROM checklist_items ORDER BY sort_order,id")]
        st.dataframe(pd.DataFrame(items), use_container_width=True, hide_index=True)
        st.caption("기존 FastAPI 관리자 화면에서는 체크리스트 항목 추가·수정·삭제까지 지원합니다. Streamlit 공개 데모에서는 현재 상태와 수동 확인을 중심으로 제공합니다.")

# --- Help --------------------------------------------------------------------
else:
    hero(texts.get("st_help_kicker","HOW TO USE"), texts.get("st_help_title","5분이면 흐름을 볼 수 있습니다."), "샘플자료를 넣고 자동검토 → 증빙 연결 → 이슈 처리 → 월별 보고 → 감사대응 자료까지 순서대로 체험해 보세요.")
    st.markdown("""
### 추천 체험 순서
1. **자료 업로드** → `sample_data/02_오류포함_검토연습_2026-08.xlsx` 업로드
2. **회계자료 검토** → 누락·중복·금액·증빙 이슈 확인
3. **확인·보완 업무** → 담당자와 처리기한 지정
4. **증빙·문서 관리** → `AUG-RC-009` 영수증을 연결해 증빙 누락 Rule 변화 확인
5. **제출 전 체크리스트** → 자동 Rule + 수동 확인
6. **월별 보고** → 계정과목/거래처 증감 기여 확인
7. **AI 보고 보조** → 검증된 숫자만 포함된 보고 프롬프트 확인
8. **관리자 설정** → 프로그램명·메뉴·문구·Rule 기준을 코드 없이 변경

### 중요한 한계
- 이 앱은 회계·세무·감사인의 전문적 판단을 대체하지 않습니다.
- 이미지 증빙의 OCR 자동확정은 하지 않습니다.
- Streamlit Community Cloud의 로컬 SQLite는 **영구 저장이 보장되지 않습니다.** 공개 포트폴리오 데모로 사용하고, 실제 운영 시 외부 DB를 연결해야 합니다.
""")
