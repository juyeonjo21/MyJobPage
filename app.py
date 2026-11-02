"""스탠다드뱅크 인증 프로젝트 접수·진행현황 대시보드 (교육용)

실행: streamlit run app.py
- 외부 시스템 연동·메일 발송 없음. 로컬 CSV만 읽습니다.
- 시험분야 관리로 바뀐 데이터는 session_state에만 보관하며 원본 CSV는 수정하지 않습니다.
"""
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
# 과제 명세 파일명과 실제 배포 파일명 모두 지원
CSV_CANDIDATES = ["SB_projects_demo.csv", "SB_project_demo.csv"]

COLUMNS = [
    "의뢰번호", "고객사", "제품명", "모델명", "시험분야", "적용규격",
    "담당엔지니어", "접수일", "완료예정일", "진행상태", "시료상태", "메모",
]
STATUSES = ["접수", "자료확인", "시험진행", "성적서검토", "보완요청", "완료"]
FILTER_FIELDS = ["시험분야", "적용규격", "담당엔지니어", "진행상태", "시료상태"]
INITIAL_CATEGORIES = ["EMC", "무선(RF)", "인증서류", "SAR"]
DATE_FORMAT = "%Y-%m-%d"
DEFAULT_BASE_DATE = date(2026, 10, 23)

# 일정 상태 라벨
SCHEDULE_DONE = "완료"
SCHEDULE_OVERDUE = "지연 후보(일정 확인 필요)"
SCHEDULE_TODAY = "오늘 마감"
SCHEDULE_PLANNED = "진행 예정"
SCHEDULE_CHECK = "일정 확인 필요"
SCHEDULE_LABELS = [SCHEDULE_OVERDUE, SCHEDULE_TODAY, SCHEDULE_PLANNED, SCHEDULE_CHECK, SCHEDULE_DONE]

EXCEL_NOTICE = "교육용 검토 파일 — 공식 보고서 아님"
SESSION_NOTICE = (
    "시험분야 관리로 변경한 내용은 현재 브라우저 세션(session_state)에만 보관됩니다. "
    "새로고침하거나 세션이 종료되면 사라질 수 있으며, 원본 CSV는 수정되지 않습니다."
)

NAVY = "#0B2545"
BLUE = "#1F5FBF"
LIGHT_BLUE = "#E8F0FB"
STATUS_COLORS = ["#A9C3EC", "#7BA2E0", "#4A80D3", BLUE, "#D9822B", NAVY]

st.set_page_config(page_title="인증 프로젝트 진행현황", page_icon="📋", layout="wide")

st.markdown(
    f"""
    <style>
    .stApp {{ background-color: #FFFFFF; }}
    h1, h2, h3 {{ color: {NAVY}; }}
    .notice {{
        background: {LIGHT_BLUE}; color: {NAVY};
        border-left: 5px solid {BLUE}; border-radius: 6px;
        padding: 10px 14px; font-weight: 600; margin-bottom: 12px;
    }}
    .card {{
        background: #FFFFFF; border: 1px solid #D6E2F3;
        border-top: 4px solid {BLUE}; border-radius: 8px;
        padding: 12px 10px; text-align: center;
    }}
    .card.total {{ border-top-color: {NAVY}; background: {NAVY}; }}
    .card .label {{ font-size: 0.9rem; color: #4A5A75; }}
    .card .value {{ font-size: 1.8rem; font-weight: 700; color: {NAVY}; }}
    .card.total .label, .card.total .value {{ color: #FFFFFF; }}
    </style>
    """,
    unsafe_allow_html=True,
)


def read_csv(source) -> pd.DataFrame:
    """UTF-8(BOM 포함) 우선, 실패 시 CP949로 읽기."""
    for enc in ("utf-8-sig", "cp949"):
        try:
            if hasattr(source, "seek"):
                source.seek(0)
            return pd.read_csv(source, encoding=enc, dtype=str)
        except UnicodeDecodeError:
            continue
    raise ValueError("CSV 인코딩을 인식할 수 없습니다 (UTF-8 또는 CP949 권장).")


def prepare(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.str.strip()
    missing = [c for c in COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"필수 필드가 없습니다: {', '.join(missing)}")
    df = df[COLUMNS].fillna("")
    for col in COLUMNS:
        df[col] = df[col].str.strip()
    # YYYY-MM-DD 가 아니거나 비어 있으면 NaT
    for col in ("접수일", "완료예정일"):
        df[col] = pd.to_datetime(df[col], format=DATE_FORMAT, errors="coerce")
    return df


def schedule_status(progress: str, remaining) -> str:
    """일정 상태 판정. 판정 순서가 곧 우선순위입니다."""
    if progress == "완료":
        return SCHEDULE_DONE  # 예정일이 지났어도 지연 아님
    if pd.isna(remaining):
        return SCHEDULE_CHECK  # 완료예정일 누락/형식 오류
    if remaining < 0:
        return SCHEDULE_OVERDUE
    if remaining == 0:
        return SCHEDULE_TODAY
    return SCHEDULE_PLANNED


def add_schedule(df: pd.DataFrame, base_date: date) -> pd.DataFrame:
    df = df.copy()
    df["잔여일수"] = (df["완료예정일"] - pd.Timestamp(base_date)).dt.days.astype("Int64")
    df["일정 상태"] = [
        schedule_status(p, r) for p, r in zip(df["진행상태"], df["잔여일수"])
    ]
    return df


def find_local_csv() -> Path | None:
    for name in CSV_CANDIDATES:
        path = DATA_DIR / name
        if path.exists():
            return path
    return None


def options_for(df: pd.DataFrame, field: str) -> list[str]:
    if field == "시험분야":
        return list(st.session_state.categories)
    values = [v for v in df[field].unique() if v]
    if field == "진행상태":
        return [s for s in STATUSES if s in values] + sorted(set(values) - set(STATUSES))
    return sorted(values)


def display_table(df: pd.DataFrame) -> pd.DataFrame:
    table = df[COLUMNS[:9] + ["잔여일수", "일정 상태", "진행상태", "시료상태", "메모"]].copy()
    for col in ("접수일", "완료예정일"):
        table[col] = table[col].dt.date
    return table


# ---------- 시험분야 관리 (session_state) ----------
def init_state(df: pd.DataFrame, source_label: str) -> None:
    """데이터 출처가 바뀔 때만 세션 데이터를 새로 만든다."""
    if st.session_state.get("data_source") == source_label:
        return
    st.session_state.data_source = source_label
    st.session_state.projects = df
    extras = [v for v in df["시험분야"].unique() if v and v not in INITIAL_CATEGORIES]
    st.session_state.categories = INITIAL_CATEGORIES + sorted(extras)


def flash(kind: str, message: str) -> None:
    st.session_state.flash = (kind, message)


def usage_count(category: str) -> int:
    return int((st.session_state.projects["시험분야"] == category).sum())


def replace_in_filter(old: str, new: str | None) -> None:
    key = "filter_시험분야"
    selected = st.session_state.get(key, [])
    if old in selected:
        st.session_state[key] = [new if v == old else v for v in selected if new or v != old]


def add_category() -> None:
    name = st.session_state.cat_new.strip()
    categories = st.session_state.categories
    if not name:
        flash("warning", "등록할 시험분야 이름을 입력하세요.")
    elif name in categories:
        flash("warning", f"'{name}'은(는) 이미 등록된 시험분야입니다.")
    else:
        categories.append(name)
        st.session_state.cat_new = ""
        flash("success", f"시험분야 '{name}'을(를) 등록했습니다.")


def rename_category() -> None:
    old = st.session_state.cat_rename_from
    new = st.session_state.cat_rename_to.strip()
    categories = st.session_state.categories
    if not new:
        flash("warning", "새 이름을 입력하세요.")
    elif new == old:
        flash("warning", "기존 이름과 같습니다.")
    elif new in categories:
        flash("warning", f"'{new}'은(는) 이미 등록된 시험분야입니다. 합치려면 삭제 탭의 재분류를 사용하세요.")
    else:
        categories[categories.index(old)] = new
        projects = st.session_state.projects
        changed = usage_count(old)
        projects.loc[projects["시험분야"] == old, "시험분야"] = new
        replace_in_filter(old, new)
        st.session_state.cat_rename_from = new
        st.session_state.cat_rename_to = ""
        flash("success", f"'{old}' → '{new}' 이름을 변경했습니다. 관련 의뢰 {changed}건도 함께 갱신했습니다.")


def reclassify_category() -> None:
    src = st.session_state.cat_delete
    dst = st.session_state.cat_reclass_to
    projects = st.session_state.projects
    moved = usage_count(src)
    projects.loc[projects["시험분야"] == src, "시험분야"] = dst
    flash("success", f"'{src}' 의뢰 {moved}건을 '{dst}'(으)로 재분류했습니다. 이제 '{src}'을(를) 삭제할 수 있습니다.")


def delete_category() -> None:
    name = st.session_state.cat_delete
    if usage_count(name):
        flash("error", f"'{name}'은(는) 사용 중이라 삭제할 수 없습니다. 먼저 다른 분야로 재분류하세요.")
        return
    st.session_state.categories.remove(name)
    replace_in_filter(name, None)
    del st.session_state["cat_delete"]
    flash("success", f"시험분야 '{name}'을(를) 삭제했습니다.")


def render_category_manager() -> None:
    if "flash" in st.session_state:
        kind, message = st.session_state.pop("flash")
        getattr(st, kind)(message)

    categories = st.session_state.categories
    usage = pd.DataFrame(
        {"시험분야": categories, "의뢰 건수": [usage_count(c) for c in categories]}
    )
    st.dataframe(usage, hide_index=True, width="stretch")

    tab_add, tab_rename, tab_delete = st.tabs(["등록", "이름 변경", "삭제"])
    with tab_add:
        st.text_input("새 시험분야", key="cat_new", placeholder="예: 전기안전")
        st.button("등록", on_click=add_category, width="stretch")
    with tab_rename:
        if not categories:
            st.caption("등록된 시험분야가 없습니다.")
        else:
            st.selectbox("변경할 시험분야", categories, key="cat_rename_from")
            st.text_input("새 이름", key="cat_rename_to")
            st.caption("관련 의뢰의 시험분야도 함께 바뀝니다.")
            st.button("이름 변경", on_click=rename_category, width="stretch")
    with tab_delete:
        if not categories:
            st.caption("등록된 시험분야가 없습니다.")
            return
        target = st.selectbox("삭제할 시험분야", categories, key="cat_delete")
        in_use = usage_count(target)
        if in_use:
            st.warning(
                f"'{target}'은(는) 의뢰 {in_use}건에서 사용 중이라 바로 삭제할 수 없습니다. "
                "아래에서 다른 분야로 재분류한 뒤 삭제하세요."
            )
            others = [c for c in categories if c != target]
            if others:
                st.selectbox("재분류할 시험분야", others, key="cat_reclass_to")
                st.button("재분류", on_click=reclassify_category, width="stretch")
            else:
                st.caption("재분류할 다른 시험분야가 없습니다. 먼저 새 분야를 등록하세요.")
        st.button("삭제", on_click=delete_category, disabled=bool(in_use), width="stretch")


# ---------- Excel 내보내기 ----------
def status_summary(df: pd.DataFrame) -> pd.DataFrame:
    total = len(df)
    rows = []
    for group, field, labels in (
        ("진행상태", "진행상태", STATUSES),
        ("일정 상태", "일정 상태", SCHEDULE_LABELS),
    ):
        counts = df[field].value_counts()
        for label in labels:
            n = int(counts.get(label, 0))
            rows.append({"구분": group, "상태": label, "건수": n,
                         "비율(%)": round(n / total * 100, 1) if total else 0.0})
    return pd.DataFrame(rows)


def engineer_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["담당엔지니어", "전체", *STATUSES, "지연 후보"])
    engineers = df["담당엔지니어"].replace("", "(미지정)")
    table = pd.crosstab(engineers, df["진행상태"]).reindex(columns=STATUSES, fill_value=0)
    table.insert(0, "전체", table.sum(axis=1))
    table["지연 후보"] = (df["일정 상태"] == SCHEDULE_OVERDUE).groupby(engineers).sum()
    table = table.sort_index()
    table.loc["합계"] = table.sum()
    return table.rename_axis("담당엔지니어").reset_index()


def conditions_sheet(scope: str, base_date: date, selections: dict, rows: int,
                     source_label: str, generated: datetime) -> pd.DataFrame:
    items = [
        ("문서 구분", EXCEL_NOTICE),
        ("데이터 범위", scope),
        ("기준일", base_date.strftime(DATE_FORMAT)),
    ]
    for field in FILTER_FIELDS:
        chosen = selections.get(field) or []
        items.append((f"필터: {field}", ", ".join(chosen) if chosen else "전체(미적용)"))
    items += [
        ("대상 건수", f"{rows}건"),
        ("일정 상태 기준", "잔여일수 = 완료예정일 - 기준일 / 진행상태 '완료'는 지연 아님"),
        ("데이터 출처", source_label),
        ("시험분야 변경", "세션 중 변경 사항 반영(원본 CSV 미수정)"),
        ("생성 일시", generated.strftime("%Y-%m-%d %H:%M:%S")),
    ]
    return pd.DataFrame(items, columns=["항목", "값"])


def build_excel(projects: pd.DataFrame, conditions: pd.DataFrame) -> bytes:
    sheets = {
        "프로젝트목록": display_table(projects),
        "상태별집계": status_summary(projects),
        "담당자별집계": engineer_summary(projects),
        "분석조건": conditions,
    }
    header_fill = PatternFill("solid", fgColor=NAVY.lstrip("#"))
    overdue_fill = PatternFill("solid", fgColor="FDECEC")
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False, startrow=2)
            ws = writer.sheets[name]
            ws["A1"] = EXCEL_NOTICE
            ws["A1"].font = Font(bold=True, size=12, color="C00000")
            for cell in ws[3]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = header_fill
            ws.freeze_panes = "A4"
            for idx, col in enumerate(frame.columns, start=1):
                lengths = [len(str(col))] + [len(str(v)) for v in frame[col].tolist()]
                ws.column_dimensions[get_column_letter(idx)].width = min(max(lengths) * 1.8 + 2, 60)
                if col in ("접수일", "완료예정일"):
                    for (cell,) in ws.iter_rows(min_row=4, min_col=idx, max_col=idx):
                        cell.number_format = "yyyy-mm-dd"
            if name == "프로젝트목록" and not frame.empty:
                status_col = list(frame.columns).index("일정 상태")
                for row in ws.iter_rows(min_row=4):
                    if row[status_col].value == SCHEDULE_OVERDUE:
                        for cell in row:
                            cell.fill = overdue_fill
    return buffer.getvalue()


# ---------- 헤더 ----------
st.markdown(
    '<div class="notice">⚠️ 교육용 가상 데이터 기준 — 실제 업무 시스템 아님</div>',
    unsafe_allow_html=True,
)
st.title("📋 인증 프로젝트 접수 · 진행현황 대시보드")

# ---------- 데이터 로드 ----------
local_csv = find_local_csv()
if local_csv:
    raw = read_csv(local_csv)
    source_label = f"data/{local_csv.name}"
else:
    st.info(
        f"`data/{CSV_CANDIDATES[0]}` 파일을 찾을 수 없습니다. "
        "아래에서 CSV 파일을 업로드해 주세요."
    )
    uploaded = st.file_uploader("프로젝트 CSV 업로드", type=["csv"])
    if uploaded is None:
        st.caption("필드: " + ", ".join(COLUMNS))
        st.stop()
    raw = read_csv(uploaded)
    source_label = f"업로드: {uploaded.name}"

try:
    init_state(prepare(raw), source_label)
except ValueError as e:
    st.error(str(e))
    st.stop()
df = st.session_state.projects

st.caption(f"💾 {SESSION_NOTICE}")

# ---------- 사이드바: 기준일 / 필터 / 카테고리 관리 ----------
with st.sidebar:
    st.header("🗓️ 조회 조건")
    base_date = st.date_input("기준일", value=DEFAULT_BASE_DATE)
    st.caption("기준일 이후 접수된 의뢰는 집계에서 제외합니다.")

    st.divider()
    st.header("🔎 필터")
    st.caption("선택하지 않으면 전체가 대상입니다.")
    selections = {}
    for field in FILTER_FIELDS:
        options = options_for(df, field)
        key = f"filter_{field}"
        if key in st.session_state:  # 이름 변경·삭제로 사라진 값 정리
            st.session_state[key] = [v for v in st.session_state[key] if v in options]
        selections[field] = st.multiselect(field, options, key=key, placeholder="전체")

    st.divider()
    st.header("🗂️ 카테고리 관리")
    with st.container(border=True):
        st.subheader("시험분야 관리")
        render_category_manager()
        st.caption(f"⚠️ {SESSION_NOTICE}")

    st.divider()
    st.caption(f"데이터 출처: {source_label}")

# 접수일 형식 오류 행은 숨기지 않고 포함 (일정 상태로 드러나도록)
in_scope = df[~(df["접수일"] > pd.Timestamp(base_date))]
mask = pd.Series(True, index=in_scope.index)
for field, chosen in selections.items():
    if chosen:
        mask &= in_scope[field].isin(chosen)
view = add_schedule(in_scope[mask], base_date)

active = [f"{f}: {', '.join(v)}" for f, v in selections.items() if v]
st.caption(
    f"기준일 **{base_date:%Y-%m-%d}** · 조회 {len(view)}건 / 전체 {len(df)}건"
    + (f" · 필터 — {' | '.join(active)}" if active else "")
)

# ---------- 요약 카드 ----------
counts = view["진행상태"].value_counts()
cards = [("전체", len(view), "total")] + [(s, int(counts.get(s, 0)), "") for s in STATUSES]
for col, (label, value, cls) in zip(st.columns(len(cards)), cards):
    col.markdown(
        f'<div class="card {cls}"><div class="label">{label}</div>'
        f'<div class="value">{value}</div></div>',
        unsafe_allow_html=True,
    )

# ---------- Excel 다운로드 ----------
generated = datetime.now()
stamp = generated.strftime("%Y%m%d")
with st.expander("📥 Excel 다운로드 (.xlsx)", expanded=False):
    st.caption(
        "시트: 프로젝트목록 / 상태별집계 / 담당자별집계 / 분석조건 · "
        f"각 시트 상단에 '{EXCEL_NOTICE}' 표시"
    )
    dl_all, dl_filtered = st.columns(2)
    all_projects = add_schedule(df, base_date)
    dl_all.download_button(
        f"전체 목록 ({len(all_projects)}건)",
        data=build_excel(
            all_projects,
            conditions_sheet("전체 목록 (필터·기준일 접수 제외 미적용)", base_date, {},
                             len(all_projects), source_label, generated),
        ),
        file_name=f"standardbank_project_status_{stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        on_click="ignore",
        width="stretch",
    )
    dl_filtered.download_button(
        f"현재 필터 결과 ({len(view)}건)",
        data=build_excel(
            view,
            conditions_sheet("현재 필터 결과 (기준일 이후 접수 제외)", base_date, selections,
                             len(view), source_label, generated),
        ),
        file_name=f"standardbank_project_status_filtered_{stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        on_click="ignore",
        disabled=view.empty,
        width="stretch",
    )

if view.empty:
    st.info("조건에 맞는 인증 프로젝트가 없습니다")
    st.stop()

schedule_counts = view["일정 상태"].value_counts()
if schedule_counts.get(SCHEDULE_OVERDUE, 0):
    st.warning(
        f"완료예정일이 지난 미완료 의뢰(지연 후보)가 {schedule_counts[SCHEDULE_OVERDUE]}건 있습니다. "
        "일정을 확인해 주세요."
    )
if schedule_counts.get(SCHEDULE_CHECK, 0):
    st.info(f"완료예정일이 비어 있거나 형식이 올바르지 않은 의뢰가 {schedule_counts[SCHEDULE_CHECK]}건 있습니다.")

# ---------- 차트 ----------
chart_left, chart_right = st.columns(2)

with chart_left:
    st.subheader("시험분야별 건수")
    field_counts = view["시험분야"].replace("", "(미입력)").value_counts().reset_index()
    field_counts.columns = ["시험분야", "건수"]
    bar = (
        alt.Chart(field_counts)
        .mark_bar(color=BLUE, cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
        .encode(
            x=alt.X("시험분야:N", sort="-y", axis=alt.Axis(labelAngle=0)),
            y=alt.Y("건수:Q", axis=alt.Axis(tickMinStep=1)),
            tooltip=["시험분야", "건수"],
        )
    )
    labels = bar.mark_text(dy=-8, color=NAVY, fontWeight="bold").encode(text="건수:Q")
    st.altair_chart(bar + labels, width="stretch")

with chart_right:
    st.subheader("진행상태 비율")
    status_counts = view["진행상태"].value_counts().reset_index()
    status_counts.columns = ["진행상태", "건수"]
    status_counts["비율"] = status_counts["건수"] / status_counts["건수"].sum()
    donut = (
        alt.Chart(status_counts)
        .mark_arc(innerRadius=60, stroke="#FFFFFF")
        .encode(
            theta="건수:Q",
            color=alt.Color(
                "진행상태:N",
                scale=alt.Scale(domain=STATUSES, range=STATUS_COLORS),
                sort=STATUSES,
            ),
            tooltip=["진행상태", "건수", alt.Tooltip("비율:Q", format=".0%")],
        )
    )
    st.altair_chart(donut, width="stretch")

# ---------- 의뢰 목록 ----------
st.subheader("전체 의뢰 목록")


def highlight_overdue(row):
    style = "background-color: #FDECEC; color: #9B1C1C" if row["일정 상태"] == SCHEDULE_OVERDUE else ""
    return [style] * len(row)


st.dataframe(
    display_table(view).style.apply(highlight_overdue, axis=1),
    width="stretch",
    hide_index=True,
    column_config={
        "접수일": st.column_config.DateColumn(format="YYYY-MM-DD"),
        "완료예정일": st.column_config.DateColumn(format="YYYY-MM-DD"),
        "잔여일수": st.column_config.NumberColumn(help="완료예정일 - 기준일 (일)", format="%d"),
    },
)
