import gspread
from google.oauth2.service_account import Credentials

@st.cache_resource
def init_connection():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )
    client = gspread.authorize(creds)
    # スプレッドシートURLの /d/ から /edit の間にある文字列を指定
    return client.open_by_key("あなたのスプレッドシートのID").sheet1

worksheet = init_connection()

import io
from datetime import date, datetime

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# ==========================================
# 1. アプリの基本設定
# ==========================================
st.set_page_config(
    page_title="BS活動評価システム app8",
    page_icon="🏕️",
    layout="wide",
)

st.markdown(
    """
    <style>
    .bs-color { color: #00664B; font-weight: bold; }
    hr { margin: 2em 0; border: none; border-top: 2px dashed #ccc; }
    </style>
    """,
    unsafe_allow_html=True,
)

LEADERS = ["隊長", "副長", "副長補", "インストラクター"]
SCOUTS_PATROL_1 = [
    "スカウトA (シロクマ班)",
    "スカウトB (シロクマ班)",
    "スカウトC (シロクマ班)",
]
SCOUTS_PATROL_2 = [
    "スカウトD (トナカイ班)",
    "スカウトE (トナカイ班)",
    "スカウトF (トナカイ班)",
]
SCOUTS_ALL = SCOUTS_PATROL_1 + SCOUTS_PATROL_2

BS_RANKS = ["初級", "2級", "1級", "菊"]
SKILL_CATEGORIES = [
    "基本 (ちかい・おきて・国旗など)",
    "健康と発達 (救急・体力向上など)",
    "スカウト技能 (ロープ・読図・キャンプなど)",
    "善行・奉仕",
    "信仰奨励",
    "班長会議",
]
MERIT_BADGES = [
    "なし", "読図章", "野営章", "野外炊事章", "パイオニアリング章",
    "リーダーシップ章", "ハイキング章", "スカウトソング章", "救急章", "その他",
]

DATA_COLUMNS = [
    "評価種別", "日付", "活動対象", "場所", "当番班", "テーマ",
    "班議題1", "班議題2", "班議題3", "班議題4", "班議題5",
    "GB議題1", "GB議題2", "GB議題3", "タイムテーブル",
    "担当者1", "担当者2", "スカウト名", "技能評価", "関心評価",
    "進級目標", "課目カテゴリ", "具体的な取り組み", "技能章",
    "指導者1コメント", "指導者2コメント",
]

if "eval_data" not in st.session_state:
    st.session_state.eval_data = pd.DataFrame(columns=DATA_COLUMNS)

st.title("🏕️ ボーイスカウト活動評価")
st.markdown(
    '<p class="bs-color">■ ボーイスカウト部門 専用システム</p>',
    unsafe_allow_html=True,
)

# ==========================================
# 2. 計画情報の状態管理
# ウィジェットを作る前に session_state を更新することが重要
# ==========================================
def default_plan_values():
    return {
        "日時": date.today(),
        "場所": "",
        "テーマ": "",
        "当番班": "シロクマ班",
        "班議題1": "",
        "班議題2": "",
        "班議題3": "",
        "班議題4": "",
        "班議題5": "",
        "GB議題1": "",
        "GB議題2": "",
        "GB議題3": "",
        "タイムテーブル": "09:00 集合・点検\n12:00 昼食",
    }


def widget_key_map(eval_type):
    return {
        "日時": f"date_{eval_type}",
        "場所": f"loc_{eval_type}",
        "当番班": f"duty_{eval_type}",
        "テーマ": f"theme_{eval_type}",
        "班議題1": f"b1_{eval_type}",
        "班議題2": f"b2_{eval_type}",
        "班議題3": f"b3_{eval_type}",
        "班議題4": f"b4_{eval_type}",
        "班議題5": f"b5_{eval_type}",
        "GB議題1": f"g1_{eval_type}",
        "GB議題2": f"g2_{eval_type}",
        "GB議題3": f"g3_{eval_type}",
        "タイムテーブル": f"time_{eval_type}",
    }


def initialize_plan_widgets(eval_type):
    defaults = default_plan_values()
    for field, key in widget_key_map(eval_type).items():
        if key not in st.session_state:
            st.session_state[key] = defaults[field]


def set_plan_widgets(eval_type, data):
    """対応するウィジェットが生成される前に呼び出す。"""
    for field, key in widget_key_map(eval_type).items():
        if field in data:
            value = data[field]
            if field == "日時" and isinstance(value, datetime):
                value = value.date()
            if field == "日時" and not isinstance(value, date):
                value = date.today()
            if field == "当番班" and value not in ["シロクマ班", "トナカイ班", "(なし)"]:
                value = "シロクマ班"
            st.session_state[key] = value

# ==========================================
# 3. Excel実施計画書の解析
# ==========================================
@st.cache_data(show_spinner=False, max_entries=10)
def parse_plan_excel(file_bytes):
    wb = None
    try:
        wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        ws = wb.active

        # read_only ワークシートでのセル再アクセスを避け、必要範囲を一度に読む
        rows = list(ws.iter_rows(min_row=1, max_row=50, max_col=10, values_only=True))

        def get_val(r, c):
            """元コードと同じ0始まりの行・列番号で値を取得する。"""
            if r >= len(rows) or c >= len(rows[r]):
                return ""
            value = rows[r][c]
            if value is None:
                return ""
            if isinstance(value, str):
                return value.strip()
            return value

        def to_date(value):
            if isinstance(value, datetime):
                return value.date()
            if isinstance(value, date):
                return value
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                # Excelの日付シリアル値
                return pd.to_datetime(
                    value, unit="D", origin="1899-12-30", errors="raise"
                ).date()
            if isinstance(value, str) and value.strip():
                return pd.to_datetime(value.strip(), errors="raise").date()
            # 日付が空欄の場合のみ今日を初期値として使用
            return date.today()

        def as_text(value):
            return "" if value is None else str(value).strip()

        def clean_agenda(value, prefix):
            text = as_text(value)
            return text.replace(prefix, "", 1).strip() if text else ""

        duty_raw = as_text(get_val(5, 2))
        default_duty = "シロクマ班"
        if "トナカイ" in duty_raw:
            default_duty = "トナカイ班"
        elif "なし" in duty_raw:
            default_duty = "(なし)"

        date_raw = get_val(2, 3)
        try:
            parsed_date = to_date(date_raw)
        except (ValueError, TypeError, OverflowError):
            # 不正な日付を黙って今日に置き換えず、原因がわかるようにする
            raise ValueError(f"活動日のセルを日付として読み取れません: {date_raw!r}")

        timetable_lines = []
        for r in range(16, 50):
            time_val = as_text(get_val(r, 2))
            prog_val = as_text(get_val(r, 3))
            content_val = as_text(get_val(r, 5))
            if time_val or prog_val:
                line = f"{time_val} {prog_val}".strip()
                if content_val and content_val.lower() != "nan":
                    line += f" ({content_val})"
                timetable_lines.append(line)

        timetable_str = "\n".join(timetable_lines)
        if not timetable_str:
            timetable_str = "09:00 集合\n12:00 昼食"

        return {
            "日時": parsed_date,
            "場所": as_text(get_val(3, 2)),
            "テーマ": as_text(get_val(3, 6)),
            "当番班": default_duty,
            "班議題1": clean_agenda(get_val(8, 3), "議題①："),
            "班議題2": clean_agenda(get_val(9, 3), "議題②："),
            "班議題3": clean_agenda(get_val(10, 3), "議題③："),
            "班議題4": clean_agenda(get_val(11, 3), "議題④："),
            "班議題5": clean_agenda(get_val(12, 3), "議題⑤："),
            "GB議題1": clean_agenda(get_val(13, 3), "訓練①："),
            "GB議題2": clean_agenda(get_val(14, 3), "訓練②："),
            "GB議題3": clean_agenda(get_val(15, 3), "訓練③："),
            "タイムテーブル": timetable_str,
        }
    except Exception as exc:
        return {"エラー": str(exc)}
    finally:
        if wb is not None:
            wb.close()

# ==========================================
# 4. 評価フォーム
# ==========================================
def render_evaluation_form(eval_type):
    initialize_plan_widgets(eval_type)

    # クリア・Excel反映は入力ウィジェットの生成前に処理する
    col_spacer, col_clear = st.columns([4, 1])
    with col_clear:
        clear_clicked = st.button("🗑️ 入力値をクリア", key=f"clear_{eval_type}")
    if clear_clicked:
        set_plan_widgets(eval_type, default_plan_values())
        st.rerun()

    uploaded_plan = st.file_uploader(
        "📂 実施計画書（Excel）を選択・ドロップ",
        type=["xlsx"],
        key=f"up_{eval_type}",
    )
    if uploaded_plan is not None:
        if st.button(
            "🔄 計画書のデータを入力枠に反映する",
            type="primary",
            key=f"btn_apply_{eval_type}",
        ):
            with st.spinner("計画書を読み込んでいます..."):
                parsed = parse_plan_excel(uploaded_plan.getvalue())
            if "エラー" in parsed:
                st.error(f"Excelの解析に失敗しました: {parsed['エラー']}")
            else:
                set_plan_widgets(eval_type, parsed)
                st.success("計画書のデータを反映しました。")
                st.rerun()

    st.subheader("📌 1. 集会・計画情報")
    activity_type_eval = st.radio(
        "活動対象",
        ["隊全体", "シロクマ班", "トナカイ班"],
        horizontal=True,
        key=f"radio_{eval_type}",
    )
    current_scouts = (
        SCOUTS_ALL if activity_type_eval == "隊全体"
        else SCOUTS_PATROL_1 if activity_type_eval == "シロクマ班"
        else SCOUTS_PATROL_2
    )

    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        date_val = st.date_input("活動日時", key=f"date_{eval_type}")
    with col_m2:
        location = st.text_input("場所", key=f"loc_{eval_type}")
    with col_m3:
        duty_patrol = st.selectbox(
            "当番班",
            ["シロクマ班", "トナカイ班", "(なし)"],
            key=f"duty_{eval_type}",
        )

    theme = st.text_input("テーマ・目的", key=f"theme_{eval_type}")

    col_b1, col_b2 = st.columns(2)
    with col_b1:
        b_agenda1 = st.text_input("班議題①", key=f"b1_{eval_type}")
        b_agenda2 = st.text_input("班議題②", key=f"b2_{eval_type}")
        b_agenda3 = st.text_input("班議題③", key=f"b3_{eval_type}")
    with col_b2:
        b_agenda4 = st.text_input("班議題④", key=f"b4_{eval_type}")
        b_agenda5 = st.text_input("班議題⑤", key=f"b5_{eval_type}")

    col_g1, col_g2, col_g3 = st.columns(3)
    with col_g1:
        g_agenda1 = st.text_input("GB議題①", key=f"g1_{eval_type}")
    with col_g2:
        g_agenda2 = st.text_input("GB議題②", key=f"g2_{eval_type}")
    with col_g3:
        g_agenda3 = st.text_input("GB議題③", key=f"g3_{eval_type}")

    timetable = st.text_area(
        "タイムテーブル（時間・プログラム）",
        height=120,
        key=f"time_{eval_type}",
    )

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("👤 2. スカウト個別評価（1人ずつ連続入力）")

    with st.form(f"form_{eval_type}", clear_on_submit=True):
        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1:
            e_scout = st.selectbox("対象スカウト", current_scouts)
        with col_p2:
            e_leader1 = st.selectbox("指導者1", LEADERS)
        with col_p3:
            e_leader2 = st.selectbox("指導者2", ["(なし)"] + LEADERS)

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            e_skill = st.slider("技能（技術の習得度など）", 1, 5, 3)
        with col_s2:
            e_interest = st.slider("関心（積極性・態度など）", 1, 5, 3)

        col_c1, col_c2, col_c3 = st.columns(3)
        with col_c1:
            e_rank = st.selectbox("進級目標", BS_RANKS)
        with col_c2:
            e_cat = st.selectbox("評価カテゴリ", SKILL_CATEGORIES)
        with col_c3:
            e_badge = st.selectbox("技能章", MERIT_BADGES)

        e_task = st.text_input("具体的な取り組み・細目")
        e_com1 = st.text_area("指導者1からのコメント", height=120)
        e_com2 = st.text_area("指導者2からのコメント（任意）", height=120)

        submitted = st.form_submit_button(
            "＋ このスカウトの評価を記録する",
            type="primary",
            use_container_width=True,
        )

   if submitted:
            new_row = [
                eval_type, str(date_val), activity_type_eval, location, duty_patrol, theme,
                b_agenda1, b_agenda2, b_agenda3, b_agenda4, b_agenda5,
                g_agenda1, g_agenda2, g_agenda3, timetable,
                e_leader1, "" if e_leader2 == "(なし)" else e_leader2, e_scout,
                e_skill, e_interest, e_rank, e_cat, e_task, e_badge, e_com1, e_com2
            ]
            worksheet.append_row(new_row)
            st.success(f"✅ {e_scout} の評価をスプレッドシートに記録しました。")
# ==========================================
# 5. 画面
# ==========================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📝 各集会評価",
    "🏕️ GB集会・訓練評価",
    "📈 進歩・進級記録",
    "📊 データ出力",
    "📶 圏外同期",
])

with tab1:
    render_evaluation_form("各集会評価")
with tab2:
    render_evaluation_form("GB集会・訓練評価")
with tab3:
    render_evaluation_form("進歩・進級記録")

with tab4:
    st.subheader("蓄積データの一覧")
    st.dataframe(st.session_state.eval_data, use_container_width=True)

    if not st.session_state.eval_data.empty:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter", date_format="yyyy-mm-dd") as writer:
            st.session_state.eval_data.to_excel(writer, sheet_name="統合記録", index=False)
        st.download_button(
            "📥 全データをExcelで一括ダウンロード",
            data=buffer.getvalue(),
            file_name=f"BS_Scout_Data_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

with tab5:
    st.subheader("圏外での入力用テンプレートとデータ同期")
    template_df = pd.DataFrame(columns=DATA_COLUMNS)
    template_buffer = io.BytesIO()
    with pd.ExcelWriter(template_buffer, engine="xlsxwriter") as writer:
        template_df.to_excel(writer, sheet_name="入力用", index=False)
    st.download_button(
        "📄 オフライン入力用テンプレートをダウンロード",
        data=template_buffer.getvalue(),
        file_name="Offline_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    uploaded_file = st.file_uploader(
        "オフラインで入力したExcelファイルをアップロード",
        type=["xlsx"],
        key="offline_sync",
    )
    if uploaded_file is not None and st.button("同期（データを追加）する", key="btn_offline_sync"):
        try:
            offline_data = pd.read_excel(
                io.BytesIO(uploaded_file.getvalue()),
                engine="openpyxl",
            )
            if list(offline_data.columns) != DATA_COLUMNS:
                st.error(
                    "Excelの列構成がテンプレートと異なります。"
                    "このアプリからダウンロードしたテンプレートを使用してください。"
                )
            elif offline_data.empty:
                st.warning("追加できるデータがありません。")
            else:
                # 列順を明示してから追加する
                offline_data = offline_data[DATA_COLUMNS]
                st.session_state.eval_data = pd.concat(
                    [st.session_state.eval_data, offline_data],
                    ignore_index=True,
                )
                st.success(f"✅ {len(offline_data)}件のデータを追加しました。")
        except Exception as exc:
            st.error(f"同期に失敗しました: {exc}")

st.caption(
    "注意：この版の記録データはStreamlitのセッション内に保存されます。 "
    "アプリ再起動やセッション切断後も確実に保持するには、外部ファイルやデータベースへの保存機能が必要です。"
)
