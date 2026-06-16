"""esグループ月次会議用 3社財務ダッシュボード"""
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import yaml
from pathlib import Path

CONFIG_DIR = Path(__file__).parent / "config"
DATA_DIR = Path(__file__).parent.parent / "data" / "monthly"
COMPANIES_FILE = CONFIG_DIR / "companies.yaml"

COMPANY_COLORS = {
    "es-entertainment": "#c0392b",
    "s-create": "#2980b9",
    "life-work": "#27ae60",
}
COMPANY_SHORT = {
    "es-entertainment": "esエンタメ",
    "s-create": "エスクリエイト",
    "life-work": "ライフワーク",
}


@st.cache_data(ttl=300)
def load_config():
    with open(COMPANIES_FILE, "r") as f:
        return yaml.safe_load(f)


def find_available_months():
    files = sorted(DATA_DIR.glob("*.xlsx"))
    months = set()
    for f in files:
        import re
        m = re.search(r"(\d{6})", f.name)
        if m:
            months.add(m.group(1))
    return sorted(months, reverse=True)


def find_file(company, yyyymm):
    prefix = company["file_prefix"]
    candidates = [
        DATA_DIR / f"{prefix}-{yyyymm}.xlsx",
        DATA_DIR / f"{prefix}{yyyymm}.xlsx",
    ]
    for c in candidates:
        if c.exists():
            return c
    return None


# ============================================================
# Excel読み込み（キャッシュ付き）
# ============================================================

@st.cache_data(ttl=300)
def read_pl(file_path: str, company_id: str):
    config = load_config()
    company = next(c for c in config["companies"] if c["id"] == company_id)
    sheet = company["sheets"].get("pl", "月次PL")
    return pd.read_excel(file_path, sheet_name=sheet, header=None)


@st.cache_data(ttl=300)
def read_pl_trend(file_path: str, company_id: str):
    config = load_config()
    company = next(c for c in config["companies"] if c["id"] == company_id)
    sheet = company["sheets"].get("pl_trend", "PL推移")
    return pd.read_excel(file_path, sheet_name=sheet, header=None)


@st.cache_data(ttl=300)
def read_bs(file_path: str, company_id: str):
    config = load_config()
    company = next(c for c in config["companies"] if c["id"] == company_id)
    sheet = company["sheets"].get("bs_thousand", "BS千円")
    return pd.read_excel(file_path, sheet_name=sheet, header=None)


@st.cache_data(ttl=300)
def read_cf(file_path: str, company_id: str):
    config = load_config()
    company = next(c for c in config["companies"] if c["id"] == company_id)
    sheet = company["sheets"].get("cf", "CF推移")
    return pd.read_excel(file_path, sheet_name=sheet, header=None)


# ============================================================
# データ抽出
# ============================================================

def extract_monthly_pl(df):
    """月次PLから主要項目の月別推移を抽出する。"""
    results = {}
    targets = {
        "売上高合計": "売上高",
        "売上原価": "原価",
        "売上総利益": "粗利",
    }
    header_row = None
    for i in range(min(5, len(df))):
        row_vals = [str(v) for v in df.iloc[i] if pd.notna(v)]
        if any("月" in v for v in row_vals) or any("勘定科目" in v for v in row_vals):
            header_row = i
            break

    months = []
    if header_row is not None:
        for c in range(2, df.shape[1]):
            val = df.iloc[header_row, c]
            if pd.notna(val):
                s = str(val).replace(".0", "")
                if "月" in s or s.isdigit():
                    months.append(s if "月" in s else f"{s}月")
                elif s == "累計":
                    break

    for row_idx in range(len(df)):
        cell = str(df.iloc[row_idx, 0]) if pd.notna(df.iloc[row_idx, 0]) else ""
        for key, label in targets.items():
            if key in cell:
                vals = []
                for c in range(2, 2 + len(months)):
                    if c < df.shape[1]:
                        v = df.iloc[row_idx, c]
                        vals.append(float(v) if pd.notna(v) and isinstance(v, (int, float)) else 0)
                results[label] = vals
    return months, results


def extract_pl_trend_data(df):
    """PL推移シートから売上・粗利・販管費・営業利益の千円ベース推移を抽出する。"""
    results = {}
    months = []

    header_row = None
    for i in range(len(df)):
        for j in range(df.shape[1]):
            val = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
            if "項目" in val:
                header_row = i
                break
        if header_row is not None:
            break

    if header_row is None:
        return months, results

    for c in range(3, df.shape[1], 2):
        if c < df.shape[1]:
            val = df.iloc[header_row, c]
            if pd.notna(val):
                s = str(val).replace(".0", "")
                if "月" in s or s.isdigit():
                    months.append(s if "月" in s else f"{s}月")

    targets = {"売上高": "売上高", "粗利益": "粗利", "販売管理費": "販管費"}
    for row_idx in range(header_row + 1, len(df)):
        for col_idx in range(df.shape[1]):
            cell = str(df.iloc[row_idx, col_idx]) if pd.notna(df.iloc[row_idx, col_idx]) else ""
            for key, label in targets.items():
                if cell == key:
                    vals = []
                    for m_idx in range(len(months)):
                        data_col = 3 + m_idx * 2
                        if data_col < df.shape[1]:
                            v = df.iloc[row_idx, data_col]
                            vals.append(float(v) if pd.notna(v) and isinstance(v, (int, float)) else 0)
                    results[label] = vals
                    break

    if "売上高" in results and "販管費" in results:
        results["営業利益"] = [
            (results.get("粗利", results["売上高"])[i] - results["販管費"][i])
            if i < len(results.get("粗利", results["売上高"])) and i < len(results["販管費"])
            else 0
            for i in range(len(months))
        ]

    return months, results


def extract_pl_detail(df):
    """PL推移シートから営業損益・経常損益・販管費内訳・営業外損益を抽出する。"""
    months = []
    detail = {}

    header_row = None
    for i in range(len(df)):
        for j in range(df.shape[1]):
            val = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
            if "項目" in val:
                header_row = i
                break
        if header_row is not None:
            break
    if header_row is None:
        return months, detail

    for c in range(3, df.shape[1], 2):
        if c < df.shape[1]:
            val = df.iloc[header_row, c]
            if pd.notna(val):
                s = str(val).replace(".0", "")
                if "月" in s or s.isdigit():
                    months.append(s if "月" in s else f"{s}月")

    def _read_row(row_idx):
        vals = []
        for m_idx in range(len(months)):
            data_col = 3 + m_idx * 2
            if data_col < df.shape[1]:
                v = df.iloc[row_idx, data_col]
                vals.append(float(v) if pd.notna(v) and isinstance(v, (int, float)) else 0)
        return vals

    main_targets = {
        "売上高": "売上高", "売上原価": "売上原価", "粗利益": "粗利",
        "販売管理費": "販管費", "営業損益": "営業利益",
        "営業外収益": "営業外収益", "営業外費用": "営業外費用", "経常損益": "経常利益",
    }
    sga_items = [
        "人件費", "旅費交通費", "通信費", "開発費", "交際費", "会議費",
        "減価償却費", "賃借料", "地代家賃", "リース料", "修繕費", "保険料",
        "水道光熱費", "消耗品費", "租税公課", "広告宣伝費", "支払手数料",
        "新聞図書費", "管理諸費", "雑費", "市場調査費",
    ]

    for row_idx in range(header_row + 1, len(df)):
        cell0 = str(df.iloc[row_idx, 0]) if pd.notna(df.iloc[row_idx, 0]) else ""
        cell1 = str(df.iloc[row_idx, 1]) if df.shape[1] > 1 and pd.notna(df.iloc[row_idx, 1]) else ""

        for key, label in main_targets.items():
            if cell0 == key or cell1 == key:
                detail[label] = _read_row(row_idx)
                break

        for item in sga_items:
            if cell1 == item:
                detail[f"販管費_{item}"] = _read_row(row_idx)

    return months, detail


def extract_segments(df, segments):
    """PL推移シートからセグメント別売上を抽出する。"""
    if not segments:
        return [], {}

    months = []
    header_row = None
    for i in range(len(df)):
        for j in range(df.shape[1]):
            val = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
            if "項目" in val:
                header_row = i
                break
        if header_row is not None:
            break

    if header_row is None:
        return [], {}

    for c in range(3, df.shape[1], 2):
        if c < df.shape[1]:
            val = df.iloc[header_row, c]
            if pd.notna(val):
                s = str(val).replace(".0", "")
                if "月" in s or s.isdigit():
                    months.append(s if "月" in s else f"{s}月")

    seg_data = {}
    for row_idx in range(header_row + 1, header_row + 20):
        if row_idx >= len(df):
            break
        for col_idx in range(df.shape[1]):
            cell = str(df.iloc[row_idx, col_idx]).strip() if pd.notna(df.iloc[row_idx, col_idx]) else ""
            for seg in segments:
                if seg in cell or cell in seg:
                    vals = []
                    for m_idx in range(len(months)):
                        data_col = 3 + m_idx * 2
                        if data_col < df.shape[1]:
                            v = df.iloc[row_idx, data_col]
                            vals.append(float(v) if pd.notna(v) and isinstance(v, (int, float)) else 0)
                    seg_data[seg] = vals
                    break

    return months, seg_data


def extract_bs_trend(df, fy_start=7):
    """BS千円シートから主要残高の月別推移を抽出する。"""
    results = {}
    targets = ["現金及び預金合計", "流動資産合計", "流動負債合計", "純資産の部合計"]

    n_data_cols = df.shape[1] - 2
    months = ["期首残"]
    for i in range(1, n_data_cols):
        m = ((fy_start - 1 + (i - 1)) % 12) + 1
        months.append(f"{m}月")

    for row_idx in range(len(df)):
        cell = str(df.iloc[row_idx, 0]) if pd.notna(df.iloc[row_idx, 0]) else ""
        for target in targets:
            if cell.strip().startswith(target):
                vals = []
                for c in range(2, 2 + len(months)):
                    if c < df.shape[1]:
                        v = df.iloc[row_idx, c]
                        vals.append(float(v) if pd.notna(v) and isinstance(v, (int, float)) else 0)
                results[target] = vals
    return months, results


# BS内訳の階層定義: 各合計項目 → 構成する小項目
BS_DETAIL_MAP = {
    "現金及び預金合計": {
        "sub_items": [
            ("現金", 1), ("普通預金", 1), ("小口現金", 1),
        ],
    },
    "流動資産合計": {
        "sub_items": [
            ("現金及び預金合計", 0), ("売上債権合計", 0),
            ("有価証券合計", 0), ("棚卸資産合計", 0),
            ("その他流動資産合計", 0),
        ],
    },
    "流動負債合計": {
        "sub_items": [
            ("仕入債務合計", 0), ("その他流動負債合計", 0),
        ],
    },
    "純資産の部合計": {
        "sub_items": [
            ("資本金合計", 0), ("資本剰余金合計", 0),
            ("利益剰余金合計", 0), ("株主資本合計", 0),
        ],
    },
}

# さらに深い内訳（合計項目→明細）
BS_SUB_DETAIL_MAP = {
    "売上債権合計": [("売掛金", 1)],
    "棚卸資産合計": [("商品", 1), ("貯蔵品", 1)],
    "その他流動資産合計": [
        ("前払費用", 1), ("立替金", 1), ("未収入金", 1),
        ("仮払金", 1), ("仮払消費税", 1),
    ],
    "仕入債務合計": [("買掛金", 1)],
    "その他流動負債合計": [
        ("未払金", 1), ("未払費用", 1), ("前受金", 1),
        ("預り金", 1), ("仮受金", 1), ("未払消費税", 1),
        ("未払法人税等", 1), ("仮受消費税", 1),
        ("一年以内返済長期借入金", 1),
    ],
    "利益剰余金合計": [("繰越利益剰余金", 0), ("（うち当期純利益）", 0)],
}


def extract_bs_full(df, fy_start=7):
    """BS千円シートから全勘定科目の月別推移を抽出する。"""
    results = {}

    n_data_cols = df.shape[1] - 2
    months = ["期首残"]
    for i in range(1, n_data_cols):
        m = ((fy_start - 1 + (i - 1)) % 12) + 1
        months.append(f"{m}月")

    all_targets = set()
    for info in BS_DETAIL_MAP.values():
        for name, _ in info["sub_items"]:
            all_targets.add(name)
    for subs in BS_SUB_DETAIL_MAP.values():
        for name, _ in subs:
            all_targets.add(name)
    all_targets.add("固定資産合計")
    all_targets.add("固定負債合計")

    for row_idx in range(len(df)):
        c0 = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ""
        c1 = str(df.iloc[row_idx, 1]).strip() if pd.notna(df.iloc[row_idx, 1]) else ""
        for target in all_targets:
            matched = False
            if c0 and c0.startswith(target):
                matched = True
            elif c1 and c1 == target:
                matched = True
            if matched and target not in results:
                vals = []
                for c in range(2, 2 + len(months)):
                    if c < df.shape[1]:
                        v = df.iloc[row_idx, c]
                        vals.append(float(v) if pd.notna(v) and isinstance(v, (int, float)) else 0)
                if any(v != 0 for v in vals):
                    results[target] = vals

    return months, results


def extract_cf_data(df):
    """CF推移シートから営業CF/投資CF/財務CF/月次CFを抽出する。"""
    results = {}
    months = []
    targets = {
        "１．営業活動CF": "営業CF",
        "２．投資活動CF": "投資CF",
        "３．財務活動CF": "財務CF",
        "４．月次CF": "月次CF",
    }

    header_row = None
    for i in range(min(10, len(df))):
        for j in range(df.shape[1]):
            val = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
            if "年度" in val:
                header_row = i
                break
        if header_row is not None:
            break

    if header_row is not None:
        for c in range(3, df.shape[1]):
            val = df.iloc[header_row, c]
            if pd.notna(val):
                s = str(val).replace(".0", "")
                if "月" in s or s.isdigit():
                    months.append(s if "月" in s else f"{s}月")
                elif "単位" in s:
                    continue

    for row_idx in range(len(df)):
        for col_idx in range(min(3, df.shape[1])):
            cell = str(df.iloc[row_idx, col_idx]) if pd.notna(df.iloc[row_idx, col_idx]) else ""
            for key, label in targets.items():
                if key in cell:
                    vals = []
                    start_col = 4 if col_idx <= 1 else col_idx + 2
                    for c in range(start_col, start_col + len(months)):
                        if c < df.shape[1]:
                            v = df.iloc[row_idx, c]
                            vals.append(float(v) if pd.notna(v) and isinstance(v, (int, float)) else 0)
                    results[label] = vals
                    break

    return months, results


def extract_cf_detail(df):
    """CF推移シートから営業/投資/財務CFの内訳データを抽出する。"""
    months = []
    detail = {}
    signs = {}

    header_row = None
    for i in range(min(10, len(df))):
        for j in range(df.shape[1]):
            val = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
            if "年度" in val:
                header_row = i
                break
        if header_row is not None:
            break
    if header_row is None:
        return months, detail, signs

    for c in range(3, df.shape[1]):
        val = df.iloc[header_row, c]
        if pd.notna(val):
            s = str(val).replace(".0", "")
            if "月" in s or s.isdigit():
                months.append(s if "月" in s else f"{s}月")
            elif "単位" in s:
                continue

    sub_items = [
        ("税引前利益", "税引前利益"), ("決算関係納税", "納税"),
        ("減価償却費", "減価償却"), ("売掛債権", "売掛金"),
        ("棚卸資産", "棚卸資産"), ("その他流動資産", "その他流動資産"),
        ("買掛未払債務", "買掛金"), ("その他流動負債", "その他流動負債"),
        ("固定資産への投資", "設備投資"), ("投資勘定への支出", "投資勘定"),
        ("長期借入金返済", "借入返済"), ("長期借入金調達", "借入調達"),
        ("短期借入金", "短期借入"),
    ]

    for row_idx in range(header_row + 1, min(header_row + 25, len(df))):
        for col_idx in range(min(4, df.shape[1])):
            cell = str(df.iloc[row_idx, col_idx]).strip() if pd.notna(df.iloc[row_idx, col_idx]) else ""
            for key, label in sub_items:
                if cell == key:
                    sign_val = ""
                    if 3 < df.shape[1] and pd.notna(df.iloc[row_idx, 3]):
                        sign_val = str(df.iloc[row_idx, 3]).strip()
                    signs[label] = "+" if "＋" in sign_val else "-"
                    vals = []
                    for c in range(4, 4 + len(months)):
                        if c < df.shape[1]:
                            v = df.iloc[row_idx, c]
                            vals.append(float(v) if pd.notna(v) and isinstance(v, (int, float)) else 0)
                    detail[label] = vals
                    break

    return months, detail, signs


# ============================================================
# 暦月変換ヘルパー
# ============================================================

def fiscal_months_to_calendar(month_labels, fy_start, selected_month):
    """決算期の月ラベルを暦年/月の文字列に変換する。"""
    cal = []
    sel_year = int(selected_month[:4])
    sel_month = int(selected_month[4:6])
    fy_year = sel_year if sel_month >= fy_start else sel_year - 1
    for m_label in month_labels:
        m_num = int(m_label.replace("月", ""))
        year = fy_year if m_num >= fy_start else fy_year + 1
        cal.append(f"{year}/{m_num:02d}")
    return cal


def align_to_calendar(all_data, calendar_months, key):
    """各社データを暦月に揃えて合算する。"""
    totals = []
    for cal_m in calendar_months:
        total = 0
        for cid, info in all_data.items():
            if cal_m in info["months"] and key in info["data"]:
                idx = info["months"].index(cal_m)
                vals = info["data"][key]
                if idx < len(vals):
                    total += vals[idx]
        totals.append(total)
    return totals


def align_to_calendar_by_company(all_data, calendar_months, key):
    """各社データを暦月に揃えて、会社別の辞書で返す。"""
    result = {}
    for cid, info in all_data.items():
        vals = []
        for cal_m in calendar_months:
            v = 0
            if cal_m in info["months"] and key in info["data"]:
                idx = info["months"].index(cal_m)
                data = info["data"][key]
                if idx < len(data):
                    v = data[idx]
            vals.append(v)
        result[cid] = vals
    return result


# ============================================================
# KPIカード
# ============================================================

def kpi_card(label, value, unit="千円", delta=None, color="#333"):
    if isinstance(value, (int, float)):
        formatted = f"{value:,.0f}"
    else:
        formatted = str(value)
    delta_html = ""
    if delta is not None and isinstance(delta, (int, float)):
        d_color = "#27ae60" if delta >= 0 else "#c0392b"
        d_sign = "+" if delta >= 0 else ""
        delta_html = f"<div style='font-size:0.85rem; color:{d_color};'>{d_sign}{delta:,.0f} {unit}</div>"
    st.markdown(f"""
    <div style="background: white; border-radius: 10px; padding: 1rem; border-left: 4px solid {color};
                box-shadow: 0 1px 3px rgba(0,0,0,0.1); text-align:center;">
        <div style="font-size:0.8rem; color:#888;">{label}</div>
        <div style="font-size:1.5rem; font-weight:700; color:{color};">{formatted}</div>
        <div style="font-size:0.75rem; color:#aaa;">{unit}</div>
        {delta_html}
    </div>""", unsafe_allow_html=True)


# ============================================================
# 各社PL/CFグラフ描画
# ============================================================

def _trim_zero_tail(months, data_dict):
    """末尾の全項目0の月を除外して実績データだけにする。"""
    if not months:
        return months, data_dict
    last_active = 0
    for i in range(len(months)):
        for vals in data_dict.values():
            if i < len(vals) and vals[i] != 0:
                last_active = i
    trimmed_months = months[:last_active + 1]
    trimmed = {k: v[:last_active + 1] for k, v in data_dict.items()}
    return trimmed_months, trimmed


def _settlement_month_label(fy_start):
    """決算月のラベルを返す（例: fy_start=7 → '6月'）。"""
    m = fy_start - 1 if fy_start > 1 else 12
    return f"{m}月"


def _drop_settlement_month(months, data_dict, fy_start):
    """決算月が末尾にあれば除外する（決算月のデータは未確定のため）。"""
    if not months:
        return months, data_dict
    label = _settlement_month_label(fy_start)
    if months[-1] == label:
        trimmed_months = months[:-1]
        trimmed = {k: v[:-1] for k, v in data_dict.items()}
        return trimmed_months, trimmed
    return months, data_dict


CHART_LAYOUT = dict(
    plot_bgcolor="white",
    legend=dict(orientation="h", y=-0.18, x=0.5, xanchor="center",
                font=dict(size=12), bgcolor="rgba(255,255,255,0.8)"),
    margin=dict(l=60, r=60, t=30, b=60),
    hoverlabel=dict(font_size=13),
)

Y_AXIS_FORMAT = dict(
    gridcolor="#eee", gridwidth=1, zeroline=True, zerolinecolor="#ccc",
    tickformat=",", separatethousands=True,
)


def render_company_pl(fpath, cid, color, company, fy_start=7):
    """各社のPL推移 + サマリーテーブルを描画する。"""
    try:
        trend_df = read_pl_trend(str(fpath), cid)
        months_t, trend_data = extract_pl_trend_data(trend_df)
        months_d, pl_detail = extract_pl_detail(trend_df)
    except Exception as e:
        st.error(f"PL推移の読み込みエラー: {e}")
        return

    if not months_t or not trend_data:
        st.info("PL推移データを読み取れませんでした")
        return

    months_t, trend_data = _trim_zero_tail(months_t, trend_data)
    months_t, trend_data = _drop_settlement_month(months_t, trend_data, fy_start)
    if months_d and pl_detail:
        months_d, pl_detail = _trim_zero_tail(months_d, pl_detail)
        months_d, pl_detail = _drop_settlement_month(months_d, pl_detail, fy_start)
    n = len(months_t)

    op_vals = pl_detail.get("営業利益", trend_data.get("営業利益", []))[:n]
    ord_vals = pl_detail.get("経常利益", [])[:n]

    # ── KPIカード ──
    kpi_cols = st.columns(4)
    op_total = sum(op_vals) if op_vals else 0
    ord_total = sum(ord_vals) if ord_vals else 0
    rev_total = sum(trend_data.get("売上高", [])[:n])
    sga_total = sum(trend_data.get("販管費", [])[:n])
    with kpi_cols[0]:
        kpi_card("売上高累計", rev_total, color=color)
    with kpi_cols[1]:
        kpi_card("販管費累計", sga_total, color="#e74c3c" if sga_total > rev_total * 0.9 else "#888")
    with kpi_cols[2]:
        kpi_card("営業利益累計", op_total, color="#27ae60" if op_total >= 0 else "#c0392b")
    with kpi_cols[3]:
        kpi_card("経常利益累計", ord_total, color="#27ae60" if ord_total >= 0 else "#c0392b")

    # ── 営業利益・経常利益の月次推移グラフ ──
    st.markdown("#### 📊 月次 営業利益・経常利益")
    fig_profit = go.Figure()
    hover_yen = "<b>%{x}</b><br>%{fullData.name}: %{y:,.0f} 千円<extra></extra>"

    if op_vals:
        fig_profit.add_trace(go.Bar(
            x=months_t[:len(op_vals)], y=op_vals,
            name="営業利益",
            marker_color=["#27ae60" if v >= 0 else "#e74c3c" for v in op_vals],
            text=[f"{v:,.0f}" for v in op_vals],
            textposition="outside", textfont=dict(size=11),
            hovertemplate=hover_yen,
        ))
    if ord_vals:
        fig_profit.add_trace(go.Scatter(
            x=months_t[:len(ord_vals)], y=ord_vals,
            name="経常利益", mode="lines+markers+text",
            line=dict(color="#8e44ad", width=3), marker=dict(size=9),
            text=[f"{v:,.0f}" for v in ord_vals],
            textposition="top center", textfont=dict(size=10, color="#8e44ad"),
            hovertemplate=hover_yen,
        ))
    fig_profit.add_hline(y=0, line_dash="dash", line_color="#aaa", line_width=1)
    fig_profit.update_layout(height=400, yaxis=dict(title="千円", **Y_AXIS_FORMAT), **CHART_LAYOUT)
    st.plotly_chart(fig_profit, use_container_width=True)

    # ── なぜその金額か：月別PL構造テーブル ──
    st.markdown("#### 📋 月別PL構造（なぜこの利益か）")

    rev = pl_detail.get("売上高", trend_data.get("売上高", []))[:n]
    cogs = pl_detail.get("売上原価", [0] * n)[:n]
    gross = pl_detail.get("粗利", trend_data.get("粗利", []))[:n]
    sga = pl_detail.get("販管費", trend_data.get("販管費", []))[:n]
    op_ext_income = pl_detail.get("営業外収益", [0] * n)[:n]
    op_ext_expense = pl_detail.get("営業外費用", [0] * n)[:n]

    pl_rows = []
    for label, vals in [
        ("売上高", rev), ("売上原価", cogs), ("粗利", gross),
        ("販管費", sga), ("営業利益", op_vals),
        ("営業外収益", op_ext_income), ("営業外費用", op_ext_expense),
        ("経常利益", ord_vals),
    ]:
        if vals:
            row = {"項目": label}
            for idx, m in enumerate(months_t):
                if idx < len(vals):
                    row[m] = f"{vals[idx]:,.0f}"
            active = [v for v in vals if v != 0]
            row["累計"] = f"{sum(vals):,.0f}"
            pl_rows.append(row)

    if pl_rows:
        df_pl = pd.DataFrame(pl_rows)
        st.dataframe(df_pl, use_container_width=True, hide_index=True)

    # ── 販管費内訳テーブル ──
    sga_keys = [(k, k.replace("販管費_", "")) for k in sorted(pl_detail.keys()) if k.startswith("販管費_")]
    if sga_keys:
        with st.expander("📂 販管費の内訳（月別）"):
            sga_rows = []
            for full_key, short_name in sga_keys:
                vals = pl_detail[full_key][:n]
                if any(v != 0 for v in vals):
                    row = {"費目": short_name}
                    for idx, m in enumerate(months_t):
                        if idx < len(vals):
                            row[m] = f"{vals[idx]:,.0f}"
                    row["累計"] = f"{sum(vals):,.0f}"
                    row["月平均"] = f"{sum(vals)/max(len([v for v in vals if v != 0]), 1):,.0f}"
                    sga_rows.append(row)
            if sga_rows:
                sga_rows.sort(key=lambda r: -abs(float(r["累計"].replace(",", ""))))
                st.dataframe(pd.DataFrame(sga_rows), use_container_width=True, hide_index=True)

    # ── ウォーターフォール：直近月のPL構造 ──
    if rev and len(rev) >= 1:
        last_idx = len(months_t) - 1
        last_m = months_t[last_idx]
        st.markdown(f"#### 🔍 {last_m} PL構造（ウォーターフォール）")

        last_rev = rev[last_idx] if last_idx < len(rev) else 0
        last_cogs = cogs[last_idx] if last_idx < len(cogs) else 0
        last_sga = sga[last_idx] if last_idx < len(sga) else 0
        last_gross = last_rev - last_cogs
        last_op = last_gross - last_sga

        wf_labels = ["売上高", "売上原価", "粗利", "販管費", "営業利益"]
        wf_values = [last_rev, -last_cogs, 0, -last_sga, 0]
        wf_text = [f"{last_rev:,.0f}", f"{-last_cogs:,.0f}", f"{last_gross:,.0f}",
                    f"{-last_sga:,.0f}", f"{last_op:,.0f}"]
        wf_measures = ["absolute", "relative", "total", "relative", "total"]

        if ord_vals and last_idx < len(ord_vals):
            ext_inc = op_ext_income[last_idx] if last_idx < len(op_ext_income) else 0
            ext_exp = op_ext_expense[last_idx] if last_idx < len(op_ext_expense) else 0
            last_ord = last_op + ext_inc - ext_exp
            wf_labels += ["営業外収益", "営業外費用", "経常利益"]
            wf_values += [ext_inc, -ext_exp, 0]
            wf_text += [f"{ext_inc:,.0f}", f"{-ext_exp:,.0f}", f"{last_ord:,.0f}"]
            wf_measures += ["relative", "relative", "total"]

        fig_wf = go.Figure(go.Waterfall(
            x=wf_labels, y=wf_values, measure=wf_measures,
            connector=dict(line=dict(color="#ccc")),
            increasing=dict(marker_color="#27ae60"),
            decreasing=dict(marker_color="#e74c3c"),
            totals=dict(marker_color="#3498db"),
            text=wf_text, textposition="outside",
            textfont=dict(size=11),
        ))
        fig_wf.update_layout(height=400, yaxis=dict(title="千円", **Y_AXIS_FORMAT), **CHART_LAYOUT)
        st.plotly_chart(fig_wf, use_container_width=True)

    # ── 売上・粗利・販管費の推移グラフ ──
    with st.expander("📈 売上・粗利・販管費の推移"):
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        hover_pct = "<b>%{x}</b><br>%{fullData.name}: %{y:.1f}%<extra></extra>"

        if "売上高" in trend_data:
            fig.add_trace(go.Bar(
                x=months_t, y=trend_data["売上高"][:n],
                name="売上高", marker_color=color, opacity=0.75,
                text=[f"{v:,.0f}" for v in trend_data["売上高"][:n]],
                textposition="outside", textfont=dict(size=10),
                hovertemplate=hover_yen,
            ), secondary_y=False)
        if "粗利" in trend_data:
            fig.add_trace(go.Scatter(
                x=months_t, y=trend_data["粗利"][:n],
                name="粗利", mode="lines+markers",
                line=dict(color="#f39c12", width=3), marker=dict(size=8),
                hovertemplate=hover_yen,
            ), secondary_y=False)
        if "販管費" in trend_data:
            fig.add_trace(go.Scatter(
                x=months_t, y=trend_data["販管費"][:n],
                name="販管費", mode="lines+markers",
                line=dict(color="#e74c3c", width=3, dash="dot"), marker=dict(size=8),
                hovertemplate=hover_yen,
            ), secondary_y=False)
        if "売上高" in trend_data and "粗利" in trend_data:
            gross_rates = []
            for i in range(min(len(trend_data["売上高"]), len(trend_data["粗利"]), n)):
                r = (trend_data["粗利"][i] / trend_data["売上高"][i] * 100) if trend_data["売上高"][i] else 0
                gross_rates.append(r)
            fig.add_trace(go.Scatter(
                x=months_t[:len(gross_rates)], y=gross_rates,
                name="粗利率", mode="lines+markers",
                line=dict(color="#9b59b6", width=2, dash="dashdot"), marker=dict(size=7),
                hovertemplate=hover_pct,
            ), secondary_y=True)
        fig.update_layout(height=480, barmode="group", **CHART_LAYOUT)
        fig.update_yaxes(title_text="千円", secondary_y=False, **Y_AXIS_FORMAT)
        fig.update_yaxes(title_text="粗利率 (%)", secondary_y=True,
                         gridcolor="rgba(0,0,0,0)", ticksuffix="%")
        st.plotly_chart(fig, use_container_width=True)

    # セグメント別（ある場合）
    segments = company.get("segments", [])
    if segments:
        with st.expander("📊 セグメント別売上"):
            try:
                months_s, seg_data = extract_segments(trend_df, segments)
                if months_s and seg_data:
                    months_s, seg_data = _trim_zero_tail(months_s, seg_data)
                    months_s, seg_data = _drop_settlement_month(months_s, seg_data, fy_start)
                    fig_seg = go.Figure()
                    seg_colors = px.colors.qualitative.Set2
                    for j, (seg_name, seg_vals) in enumerate(seg_data.items()):
                        fig_seg.add_trace(go.Bar(
                            x=months_s[:len(seg_vals)], y=seg_vals, name=seg_name,
                            marker_color=seg_colors[j % len(seg_colors)],
                            hovertemplate="<b>%{x}</b><br>%{fullData.name}: %{y:,.0f} 千円<extra></extra>",
                        ))
                    fig_seg.update_layout(
                        barmode="stack", height=400,
                        yaxis=dict(title="千円", **Y_AXIS_FORMAT),
                        **CHART_LAYOUT,
                    )
                    st.plotly_chart(fig_seg, use_container_width=True)
            except Exception:
                pass


def _cf_bar_colors(vals):
    """プラスなら緑、マイナスなら赤の色リストを返す。"""
    return ["#27ae60" if v >= 0 else "#e74c3c" for v in vals]


def _cf_impact(detail, signs, item, idx):
    """CFへの寄与額（＋項目はそのまま、－項目は符号反転）。"""
    if item not in detail or idx >= len(detail[item]):
        return 0
    val = detail[item][idx]
    return val if signs.get(item) == "+" else -val


def _generate_cf_diagnosis(months, cf_data, detail=None, signs=None):
    """CFデータから原因分析テキストを生成する。"""
    if "月次CF" not in cf_data:
        return ""
    detail = detail or {}
    signs = signs or {}
    monthly = cf_data["月次CF"]
    n = min(len(months), len(monthly))
    active = [(i, months[i]) for i in range(n) if monthly[i] != 0]
    if not active:
        return ""

    total = sum(monthly[i] for i, _ in active)
    ops_vals = cf_data.get("営業CF", [0] * n)
    inv_vals = cf_data.get("投資CF", [0] * n)
    fin_vals = cf_data.get("財務CF", [0] * n)
    ops = sum(ops_vals[i] for i, _ in active)
    inv = sum(inv_vals[i] for i, _ in active)
    fin = sum(fin_vals[i] for i, _ in active)

    period = f"{active[0][1]}〜{active[-1][1]}" if len(active) > 1 else active[0][1]
    lines = []

    if total > 0:
        lines.append(f"✅ **{period}の累計: {total:+,.0f}千円（現金増加）**")
    elif total == 0:
        lines.append(f"➖ **{period}の累計: ±0千円**")
    else:
        lines.append(f"⚠️ **{period}の累計: {total:+,.0f}千円（現金減少）**")

    lines.append("")
    lines.append("**3つの活動の内訳:**")

    ops_items = ["税引前利益", "減価償却", "売掛金", "買掛金",
                 "棚卸資産", "その他流動資産", "その他流動負債", "納税"]
    inv_items = ["設備投資", "投資勘定"]
    fin_items = ["借入返済", "借入調達", "短期借入"]

    for name, val, items_list in [
        ("営業活動（本業の稼ぎ）", ops, ops_items),
        ("投資活動（設備・投資）", inv, inv_items),
        ("財務活動（借入・返済）", fin, fin_items),
    ]:
        emoji = "🟢" if val >= 0 else "🔴"
        lines.append(f"- {emoji} {name}: **{val:+,.0f}千円**")
        if detail and abs(val) > 100:
            contribs = []
            for item in items_list:
                if item in detail:
                    c = sum(_cf_impact(detail, signs, item, i) for i, _ in active)
                    if abs(c) > 50:
                        contribs.append((item, c))
            contribs.sort(key=lambda x: abs(x[1]), reverse=True)
            for item_name, c in contribs[:3]:
                lines.append(f"  - {item_name}: {c:+,.0f}千円")

    positive_months = sum(1 for i, _ in active if monthly[i] > 0)
    negative_months = sum(1 for i, _ in active if monthly[i] < 0)
    lines.append("")
    lines.append(f"プラス月: {positive_months}ヶ月 / マイナス月: {negative_months}ヶ月")

    worst_i, worst_m = min(active, key=lambda x: monthly[x[0]])
    best_i, best_m = max(active, key=lambda x: monthly[x[0]])

    if abs(monthly[worst_i]) > 3000:
        lines.append(f"📍 **最大減少月: {worst_m}（{monthly[worst_i]:+,.0f}千円）**")
        parts = []
        for key in ["営業CF", "投資CF", "財務CF"]:
            if key in cf_data and worst_i < len(cf_data[key]):
                v = cf_data[key][worst_i]
                if abs(v) > 100:
                    parts.append(f"{key} {v:+,.0f}")
        if parts:
            lines.append(f"  → {' / '.join(parts)}")
    if monthly[best_i] > 3000:
        lines.append(f"📍 **最大増加月: {best_m}（{monthly[best_i]:+,.0f}千円）**")

    return "\n".join(lines)


def _make_cf_waterfall(cf_data):
    """期間累計のCFウォーターフォール図を作成する。"""
    ops = sum(v for v in cf_data.get("営業CF", []) if v != 0)
    inv = sum(v for v in cf_data.get("投資CF", []) if v != 0)
    fin = sum(v for v in cf_data.get("財務CF", []) if v != 0)
    total = ops + inv + fin

    fig = go.Figure(go.Waterfall(
        orientation="v",
        measure=["relative", "relative", "relative", "total"],
        x=["営業CF<br>(本業)", "投資CF<br>(設備投資)", "財務CF<br>(借入返済)", "期間合計"],
        y=[ops, inv, fin, total],
        text=[f"{ops:+,.0f}", f"{inv:+,.0f}", f"{fin:+,.0f}", f"{total:+,.0f}"],
        textposition="outside",
        textfont=dict(size=14, weight="bold"),
        increasing=dict(marker=dict(color="#27ae60")),
        decreasing=dict(marker=dict(color="#e74c3c")),
        totals=dict(marker=dict(color="#2980b9")),
        connector=dict(line=dict(color="#999", width=1, dash="dot")),
        hovertemplate="<b>%{x}</b><br>%{y:+,.0f} 千円<extra></extra>",
    ))
    fig.update_layout(
        height=380, showlegend=False,
        yaxis=dict(title="千円", **Y_AXIS_FORMAT),
        **CHART_LAYOUT,
    )
    return fig


def render_company_cf(fpath, cid, color, fy_start=7):
    """各社のCF推移を原因分析付きで描画する。"""
    try:
        cf_df = read_cf(str(fpath), cid)
        cf_months, cf_data = extract_cf_data(cf_df)
        _, cf_detail, cf_signs = extract_cf_detail(cf_df)
    except Exception as e:
        st.error(f"CFデータの読み込みエラー: {e}")
        return

    if not cf_months or not cf_data:
        st.info("CFデータを読み取れませんでした")
        return

    cf_months, cf_data = _trim_zero_tail(cf_months, cf_data)
    cf_months, cf_data = _drop_settlement_month(cf_months, cf_data, fy_start)
    n = len(cf_months)
    cf_detail = {k: v[:n] for k, v in cf_detail.items()}

    # ── KPIカード ──
    kpi_cols = st.columns(4)
    for i, (key, label) in enumerate([
        ("営業CF", "営業CF累計"), ("投資CF", "投資CF累計"),
        ("財務CF", "財務CF累計"), ("月次CF", "月次CF累計"),
    ]):
        if key in cf_data:
            total = sum(v for v in cf_data[key] if v != 0)
            with kpi_cols[i]:
                c = "#27ae60" if total >= 0 else "#c0392b"
                kpi_card(label, total, "千円", color=c)

    # ── 診断テキスト ──
    st.markdown("---")
    st.markdown("##### 📋 CF診断：なぜ現金が増えた/減ったか")
    diagnosis = _generate_cf_diagnosis(cf_months, cf_data, cf_detail, cf_signs)
    if diagnosis:
        st.markdown(diagnosis)

    # ── ウォーターフォール ──
    st.markdown("---")
    st.markdown("##### 🏗️ 期間累計ウォーターフォール")
    st.caption("営業→投資→財務の順に、現金がどう変化したか")
    fig_wf = _make_cf_waterfall(cf_data)
    st.plotly_chart(fig_wf, use_container_width=True)

    # ── 月別推移（積み上げ棒 + 折れ線） ──
    st.markdown("##### 📊 月別CF推移")
    st.caption("棒の色: 営業(緑)・投資(黄)・財務(紫) を積み上げ。黒線が月次CF合計")
    fig_m = go.Figure()
    for key, lbl, clr in [("営業CF", "営業CF", "#27ae60"),
                          ("投資CF", "投資CF", "#f39c12"),
                          ("財務CF", "財務CF", "#8e44ad")]:
        if key in cf_data:
            fig_m.add_trace(go.Bar(
                x=cf_months, y=cf_data[key][:n], name=lbl, marker_color=clr,
                hovertemplate=f"<b>%{{x}}</b><br>{lbl}: %{{y:+,.0f}} 千円<extra></extra>",
            ))
    if "月次CF" in cf_data:
        fig_m.add_trace(go.Scatter(
            x=cf_months, y=cf_data["月次CF"][:n], name="月次CF合計",
            mode="lines+markers+text", line=dict(color="#2c3e50", width=3),
            marker=dict(size=8), text=[f"{v:+,.0f}" for v in cf_data["月次CF"][:n]],
            textposition="top center", textfont=dict(size=9),
            hovertemplate="<b>%{x}</b><br>月次CF: %{y:+,.0f} 千円<extra></extra>",
        ))
    fig_m.update_layout(barmode="relative", height=450,
                        yaxis=dict(title="千円", **Y_AXIS_FORMAT), **CHART_LAYOUT)
    fig_m.add_hline(y=0, line_color="#999", line_width=1)
    st.plotly_chart(fig_m, use_container_width=True)

    # ── 月別CFテーブル ──
    with st.expander("📝 月別CF内訳テーブル"):
        rows = []
        for i, m in enumerate(cf_months):
            row = {"月": m}
            for key in ["営業CF", "投資CF", "財務CF", "月次CF"]:
                if key in cf_data and i < len(cf_data[key]):
                    row[key] = f"{cf_data[key][i]:+,.0f}"
            rows.append(row)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # ── 営業CF内訳 ──
    if cf_detail:
        with st.expander("📖 営業CFの詳細内訳（何が本業の稼ぎを動かしているか）"):
            ops_items = ["税引前利益", "減価償却", "売掛金", "買掛金",
                         "棚卸資産", "その他流動資産", "その他流動負債", "納税"]
            d_rows = []
            for i, m in enumerate(cf_months):
                row = {"月": m}
                for item in ops_items:
                    if item in cf_detail:
                        impact = _cf_impact(cf_detail, cf_signs, item, i)
                        row[item] = f"{impact:+,.0f}"
                d_rows.append(row)
            if d_rows:
                st.dataframe(pd.DataFrame(d_rows), use_container_width=True, hide_index=True)


# ============================================================
# メインアプリ
# ============================================================

def main():
    st.set_page_config(page_title="esグループ月次会議", page_icon="📊", layout="wide")

    st.markdown("""
    <style>
    .block-container { padding-top: 1.5rem; }
    [data-testid="stMetricValue"] { font-size: 1.2rem; }
    </style>
    """, unsafe_allow_html=True)

    config = load_config()
    companies = config["companies"]
    available_months = find_available_months()

    st.title("📊 esグループ月次会議 ダッシュボード")

    if not available_months:
        st.warning("月次データが見つかりません。`data/monthly/` にExcelファイルを配置してください。")
        return

    selected_month = st.selectbox(
        "対象月", available_months,
        format_func=lambda m: f"{m[:4]}年{int(m[4:])}月"
    )

    # ====================
    # 1. 3社サマリー
    # ====================
    st.divider()
    st.subheader("📋 3社サマリー")

    cols = st.columns([1, 1, 1, 1])
    company_data = {}
    sum_rev = 0
    sum_gross = 0
    sum_ord = 0
    sum_cash = 0

    for i, company in enumerate(companies):
        cid = company["id"]
        color = COMPANY_COLORS.get(cid, "#333")
        fpath = find_file(company, selected_month)

        with cols[i]:
            st.markdown(f"### {company['name']}")
            st.caption(f"{company['legal_entity']}（{company['business_type']}）")

            if not fpath:
                st.warning("データなし")
                continue

            pl_df = read_pl(str(fpath), cid)
            bs_df = read_bs(str(fpath), cid)

            fy_s = company.get("fiscal_year_start", 7)
            months_pl, pl_data = extract_monthly_pl(pl_df)
            months_pl, pl_data = _drop_settlement_month(months_pl, pl_data, fy_s)

            revenue = pl_data.get("売上高", [])
            gross = pl_data.get("粗利", [])
            total_rev = sum(revenue)
            total_gross = sum(gross)
            gross_rate = (total_gross / total_rev * 100) if total_rev else 0

            ord_total_sum = 0
            try:
                trend_df_s = read_pl_trend(str(fpath), cid)
                _, pl_det_s = extract_pl_detail(trend_df_s)
                if pl_det_s:
                    _, pl_det_s = _trim_zero_tail(months_pl, pl_det_s)
                    _, pl_det_s = _drop_settlement_month(months_pl, pl_det_s, fy_s)
                    ord_vals_s = pl_det_s.get("経常利益", [])
                    ord_total_sum = sum(ord_vals_s)
            except Exception:
                pass

            sum_rev += total_rev / 1000
            sum_gross += total_gross / 1000
            sum_ord += ord_total_sum

            kpi_card("売上高累計", total_rev / 1000, "千円", color=color)
            st.write("")
            kpi_card("粗利累計", total_gross / 1000, "千円", color=color)
            st.write("")
            kpi_card("経常利益累計", ord_total_sum, "千円", color="#27ae60" if ord_total_sum >= 0 else "#c0392b")
            st.write("")
            kpi_card("粗利率", f"{gross_rate:.1f}%", "", color=color)

            bs_months_kpi, bs_data = extract_bs_trend(bs_df, fy_s)
            bs_months_kpi, bs_data = _drop_settlement_month(bs_months_kpi, bs_data, fy_s)
            cash_vals = bs_data.get("現金及び預金合計", [])
            if cash_vals:
                latest_cash = [v for v in cash_vals if v != 0]
                if latest_cash:
                    kpi_card("現金預金", latest_cash[-1], "千円", color=color)
                    sum_cash += latest_cash[-1]

            company_data[cid] = {
                "file": fpath,
                "months_pl": months_pl,
                "pl_data": pl_data,
                "bs_data": bs_data,
                "color": color,
                "company": company,
            }

    with cols[3]:
        st.markdown("### 3社合計")
        st.caption("連結ベース")
        kpi_card("売上高累計", sum_rev, "千円", color="#333")
        st.write("")
        kpi_card("粗利累計", sum_gross, "千円", color="#333")
        st.write("")
        kpi_card("経常利益累計", sum_ord, "千円", color="#27ae60" if sum_ord >= 0 else "#c0392b")
        st.write("")
        gross_rate_all = (sum_gross / sum_rev * 100) if sum_rev else 0
        kpi_card("粗利率", f"{gross_rate_all:.1f}%", "", color="#333")
        if sum_cash:
            kpi_card("現金預金合計", sum_cash, "千円", color="#333")

    # ====================
    # 2. 連結ビュー（PL + CF）
    # ====================
    st.divider()
    st.subheader("🏢 3社連結")

    # --- 暦月でPL/CFデータを揃える ---
    all_pl = {}
    all_cf = {}

    for company in companies:
        cid = company["id"]
        if cid not in company_data:
            continue
        fpath = company_data[cid]["file"]
        fy_start = company.get("fiscal_year_start", 7)

        # PL
        try:
            trend_df = read_pl_trend(str(fpath), cid)
            months_t, trend_data = extract_pl_trend_data(trend_df)
            months_t, trend_data = _trim_zero_tail(months_t, trend_data)
            months_t, trend_data = _drop_settlement_month(months_t, trend_data, fy_start)
            _, pl_det = extract_pl_detail(trend_df)
            if pl_det:
                _, pl_det = _trim_zero_tail(months_t, pl_det)
                _, pl_det = _drop_settlement_month(months_t, pl_det, fy_start)
                for dk in ["営業利益", "経常利益", "営業外収益", "営業外費用"]:
                    if dk in pl_det:
                        trend_data[dk] = pl_det[dk][:len(months_t)]
            cal = fiscal_months_to_calendar(months_t, fy_start, selected_month)
            all_pl[cid] = {"months": cal, "data": trend_data, "name": company["name"]}
        except Exception:
            pass

        # CF
        try:
            cf_df = read_cf(str(fpath), cid)
            cf_months, cf_data = extract_cf_data(cf_df)
            cf_months, cf_data = _trim_zero_tail(cf_months, cf_data)
            cf_months, cf_data = _drop_settlement_month(cf_months, cf_data, fy_start)
            cal = fiscal_months_to_calendar(cf_months, fy_start, selected_month)
            all_cf[cid] = {"months": cal, "data": cf_data, "name": company["name"]}
        except Exception:
            pass

    # 全社の暦月を統合
    all_cal_set = set()
    for v in list(all_pl.values()) + list(all_cf.values()):
        all_cal_set.update(v["months"])
    calendar_months = sorted(all_cal_set)
    cal_labels = [f"{int(m.split('/')[1])}月" for m in calendar_months]

    # 実績ある月（PL売上ベース）をタブ共通で計算
    active_idx = []
    if all_pl:
        rev_total = align_to_calendar(all_pl, calendar_months, "売上高")
        active_idx = [i for i, v in enumerate(rev_total) if v != 0]

    tab_pl, tab_cf = st.tabs(["📈 連結PL", "💰 連結CF"])

    # --- 連結PL ---
    with tab_pl:
        if all_pl:
            sub_total, sub_company = st.tabs(["合算", "法人別比較"])

            with sub_total:
                fig_pl = make_subplots(specs=[[{"secondary_y": True}]])
                hover_yen = "<b>%{x}</b><br>%{fullData.name}: %{y:,.0f} 千円<extra></extra>"

                gross_total = align_to_calendar(all_pl, calendar_months, "粗利")
                sgna_total = align_to_calendar(all_pl, calendar_months, "販管費")
                op_total_cal = align_to_calendar(all_pl, calendar_months, "営業利益")
                ord_total_cal = align_to_calendar(all_pl, calendar_months, "経常利益")
                # fallback: 営業利益がdetailにない場合は粗利-販管費
                if all(v == 0 for v in op_total_cal):
                    op_total_cal = [g - s for g, s in zip(gross_total, sgna_total)]
                if active_idx:
                    a_labels = [cal_labels[i] for i in active_idx]
                    a_rev = [rev_total[i] for i in active_idx]
                    a_gross = [gross_total[i] for i in active_idx]
                    a_sgna = [sgna_total[i] for i in active_idx]
                    a_op = [op_total_cal[i] for i in active_idx]
                    a_ord = [ord_total_cal[i] for i in active_idx]
                else:
                    a_labels, a_rev, a_gross, a_sgna, a_op = cal_labels, rev_total, gross_total, sgna_total, op_total_cal
                    a_ord = ord_total_cal

                # 会社別の営業利益・経常利益を暦月にアライン
                op_by_co = align_to_calendar_by_company(all_pl, calendar_months, "営業利益")
                ord_by_co = align_to_calendar_by_company(all_pl, calendar_months, "経常利益")

                # active monthsに絞る
                if active_idx:
                    for cid in op_by_co:
                        op_by_co[cid] = [op_by_co[cid][i] for i in active_idx]
                        ord_by_co[cid] = [ord_by_co[cid][i] for i in active_idx]

                # 連結合計ラインは棒グラフの合計から直接計算（ずれ防止）
                n_months = len(a_labels)
                op_line = [sum(op_by_co[cid][j] for cid in op_by_co) for j in range(n_months)]
                ord_line = [sum(ord_by_co[cid][j] for cid in ord_by_co) for j in range(n_months)]

                # 累積値を計算
                op_cum = []
                ord_cum = []
                r_op = 0
                r_ord = 0
                for j in range(n_months):
                    r_op += op_line[j]
                    r_ord += ord_line[j]
                    op_cum.append(r_op)
                    ord_cum.append(r_ord)

                def _align_zero_axes(fig, by_co, cum_vals, n_m):
                    """左右Y軸のゼロ位置を揃える"""
                    pos_stack = [0.0] * n_m
                    neg_stack = [0.0] * n_m
                    for cid in by_co:
                        for j in range(n_m):
                            v = by_co[cid][j]
                            if v >= 0:
                                pos_stack[j] += v
                            else:
                                neg_stack[j] += v
                    y1_hi = max(pos_stack) * 1.2 if max(pos_stack) > 0 else 100
                    y1_lo = min(neg_stack) * 1.2 if min(neg_stack) < 0 else -100
                    y2_hi = max(cum_vals)
                    y2_lo = min(cum_vals)
                    y2_pad = (y2_hi - y2_lo) * 0.15 or 100
                    y2_hi += y2_pad
                    y2_lo -= y2_pad
                    zero_frac = -y1_lo / (y1_hi - y1_lo)
                    if zero_frac > 0 and zero_frac < 1:
                        if abs(y2_lo) >= abs(y2_hi):
                            y2_hi = -y2_lo * (1 - zero_frac) / zero_frac
                        else:
                            y2_lo = -y2_hi * zero_frac / (1 - zero_frac)
                    fig.update_yaxes(range=[y1_lo, y1_hi], secondary_y=False)
                    fig.update_yaxes(range=[y2_lo, y2_hi], secondary_y=True)

                # ── 営業利益：会社別積み上げ + 合計ライン + 累積ライン ──
                st.markdown("#### 📊 月次 営業利益（会社別内訳）")
                fig_op = make_subplots(specs=[[{"secondary_y": True}]])
                for cid in op_by_co:
                    vals = op_by_co[cid]
                    fig_op.add_trace(go.Bar(
                        x=a_labels, y=vals,
                        name=COMPANY_SHORT.get(cid, cid),
                        marker_color=COMPANY_COLORS.get(cid, "#888"),
                        hovertemplate="<b>%{x}</b><br>" + COMPANY_SHORT.get(cid, cid) + ": %{y:,.0f} 千円<extra></extra>",
                    ), secondary_y=False)
                fig_op.add_trace(go.Scatter(
                    x=a_labels, y=op_line, name="連結合計",
                    mode="lines+markers+text",
                    line=dict(color="#2c3e50", width=3), marker=dict(size=10, symbol="diamond", color="#2c3e50"),
                    text=[f"{v:,.0f}" for v in op_line],
                    textposition="top center", textfont=dict(size=11, color="#2c3e50"),
                    hovertemplate="<b>%{x}</b><br>連結合計: %{y:,.0f} 千円<extra></extra>",
                ), secondary_y=False)
                fig_op.add_trace(go.Scatter(
                    x=a_labels, y=op_cum, name="累積営業利益",
                    mode="lines+markers+text",
                    line=dict(color="#e67e22", width=2, dash="dot"),
                    marker=dict(size=7, symbol="circle", color="#e67e22"),
                    text=[f"{v:,.0f}" for v in op_cum],
                    textposition="top center", textfont=dict(size=9, color="#e67e22"),
                    hovertemplate="<b>%{x}</b><br>累積営業利益: %{y:,.0f} 千円<extra></extra>",
                ), secondary_y=True)
                fig_op.add_hline(y=0, line_dash="dash", line_color="#aaa", line_width=1)
                fig_op.update_layout(
                    barmode="relative", height=450,
                    **{k: v for k, v in CHART_LAYOUT.items() if k != "legend"},
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
                )
                fig_op.update_yaxes(title_text="千円（月次）", secondary_y=False, **Y_AXIS_FORMAT)
                fig_op.update_yaxes(title_text="千円（累積）", secondary_y=True, gridcolor="rgba(0,0,0,0)", tickformat=",")
                _align_zero_axes(fig_op, op_by_co, op_cum, n_months)
                st.plotly_chart(fig_op, use_container_width=True)

                # ── 経常利益：会社別積み上げ + 合計ライン + 累積ライン ──
                if any(v != 0 for v in ord_line):
                    st.markdown("#### 📊 月次 経常利益（会社別内訳）")
                    fig_ord = make_subplots(specs=[[{"secondary_y": True}]])
                    for cid in ord_by_co:
                        vals = ord_by_co[cid]
                        fig_ord.add_trace(go.Bar(
                            x=a_labels, y=vals,
                            name=COMPANY_SHORT.get(cid, cid),
                            marker_color=COMPANY_COLORS.get(cid, "#888"),
                            hovertemplate="<b>%{x}</b><br>" + COMPANY_SHORT.get(cid, cid) + ": %{y:,.0f} 千円<extra></extra>",
                        ), secondary_y=False)
                    fig_ord.add_trace(go.Scatter(
                        x=a_labels, y=ord_line, name="連結合計",
                        mode="lines+markers+text",
                        line=dict(color="#2c3e50", width=3), marker=dict(size=10, symbol="diamond", color="#2c3e50"),
                        text=[f"{v:,.0f}" for v in ord_line],
                        textposition="top center", textfont=dict(size=11, color="#2c3e50"),
                        hovertemplate="<b>%{x}</b><br>連結合計: %{y:,.0f} 千円<extra></extra>",
                    ), secondary_y=False)
                    fig_ord.add_trace(go.Scatter(
                        x=a_labels, y=ord_cum, name="累積経常利益",
                        mode="lines+markers+text",
                        line=dict(color="#8e44ad", width=2, dash="dot"),
                        marker=dict(size=7, symbol="diamond", color="#8e44ad"),
                        text=[f"{v:,.0f}" for v in ord_cum],
                        textposition="top center", textfont=dict(size=9, color="#8e44ad"),
                        hovertemplate="<b>%{x}</b><br>累積経常利益: %{y:,.0f} 千円<extra></extra>",
                    ), secondary_y=True)
                    fig_ord.add_hline(y=0, line_dash="dash", line_color="#aaa", line_width=1)
                    fig_ord.update_layout(
                        barmode="relative", height=450,
                        **{k: v for k, v in CHART_LAYOUT.items() if k != "legend"},
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
                    )
                    fig_ord.update_yaxes(title_text="千円（月次）", secondary_y=False, **Y_AXIS_FORMAT)
                    fig_ord.update_yaxes(title_text="千円（累積）", secondary_y=True, gridcolor="rgba(0,0,0,0)", tickformat=",")
                    _align_zero_axes(fig_ord, ord_by_co, ord_cum, n_months)
                    st.plotly_chart(fig_ord, use_container_width=True)

                # ── 月次キャッシュフロー：会社別積み上げ + 合計ライン + 累積ライン ──
                if all_cf:
                    cf_by_co = align_to_calendar_by_company(all_cf, calendar_months, "月次CF")
                    if active_idx:
                        for cid in cf_by_co:
                            cf_by_co[cid] = [cf_by_co[cid][i] for i in active_idx]
                    cf_line = [sum(cf_by_co[cid][j] for cid in cf_by_co) for j in range(n_months)]
                    cf_cum = []
                    r_cf = 0
                    for j in range(n_months):
                        r_cf += cf_line[j]
                        cf_cum.append(r_cf)

                    if any(v != 0 for v in cf_line):
                        st.markdown("#### 📊 月次キャッシュフロー（会社別内訳）")
                        fig_cf = make_subplots(specs=[[{"secondary_y": True}]])
                        for cid in cf_by_co:
                            vals = cf_by_co[cid]
                            fig_cf.add_trace(go.Bar(
                                x=a_labels, y=vals,
                                name=COMPANY_SHORT.get(cid, cid),
                                marker_color=COMPANY_COLORS.get(cid, "#888"),
                                hovertemplate="<b>%{x}</b><br>" + COMPANY_SHORT.get(cid, cid) + ": %{y:,.0f} 千円<extra></extra>",
                            ), secondary_y=False)
                        fig_cf.add_trace(go.Scatter(
                            x=a_labels, y=cf_line, name="連結合計",
                            mode="lines+markers+text",
                            line=dict(color="#2c3e50", width=3), marker=dict(size=10, symbol="diamond", color="#2c3e50"),
                            text=[f"{v:,.0f}" for v in cf_line],
                            textposition="top center", textfont=dict(size=11, color="#2c3e50"),
                            hovertemplate="<b>%{x}</b><br>連結合計: %{y:,.0f} 千円<extra></extra>",
                        ), secondary_y=False)
                        fig_cf.add_trace(go.Scatter(
                            x=a_labels, y=cf_cum, name="累積CF",
                            mode="lines+markers+text",
                            line=dict(color="#16a085", width=2, dash="dot"),
                            marker=dict(size=7, symbol="circle", color="#16a085"),
                            text=[f"{v:,.0f}" for v in cf_cum],
                            textposition="top center", textfont=dict(size=9, color="#16a085"),
                            hovertemplate="<b>%{x}</b><br>累積CF: %{y:,.0f} 千円<extra></extra>",
                        ), secondary_y=True)
                        fig_cf.add_hline(y=0, line_dash="dash", line_color="#aaa", line_width=1)
                        fig_cf.update_layout(
                            barmode="relative", height=450,
                            **{k: v for k, v in CHART_LAYOUT.items() if k != "legend"},
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
                        )
                        fig_cf.update_yaxes(title_text="千円（月次）", secondary_y=False, **Y_AXIS_FORMAT)
                        fig_cf.update_yaxes(title_text="千円（累積）", secondary_y=True, gridcolor="rgba(0,0,0,0)", tickformat=",")
                        _align_zero_axes(fig_cf, cf_by_co, cf_cum, n_months)
                        st.plotly_chart(fig_cf, use_container_width=True)

                # 売上・粗利・販管費の推移グラフ
                with st.expander("📈 売上・粗利・販管費の推移"):
                    fig_pl = make_subplots(specs=[[{"secondary_y": True}]])
                    fig_pl.add_trace(go.Bar(
                        x=a_labels, y=a_rev, name="売上高",
                        marker_color="#34495e", opacity=0.75,
                        text=[f"{v:,.0f}" for v in a_rev],
                        textposition="outside", textfont=dict(size=10),
                        hovertemplate=hover_yen,
                    ), secondary_y=False)
                    fig_pl.add_trace(go.Scatter(
                        x=a_labels, y=a_gross, name="粗利",
                        mode="lines+markers", line=dict(color="#f39c12", width=3),
                        marker=dict(size=8), hovertemplate=hover_yen,
                    ), secondary_y=False)
                    fig_pl.add_trace(go.Scatter(
                        x=a_labels, y=a_sgna, name="販管費",
                        mode="lines+markers", line=dict(color="#e74c3c", width=3, dash="dot"),
                        marker=dict(size=8), hovertemplate=hover_yen,
                    ), secondary_y=False)
                    rev_cum = []
                    running = 0
                    for v in a_rev:
                        running += v
                        rev_cum.append(running)
                    fig_pl.add_trace(go.Scatter(
                        x=a_labels, y=rev_cum, name="売上累計",
                        mode="lines+markers", line=dict(color="#95a5a6", width=2, dash="dash"),
                        marker=dict(size=6), hovertemplate=hover_yen,
                    ), secondary_y=True)
                    fig_pl.update_layout(barmode="group", height=500, **CHART_LAYOUT)
                    fig_pl.update_yaxes(title_text="千円（月次）", secondary_y=False, **Y_AXIS_FORMAT)
                    fig_pl.update_yaxes(title_text="千円（累計）", secondary_y=True,
                                        gridcolor="rgba(0,0,0,0)", tickformat=",")
                    st.plotly_chart(fig_pl, use_container_width=True)

                # 連結PLサマリーテーブル
                pl_rows = []
                for key, label in [("売上高", "売上高"), ("粗利", "粗利"), ("販管費", "販管費"),
                                    ("営業利益", "営業利益"), ("営業外収益", "営業外収益"),
                                    ("営業外費用", "営業外費用"), ("経常利益", "経常利益")]:
                    row = {"項目": label}
                    grand = 0
                    has_data = False
                    for cid, info in all_pl.items():
                        vals = info["data"].get(key, [])
                        if vals:
                            has_data = True
                            s = sum(vals)
                            row[info["name"]] = f"{s:,.0f}"
                            grand += s
                        else:
                            row[info["name"]] = "-"
                    if has_data:
                        row["連結合計"] = f"{grand:,.0f}"
                        pl_rows.append(row)
                if pl_rows:
                    st.dataframe(pd.DataFrame(pl_rows), use_container_width=True, hide_index=True)

            with sub_company:
                hover_yen = "<b>%{x}</b><br>%{fullData.name}: %{y:,.0f} 千円<extra></extra>"
                fig_rev = go.Figure()
                for cid_c, info in all_pl.items():
                    vals = []
                    for cal_m in calendar_months:
                        if cal_m in info["months"] and "売上高" in info["data"]:
                            idx = info["months"].index(cal_m)
                            d = info["data"]["売上高"]
                            vals.append(d[idx] if idx < len(d) else 0)
                        else:
                            vals.append(0)
                    a_vals = [vals[i] for i in active_idx] if active_idx else vals
                    fig_rev.add_trace(go.Bar(
                        x=a_labels if active_idx else cal_labels, y=a_vals, name=info["name"],
                        marker_color=COMPANY_COLORS.get(cid_c, "#888"),
                        hovertemplate=hover_yen,
                    ))
                fig_rev.add_trace(go.Scatter(
                    x=a_labels if active_idx else cal_labels, y=a_rev,
                    name="連結合計", mode="lines+markers",
                    line=dict(color="#2c3e50", width=3, dash="dash"),
                    marker=dict(size=8), hovertemplate=hover_yen,
                ))
                fig_rev.update_layout(
                    barmode="group", height=480,
                    yaxis=dict(title="千円", **Y_AXIS_FORMAT),
                    title=dict(text="売上高 法人別比較", font=dict(size=15)),
                    **CHART_LAYOUT,
                )
                st.plotly_chart(fig_rev, use_container_width=True)

                fig_op = go.Figure()
                for cid_c, info in all_pl.items():
                    vals = []
                    for cal_m in calendar_months:
                        if cal_m in info["months"]:
                            idx = info["months"].index(cal_m)
                            g = info["data"].get("粗利", [])
                            s = info["data"].get("販管費", [])
                            gv = g[idx] if idx < len(g) else 0
                            sv = s[idx] if idx < len(s) else 0
                            vals.append(gv - sv)
                        else:
                            vals.append(0)
                    a_vals = [vals[i] for i in active_idx] if active_idx else vals
                    fig_op.add_trace(go.Bar(
                        x=a_labels if active_idx else cal_labels, y=a_vals, name=info["name"],
                        marker_color=COMPANY_COLORS.get(cid_c, "#888"),
                        hovertemplate=hover_yen,
                    ))
                fig_op.add_trace(go.Scatter(
                    x=a_labels if active_idx else cal_labels, y=a_op,
                    name="連結合計", mode="lines+markers",
                    line=dict(color="#2c3e50", width=3, dash="dash"),
                    marker=dict(size=8), hovertemplate=hover_yen,
                ))
                fig_op.update_layout(
                    barmode="group", height=480,
                    yaxis=dict(title="千円", **Y_AXIS_FORMAT),
                    title=dict(text="営業利益 法人別比較", font=dict(size=15)),
                    **CHART_LAYOUT,
                )
                st.plotly_chart(fig_op, use_container_width=True)
        else:
            st.info("PLデータがありません")

    # --- 連結CF ---
    with tab_cf:
        if all_cf:
            ops_total_vals = align_to_calendar(all_cf, calendar_months, "営業CF")
            inv_total_vals = align_to_calendar(all_cf, calendar_months, "投資CF")
            fin_total_vals = align_to_calendar(all_cf, calendar_months, "財務CF")
            monthly_total_vals = align_to_calendar(all_cf, calendar_months, "月次CF")

            # PL売上ベースのactive_idxを使う（PLとCFで表示月を統一）
            cf_active_idx = active_idx if active_idx else [i for i, v in enumerate(monthly_total_vals) if v != 0]
            if not cf_active_idx:
                cf_active_idx = list(range(len(cal_labels)))

            cf_a_labels = [cal_labels[i] for i in cf_active_idx]
            a_ops = [ops_total_vals[i] for i in cf_active_idx]
            a_inv = [inv_total_vals[i] for i in cf_active_idx]
            a_fin = [fin_total_vals[i] for i in cf_active_idx]
            a_monthly = [monthly_total_vals[i] for i in cf_active_idx]

            agg_cf = {"営業CF": a_ops, "投資CF": a_inv, "財務CF": a_fin, "月次CF": a_monthly}

            cf_period = f"{cf_a_labels[0]}〜{cf_a_labels[-1]}" if cf_a_labels else ""
            sub_total_cf, sub_company_cf = st.tabs(["合算", "法人別比較"])

            with sub_total_cf:
                if cf_period:
                    st.caption(f"集計期間: {cf_period}（PL売上実績のある月）")
                # KPIカード
                kpi_cols = st.columns(4)
                for i, (key, label) in enumerate([
                    ("営業CF", "営業CF累計"), ("投資CF", "投資CF累計"),
                    ("財務CF", "財務CF累計"), ("月次CF", "月次CF累計"),
                ]):
                    total = sum(v for v in agg_cf[key] if v != 0)
                    with kpi_cols[i]:
                        c = "#27ae60" if total >= 0 else "#c0392b"
                        kpi_card(f"3社{label}", total, "千円", color=c)

                # 診断テキスト
                st.markdown("---")
                st.markdown("##### 📋 3社合算CF診断：なぜ現金が増えた/減ったか")
                diagnosis = _generate_cf_diagnosis(cf_a_labels, agg_cf)
                if diagnosis:
                    st.markdown(diagnosis)

                # ウォーターフォール
                st.markdown("---")
                st.markdown("##### 🏗️ 3社合算：期間累計ウォーターフォール")
                st.caption("営業→投資→財務の順に、現金がどう変化したか")
                fig_wf = _make_cf_waterfall(agg_cf)
                st.plotly_chart(fig_wf, use_container_width=True)

                # 月別推移（積み上げ棒）
                st.markdown("##### 📊 3社合算：月別CF推移")
                st.caption("棒の色: 営業(緑)・投資(黄)・財務(紫) を積み上げ。黒線が月次CF合計")
                fig_m = go.Figure()
                for key, lbl, clr in [("営業CF", "営業CF", "#27ae60"),
                                      ("投資CF", "投資CF", "#f39c12"),
                                      ("財務CF", "財務CF", "#8e44ad")]:
                    fig_m.add_trace(go.Bar(
                        x=cf_a_labels, y=agg_cf[key], name=lbl, marker_color=clr,
                        hovertemplate=f"<b>%{{x}}</b><br>{lbl}: %{{y:+,.0f}} 千円<extra></extra>",
                    ))
                fig_m.add_trace(go.Scatter(
                    x=cf_a_labels, y=a_monthly, name="月次CF合計",
                    mode="lines+markers+text", line=dict(color="#2c3e50", width=3),
                    marker=dict(size=8), text=[f"{v:+,.0f}" for v in a_monthly],
                    textposition="top center", textfont=dict(size=9),
                    hovertemplate="<b>%{x}</b><br>月次CF: %{y:+,.0f} 千円<extra></extra>",
                ))
                fig_m.update_layout(barmode="relative", height=450,
                                    yaxis=dict(title="千円", **Y_AXIS_FORMAT), **CHART_LAYOUT)
                fig_m.add_hline(y=0, line_color="#999", line_width=1)
                st.plotly_chart(fig_m, use_container_width=True)

                # サマリーテーブル（表示月に絞った集計）
                cf_by_co_all = {}
                for cf_type in ["営業CF", "投資CF", "財務CF", "月次CF"]:
                    cf_by_co_all[cf_type] = align_to_calendar_by_company(all_cf, calendar_months, cf_type)
                    for cid_f in cf_by_co_all[cf_type]:
                        cf_by_co_all[cf_type][cid_f] = [cf_by_co_all[cf_type][cid_f][i] for i in cf_active_idx]

                cf_rows = []
                for cf_type in ["営業CF", "投資CF", "財務CF", "月次CF"]:
                    row = {"区分": cf_type}
                    grand = 0
                    for cid_f in cf_by_co_all[cf_type]:
                        name = all_cf[cid_f]["name"]
                        s = sum(cf_by_co_all[cf_type][cid_f])
                        row[name] = f"{s:+,.0f}"
                        grand += s
                    row["3社合計"] = f"{grand:+,.0f}"
                    cf_rows.append(row)
                st.dataframe(pd.DataFrame(cf_rows), use_container_width=True, hide_index=True)

            with sub_company_cf:
                # 法人別CF比較テーブル + 横棒グラフ
                st.markdown("##### 🏗️ 法人別：期間累計キャッシュフロー比較")
                if cf_period:
                    st.caption(f"集計期間: {cf_period}")

                cf_types_display = ["営業CF", "投資CF", "財務CF", "月次CF"]
                co_summaries = {}
                for cid_c, cf_info in all_cf.items():
                    sums = {}
                    for cf_type in cf_types_display:
                        vals = cf_by_co_all.get(cf_type, {}).get(cid_c, [])
                        sums[cf_type] = sum(vals)
                    co_summaries[cid_c] = {"name": cf_info["name"], "sums": sums}

                cf_categories = ["営業CF（本業）", "投資CF（設備投資）", "財務CF（借入返済）"]
                cf_keys = ["営業CF", "投資CF", "財務CF"]
                company_names = [v["name"] for v in co_summaries.values()]
                company_ids = list(co_summaries.keys())

                fig_cf_bar = go.Figure()
                for i, cid_c in enumerate(company_ids):
                    vals = [co_summaries[cid_c]["sums"][k] for k in cf_keys]
                    fig_cf_bar.add_trace(go.Bar(
                        y=cf_categories, x=vals,
                        name=company_names[i],
                        orientation="h",
                        marker_color=COMPANY_COLORS.get(cid_c, "#888"),
                        text=[f"{v:+,.0f}" for v in vals],
                        textposition="outside",
                        textfont=dict(size=13, weight="bold"),
                        hovertemplate=f"<b>{company_names[i]}</b><br>%{{y}}: %{{x:+,.0f}} 千円<extra></extra>",
                    ))
                fig_cf_bar.update_layout(
                    barmode="group", height=300,
                    xaxis=dict(title="千円", **Y_AXIS_FORMAT),
                    yaxis=dict(autorange="reversed"),
                    plot_bgcolor="white",
                    hoverlabel=dict(font_size=13),
                    legend=dict(orientation="h", y=1.15, x=0.5, xanchor="center",
                                font=dict(size=12), bgcolor="rgba(255,255,255,0.8)"),
                    margin=dict(l=160, r=80, t=40, b=40),
                )
                fig_cf_bar.add_vline(x=0, line_color="#999", line_width=1)
                st.plotly_chart(fig_cf_bar, use_container_width=True)

                # 期間合計の比較（横棒）
                st.markdown("##### 💰 法人別：期間キャッシュフロー合計")
                totals = [co_summaries[cid]["sums"]["月次CF"] for cid in company_ids]
                bar_colors = [COMPANY_COLORS.get(cid, "#888") for cid in company_ids]
                fig_total = go.Figure(go.Bar(
                    y=company_names, x=totals,
                    orientation="h",
                    marker_color=bar_colors,
                    text=[f"{v:+,.0f} 千円" for v in totals],
                    textposition="auto",
                    textfont=dict(size=14, weight="bold", color="white"),
                    insidetextanchor="middle",
                    hovertemplate="<b>%{y}</b><br>期間CF合計: %{x:+,.0f} 千円<extra></extra>",
                ))
                fig_total.update_layout(
                    height=220,
                    xaxis=dict(title="千円", **Y_AXIS_FORMAT),
                    yaxis=dict(autorange="reversed"),
                    margin=dict(l=160, r=100, t=20, b=40),
                    showlegend=False,
                    plot_bgcolor="white",
                    hoverlabel=dict(font_size=13),
                )
                fig_total.add_vline(x=0, line_color="#999", line_width=1)
                st.plotly_chart(fig_total, use_container_width=True)

                # 法人別月次CF比較
                st.markdown("##### 📊 法人別：月別CF推移")
                fig_comp = go.Figure()
                for cid_c, cf_info in all_cf.items():
                    monthly_vals = []
                    for cal_m in calendar_months:
                        if cal_m in cf_info["months"] and "月次CF" in cf_info["data"]:
                            idx = cf_info["months"].index(cal_m)
                            vals = cf_info["data"]["月次CF"]
                            monthly_vals.append(vals[idx] if idx < len(vals) else 0)
                        else:
                            monthly_vals.append(0)
                    a_vals = [monthly_vals[i] for i in cf_active_idx]
                    fig_comp.add_trace(go.Bar(
                        x=cf_a_labels, y=a_vals, name=cf_info["name"],
                        marker_color=COMPANY_COLORS.get(cid_c, "#888"),
                        hovertemplate=f"<b>%{{x}}</b><br>{cf_info['name']}: %{{y:+,.0f}} 千円<extra></extra>",
                    ))
                fig_comp.update_layout(barmode="group", height=420,
                                       yaxis=dict(title="千円", **Y_AXIS_FORMAT), **CHART_LAYOUT)
                fig_comp.add_hline(y=0, line_color="#999", line_width=1)
                st.plotly_chart(fig_comp, use_container_width=True)

                # 法人別累計推移
                st.markdown("##### 📈 法人別：CF累計推移")
                fig_cum = go.Figure()
                for cid_c, cf_info in all_cf.items():
                    monthly_vals = []
                    for cal_m in calendar_months:
                        if cal_m in cf_info["months"] and "月次CF" in cf_info["data"]:
                            idx = cf_info["months"].index(cal_m)
                            vals = cf_info["data"]["月次CF"]
                            monthly_vals.append(vals[idx] if idx < len(vals) else 0)
                        else:
                            monthly_vals.append(0)
                    a_vals = [monthly_vals[i] for i in cf_active_idx]
                    cum = []
                    r = 0
                    for v in a_vals:
                        r += v
                        cum.append(r)
                    fig_cum.add_trace(go.Scatter(
                        x=cf_a_labels, y=cum, name=cf_info["name"],
                        mode="lines+markers",
                        line=dict(color=COMPANY_COLORS.get(cid_c, "#888"), width=3),
                        marker=dict(size=8),
                        hovertemplate=f"<b>%{{x}}</b><br>{cf_info['name']} 累計: %{{y:+,.0f}} 千円<extra></extra>",
                    ))
                fig_cum.update_layout(height=400,
                                      yaxis=dict(title="千円", **Y_AXIS_FORMAT), **CHART_LAYOUT)
                fig_cum.add_hline(y=0, line_color="#999", line_width=1)
                st.plotly_chart(fig_cum, use_container_width=True)
        else:
            st.info("CFデータがありません")

    # ====================
    # 3. BSポジション
    # ====================
    st.divider()
    st.subheader("🏦 BSポジション（貸借対照表）")

    BS_COLORS = {
        "現金及び預金合計": "#3498db",
        "流動資産合計": "#2ecc71",
        "流動負債合計": "#e74c3c",
        "純資産の部合計": "#9b59b6",
    }
    BS_LABELS = {
        "現金及び預金合計": "現金預金",
        "流動資産合計": "流動資産",
        "流動負債合計": "流動負債",
        "純資産の部合計": "純資産",
    }

    all_bs = {}
    all_bs_detail = {}
    for company in companies:
        cid = company["id"]
        if cid not in company_data:
            continue
        cd = company_data[cid]
        fy_start = company.get("fiscal_year_start", 7)
        try:
            bs_raw = read_bs(str(cd["file"]), cid)
            bs_months, bs_data = extract_bs_trend(bs_raw, fy_start)
            if bs_months and bs_data:
                bs_months, bs_data = _trim_zero_tail(bs_months, bs_data)
                bs_months, bs_data = _drop_settlement_month(bs_months, bs_data, fy_start)
                skip = 0
                if bs_months and "期首" in bs_months[0]:
                    skip = 1
                month_only = bs_months[skip:]
                data_only = {k: v[skip:] for k, v in bs_data.items()}
                cal_months = fiscal_months_to_calendar(month_only, fy_start, selected_month)
                all_bs[cid] = {
                    "name": company["name"],
                    "months": cal_months,
                    "data": data_only,
                    "fy_start": fy_start,
                }
                _, full_data = extract_bs_full(bs_raw, fy_start)
                detail_trimmed = {}
                for dk, dv in full_data.items():
                    trimmed = dv[skip:len(month_only) + skip] if len(dv) > skip else dv
                    if len(trimmed) == len(cal_months):
                        detail_trimmed[dk] = trimmed
                    elif len(trimmed) > 0:
                        detail_trimmed[dk] = trimmed[:len(cal_months)]
                all_bs_detail[cid] = detail_trimmed
        except Exception:
            pass

    bs_tab_names = [c["name"] for c in companies] + ["3社合算"]
    bs_tabs = st.tabs(bs_tab_names)

    DETAIL_COLORS = [
        "#1abc9c", "#e67e22", "#3498db", "#e74c3c", "#9b59b6",
        "#f1c40f", "#2ecc71", "#e84393", "#00cec9", "#636e72",
    ]

    def _render_bs_chart(months, data, height=480):
        fig = go.Figure()
        hover_bs = "<b>%{x}</b><br>%{fullData.name}: %{y:,.0f} 千円<extra></extra>"
        for key, vals in data.items():
            fig.add_trace(go.Scatter(
                x=months[:len(vals)], y=vals,
                name=BS_LABELS.get(key, key),
                mode="lines+markers",
                line=dict(color=BS_COLORS.get(key, "#888"), width=3),
                marker=dict(size=8),
                hovertemplate=hover_bs,
            ))
        fig.update_layout(height=height, yaxis=dict(title="千円", **Y_AXIS_FORMAT), **CHART_LAYOUT)
        st.plotly_chart(fig, use_container_width=True)

    def _render_bs_table(data):
        bs_rows = []
        for key, vals in data.items():
            nonzero = [v for v in vals if v != 0]
            latest = nonzero[-1] if nonzero else 0
            first = vals[0] if vals else 0
            diff = latest - first
            bs_rows.append({
                "項目": BS_LABELS.get(key, key),
                "期首": f"{first:,.0f}",
                "直近月": f"{latest:,.0f}",
                "増減": f"{diff:+,.0f}",
            })
        st.dataframe(pd.DataFrame(bs_rows), use_container_width=True, hide_index=True)

    def _render_bs_detail_expander(parent_key, months, detail_data, parent_data=None):
        """親項目の内訳をexpanderで表示する。"""
        if parent_key not in BS_DETAIL_MAP:
            return
        sub_items = BS_DETAIL_MAP[parent_key]["sub_items"]
        available = [(name, detail_data[name]) for name, _ in sub_items if name in detail_data]
        if not available:
            return
        label = BS_LABELS.get(parent_key, parent_key)
        parent_vals = (parent_data or {}).get(parent_key) or detail_data.get(parent_key) or []
        nonzero_vals = [v for v in parent_vals if v != 0]
        latest_total = nonzero_vals[-1] if nonzero_vals else None
        header = f"📊 {label}の内訳"
        if latest_total is not None:
            header += f"（合計: {latest_total:,.0f} 千円）"
        with st.expander(header):
            fig = go.Figure()
            hover_d = "<b>%{x}</b><br>%{fullData.name}: %{y:,.0f} 千円<extra></extra>"
            rows = []
            for ci, (name, vals) in enumerate(available):
                display_name = name.replace("合計", "")
                nz = [v for v in vals if v != 0]
                latest = nz[-1] if nz else 0
                first = vals[0] if vals else 0
                if latest == 0 and first == 0:
                    continue
                fig.add_trace(go.Scatter(
                    x=months[:len(vals)], y=vals,
                    name=display_name,
                    mode="lines+markers",
                    line=dict(color=DETAIL_COLORS[ci % len(DETAIL_COLORS)], width=2),
                    marker=dict(size=6),
                    hovertemplate=hover_d,
                ))
                rows.append({
                    "項目": display_name,
                    "期初": f"{first:,.0f}",
                    "直近月": f"{latest:,.0f}",
                    "増減": f"{latest - first:+,.0f}",
                })
            if fig.data:
                fig.update_layout(
                    height=350, yaxis=dict(title="千円", **Y_AXIS_FORMAT),
                    plot_bgcolor="white", hoverlabel=dict(font_size=13),
                    legend=dict(orientation="h", y=-0.25, x=0.5, xanchor="center", font=dict(size=11)),
                    margin=dict(l=60, r=30, t=20, b=70),
                )
                st.plotly_chart(fig, use_container_width=True)
            if rows:
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    for idx, company in enumerate(companies):
        cid = company["id"]
        with bs_tabs[idx]:
            if cid in all_bs:
                bs_info = all_bs[cid]
                detail = all_bs_detail.get(cid, {})

                kpi_cols = st.columns(4)
                for i, key in enumerate(["現金及び預金合計", "流動資産合計", "流動負債合計", "純資産の部合計"]):
                    if key in bs_info["data"]:
                        nonzero = [v for v in bs_info["data"][key] if v != 0]
                        val = nonzero[-1] if nonzero else 0
                        c = "#27ae60" if val >= 0 else "#c0392b"
                        with kpi_cols[i]:
                            kpi_card(BS_LABELS[key], val, "千円", color=c)

                _render_bs_chart(bs_info["months"], bs_info["data"])
                _render_bs_table(bs_info["data"])

                if detail:
                    st.markdown("##### 🔍 項目別内訳（クリックで展開）")
                    for key in ["現金及び預金合計", "流動資産合計", "流動負債合計", "純資産の部合計"]:
                        _render_bs_detail_expander(key, bs_info["months"], detail, bs_info["data"])
            else:
                st.info("BSデータがありません")

    with bs_tabs[-1]:
        if all_bs:
            all_cal = set()
            for info in all_bs.values():
                all_cal.update(info["months"])
            sorted_cal = sorted(all_cal)

            bs_keys = ["現金及び預金合計", "流動資産合計", "流動負債合計", "純資産の部合計"]
            consolidated = {}
            for key in bs_keys:
                totals = []
                for cal_m in sorted_cal:
                    s = 0
                    for info in all_bs.values():
                        if cal_m in info["months"] and key in info["data"]:
                            mi = info["months"].index(cal_m)
                            vals = info["data"][key]
                            if mi < len(vals):
                                s += vals[mi]
                    totals.append(s)
                consolidated[key] = totals

            kpi_cols = st.columns(4)
            for i, key in enumerate(bs_keys):
                nonzero = [v for v in consolidated[key] if v != 0]
                val = nonzero[-1] if nonzero else 0
                c = "#27ae60" if val >= 0 else "#c0392b"
                with kpi_cols[i]:
                    kpi_card(f"3社{BS_LABELS[key]}", val, "千円", color=c)

            _render_bs_chart(sorted_cal, consolidated)

            st.markdown("##### 法人別内訳（直近月）")
            for key in bs_keys:
                cols = st.columns(len(all_bs) + 1)
                label = BS_LABELS[key]
                with cols[0]:
                    st.markdown(f"**{label}**")
                grand = 0
                for ci, (cid_b, info) in enumerate(all_bs.items()):
                    nonzero = [v for v in info["data"].get(key, []) if v != 0]
                    val = nonzero[-1] if nonzero else 0
                    grand += val
                    with cols[ci + 1]:
                        c = "#27ae60" if val >= 0 else "#c0392b"
                        st.metric(info["name"], f"{val:,.0f}", label_visibility="visible")

            _render_bs_table(consolidated)

            # 3社合算の内訳expander
            all_detail_keys = set()
            for d in all_bs_detail.values():
                all_detail_keys.update(d.keys())
            if all_detail_keys:
                consolidated_detail = {}
                for dk in all_detail_keys:
                    totals = []
                    for cal_m in sorted_cal:
                        s = 0
                        for cid_d, info in all_bs.items():
                            det = all_bs_detail.get(cid_d, {})
                            if cal_m in info["months"] and dk in det:
                                mi = info["months"].index(cal_m)
                                dvals = det[dk]
                                if mi < len(dvals):
                                    s += dvals[mi]
                        totals.append(s)
                    if any(v != 0 for v in totals):
                        consolidated_detail[dk] = totals
                if consolidated_detail:
                    st.markdown("##### 🔍 項目別内訳（クリックで展開）")
                    for key in bs_keys:
                        _render_bs_detail_expander(key, sorted_cal, consolidated_detail, consolidated)
        else:
            st.info("BSデータがありません")


if __name__ == "__main__":
    main()
