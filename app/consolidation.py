"""esグループ 連結決算モジュール

3法人（+関連会社）の月次PL/BSを暦月ベースで合算し、
グループ内部取引を消去して「グループ外からの真水の売上」を算出する。

## なぜ暦月ベースか

各社の決算期が揃っていない:
  - esエンターテイメント / エスクリエイト : 6月決算（期首7月）
  - esライフワーク                        : 3月決算（期首4月）

決算期が異なる会社を「期首からのn番目の月」で足すと、
エンタの7月とLWの4月が同じ列に乗ってしまい意味のない合計になる。
そのため本モジュールは全社を暦月（YYYYMM）に正規化してから合算する。

## 内部取引の消去

各社の月次Excelには取引相手先の情報が一切無い（売上はセグメント別、
費用は勘定科目別にしか分かれていない）。エンタに至っては業務委託費が
人件費に合算されている。

したがって消去すべき内部取引は `config/intercompany.yaml` で明示的に
宣言する。金額が未確認の取引は status: unknown とし、
消去額に算入せず「未確認」として警告に出す。推定値は入れない。
"""
from pathlib import Path

import pandas as pd
import yaml

APP_DIR = Path(__file__).resolve().parent
CONFIG_DIR = APP_DIR / "config"
DATA_DIR = APP_DIR.parent / "data" / "monthly"
HISTORY_DIR = APP_DIR.parent / "data" / "history"
COMPANIES_FILE = CONFIG_DIR / "companies.yaml"
INTERCOMPANY_FILE = CONFIG_DIR / "intercompany.yaml"

PL_SECTIONS = ("売上高", "売上原価", "粗利益", "販売管理費",
               "営業利益", "経常利益", "営業外収益", "営業外費用")

# 連結の「期」の起点。エンタ・エスクリエイトの期首（7月）に合わせる。
# ライフワークは3月決算のため、LWの数字は自社の期をまたいで暦月で割り当てられる。
GROUP_FY_START = 7

# 集計対象のPL項目
PL_KEYS = ("売上高", "売上原価", "粗利益", "販管費", "営業利益",
           "営業外収益", "営業外費用", "経常利益")


# ============================================================
# 設定の読み込み
# ============================================================

def load_companies():
    with open(COMPANIES_FILE, "r") as f:
        return yaml.safe_load(f)["companies"]


def load_intercompany():
    """内部取引の定義を読む。ファイルが無ければ空の定義を返す。"""
    if not INTERCOMPANY_FILE.exists():
        return {"transactions": [], "balances": []}
    with open(INTERCOMPANY_FILE, "r") as f:
        cfg = yaml.safe_load(f) or {}
    cfg.setdefault("transactions", [])
    cfg.setdefault("balances", [])
    return cfg


# ============================================================
# 暦月の正規化
# ============================================================

def fiscal_year_start_year(file_yyyymm: str, fiscal_start_month: int) -> int:
    """ファイルの対象月から、その事業年度の期首の「年」を求める。"""
    year, month = int(file_yyyymm[:4]), int(file_yyyymm[4:])
    return year if month >= fiscal_start_month else year - 1


def fiscal_columns_to_calendar(file_yyyymm: str, fiscal_start_month: int) -> list:
    """期首からn番目の月 → 暦月(YYYYMM) の対応表を12ヶ月分作る。"""
    start_year = fiscal_year_start_year(file_yyyymm, fiscal_start_month)
    out = []
    for i in range(12):
        m = (fiscal_start_month - 1 + i) % 12 + 1
        y = start_year + (1 if m < fiscal_start_month else 0)
        out.append(f"{y}{m:02d}")
    return out


# ============================================================
# 月次PLの抽出
# ============================================================

def _pl_rows(path: Path):
    """PL推移シートを (セクション, 項目, 12ヶ月の値) の形に読み替える。"""
    d = pd.read_excel(path, sheet_name="PL推移", header=None)
    rows, section = [], None
    for _, r in d.iterrows():
        cells = r.tolist()
        c1 = str(cells[1]).strip()
        c2 = str(cells[2]).strip()
        if c1 in PL_SECTIONS:
            section = c1
        # 値は 3, 5, 7, ... と1列おき（間は構成比）
        vals = []
        for k in range(12):
            idx = 3 + 2 * k
            v = cells[idx] if idx < len(cells) else None
            vals.append(float(v) if isinstance(v, (int, float)) and not pd.isna(v) else 0.0)
        rows.append((section, c1, c2, vals))
    return rows


def monthly_pl(company: dict, file_yyyymm: str) -> dict:
    """1社の月次PLを暦月ベースで返す。

    戻り値: {YYYYMM: {"売上高":x, "売上原価":x, "粗利益":x, "販管費":x, "営業利益":x}}
    """
    prefix = company["file_prefix"]
    path = DATA_DIR / f"{prefix}-{file_yyyymm}.xlsx"
    if not path.exists():
        return {}

    rows = _pl_rows(path)
    calendar = fiscal_columns_to_calendar(file_yyyymm, company["fiscal_year_start"])

    totals = {}
    for section, c1, c2, vals in rows:
        # 各セクションの見出し行（c1がセクション名そのもの）が合計値を持つ
        if c1 in PL_SECTIONS and c2 in ("", "nan"):
            totals[c1] = vals

    out = {}
    for i, ym in enumerate(calendar):
        rev = totals.get("売上高", [0] * 12)[i]
        cogs = totals.get("売上原価", [0] * 12)[i]
        gross = totals.get("粗利益", [0] * 12)[i]
        sga = totals.get("販売管理費", [0] * 12)[i]
        op = totals.get("営業利益", [0] * 12)[i]
        noi = totals.get("営業外収益", [0] * 12)[i]
        noe = totals.get("営業外費用", [0] * 12)[i]
        if rev == 0 and sga == 0 and gross == 0:
            continue  # 未入力の月は連結対象外
        gross_v = gross if gross else rev - cogs
        op_v = op if op else gross_v - sga
        out[ym] = {
            "売上高": rev,
            "売上原価": cogs,
            "粗利益": gross_v,
            "販管費": sga,
            "営業利益": op_v,
            "営業外収益": noi,
            "営業外費用": noe,
            "経常利益": op_v + noi - noe,
        }
    return out


def available_file_months(company: dict) -> list:
    """その会社の月次Excelが存在する対象月(YYYYMM)を昇順で返す。"""
    import re
    prefix = company["file_prefix"]
    out = []
    for path in DATA_DIR.glob(f"{prefix}-*.xlsx"):
        m = re.search(r"(\d{6})", path.name)
        if m:
            out.append(m.group(1))
    return sorted(set(out))


def all_monthly_pl(company: dict, upto_yyyymm: str) -> dict:
    """利用可能な月次Excelを全て併合して暦月PLを返す。

    月次Excelは1事業年度分しか持たないため、1ファイルだけを見ると
    連結できる期間が短くなる（例: 202607のファイルはエンタなら7月の1ヶ月分のみ）。
    古いファイルから順に併合し、新しいファイルで上書きすることで、
    修正が入った月は最新版が勝ちつつ、過去の月も拾える。
    """
    merged = {}
    for ym in available_file_months(company):
        if ym > upto_yyyymm:
            continue
        merged.update(monthly_pl(company, ym))
    # 対象月より後は落とす（未入力月の混入防止）
    return {k: v for k, v in merged.items() if k <= upto_yyyymm}


def monthly_segments(company: dict, file_yyyymm: str) -> dict:
    """セグメント別売上を暦月ベースで返す。 {YYYYMM: {セグメント名: 金額}}"""
    prefix = company["file_prefix"]
    path = DATA_DIR / f"{prefix}-{file_yyyymm}.xlsx"
    if not path.exists():
        return {}
    rows = _pl_rows(path)
    calendar = fiscal_columns_to_calendar(file_yyyymm, company["fiscal_year_start"])
    out = {ym: {} for ym in calendar}
    for section, c1, c2, vals in rows:
        if section == "売上高" and c2 not in ("", "nan", "0"):
            for i, ym in enumerate(calendar):
                if vals[i]:
                    out[ym][c2] = out[ym].get(c2, 0.0) + vals[i]
    return {k: v for k, v in out.items() if v}


# ============================================================
# 補助データソース（過年度の月次実績）
# ============================================================

def load_history_manifest():
    """data/history/manifest.yaml を読む。無ければ空。"""
    mf = HISTORY_DIR / "manifest.yaml"
    if not mf.exists():
        return []
    with open(mf, "r") as f:
        return (yaml.safe_load(f) or {}).get("sources", []) or []


def _add_months(yyyymm: str, n: int) -> str:
    y, m = int(yyyymm[:4]), int(yyyymm[4:])
    total = (y * 12 + m - 1) + n
    return f"{total // 12}{total % 12 + 1:02d}"


def supplementary_pl(company_id: str) -> dict:
    """過年度の月次PLを暦月ベースで返す（標準様式外のファイルに対応）。

    月次Excelは1事業年度分しか持たないため、決算期の異なる会社を暦月で
    12ヶ月連結するには過年度の実績が要る。manifest.yaml で様式を宣言する。
    """
    out = {}
    for src in load_history_manifest():
        if src.get("company") != company_id:
            continue
        path = HISTORY_DIR / src["file"]
        if not path.exists():
            continue
        try:
            df = pd.read_excel(path, sheet_name=src["sheet"], header=None)
        except Exception:
            continue

        scale = 0.001 if src.get("unit") == "円" else 1.0
        first_col = int(src.get("first_col", 2))
        count = int(src.get("month_count", 12))
        first_month = str(src["first_month"])

        # ラベル → 行 の対応を作る（行番号ではなくラベルで拾う）
        label_to_key = {v: k for k, v in src.get("rows", {}).items()}
        label_cols = src.get("label_cols", [0, 1])
        found = {}
        for _, r in df.iterrows():
            cells = r.tolist()
            for lc in label_cols:
                if lc >= len(cells):
                    continue
                label = str(cells[lc]).strip()
                key = label_to_key.get(label)
                if key and key not in found:
                    found[key] = cells
                    break

        for i in range(count):
            ym = _add_months(first_month, i)
            row = {}
            for key, cells in found.items():
                idx = first_col + i
                v = cells[idx] if idx < len(cells) else None
                row[key] = (float(v) * scale) if isinstance(v, (int, float)) and not pd.isna(v) else 0.0
            if any(row.get(k) for k in ("売上高", "販管費")):
                op_v = row.get("営業利益", 0.0)
                out[ym] = {
                    "売上高": row.get("売上高", 0.0),
                    "売上原価": row.get("売上原価", 0.0),
                    "粗利益": row.get("粗利益", 0.0),
                    "販管費": row.get("販管費", 0.0),
                    "営業利益": op_v,
                    "営業外収益": 0.0,
                    "営業外費用": 0.0,
                    "経常利益": row.get("経常利益", op_v),
                }
    return out


# ============================================================
# 内部取引の消去
# ============================================================

def elimination_for_month(cfg: dict, yyyymm: str) -> dict:
    """指定月の消去額を集計する。

    戻り値:
      {"amount": 確定分の消去額, "unknown": [金額未確認の取引名], "items": [明細]}
    """
    total, unknown, items = 0.0, [], []
    for tx in cfg.get("transactions", []):
        amounts = tx.get("amounts") or {}
        amt = amounts.get(yyyymm) or amounts.get(int(yyyymm))
        status = tx.get("status", "unknown")
        if amt is None or status == "unknown":
            if tx.get("active", True):
                unknown.append(tx.get("name", tx.get("id", "?")))
            continue
        total += float(amt)
        items.append({
            "name": tx.get("name", tx.get("id")),
            "seller": tx.get("seller"),
            "buyer": tx.get("buyer"),
            "amount": float(amt),
            "status": status,
        })
    return {"amount": total, "unknown": unknown, "items": items}


def consolidate(file_yyyymm: str, months: list = None) -> dict:
    """連結PLを計算する。

    Args:
        file_yyyymm: 各社の月次Excelの対象月（例 "202606"）
        months:      集計する暦月のリスト。Noneなら全社のデータがある月すべて。

    戻り値:
        {
          "months": [...],
          "by_company": {company_id: {YYYYMM: {...}}},
          "simple_sum": {YYYYMM: {...}},        # 単純合算
          "elimination": {YYYYMM: {...}},       # 内部取引消去額
          "consolidated": {YYYYMM: {...}},      # 連結（真水）
          "unknown": {YYYYMM: [取引名]},        # 金額未確認の取引
          "coverage": {YYYYMM: {present, missing, complete}},  # その月にデータがある会社
        }

    注意: 決算期が異なるため、1つの月次Excelに載る12ヶ月は会社ごとにずれる。
    3社すべてのデータが揃う月（coverage[ym]["complete"] が True）以外は、
    一部の会社しか含まない不完全な合計になる。連結値として使ってよいのは
    complete な月だけ。complete_months() で取得できる。
    """
    companies = load_companies()
    cfg = load_intercompany()

    by_company = {}
    for c in companies:
        pl = all_monthly_pl(c, file_yyyymm)
        hist = {k: v for k, v in supplementary_pl(c["id"]).items() if k <= file_yyyymm}
        # 月次Excelを優先し、足りない過去月を補助ソースで埋める
        merged = {**hist, **pl}
        if merged:
            by_company[c["id"]] = merged

    all_months = sorted({ym for pl in by_company.values() for ym in pl})
    if months:
        all_months = [m for m in all_months if m in months]

    simple, elim, cons, unknown, coverage = {}, {}, {}, {}, {}
    total_companies = len(by_company)
    for ym in all_months:
        agg = {k: 0.0 for k in PL_KEYS}
        present = []
        for cid, pl in by_company.items():
            row = pl.get(ym)
            if not row:
                continue
            present.append(cid)
            for k in agg:
                agg[k] += row.get(k, 0.0)
        simple[ym] = agg
        coverage[ym] = {
            "present": present,
            "missing": [c for c in by_company if c not in present],
            "complete": len(present) == total_companies,
        }

        e = elimination_for_month(cfg, ym)
        elim[ym] = e
        unknown[ym] = e["unknown"]

        # 内部売上は、売り手の売上高と買い手の費用（販管費 or 原価）を同額落とす。
        # 利益への影響はゼロ、売上・費用の総額だけが縮む。
        cons[ym] = dict(agg)
        cons[ym]["売上高"] = agg["売上高"] - e["amount"]
        cons[ym]["粗利益"] = agg["粗利益"] - e["amount"]
        cons[ym]["販管費"] = agg["販管費"] - e["amount"]

    return {
        "months": all_months,
        "by_company": by_company,
        "simple_sum": simple,
        "elimination": elim,
        "consolidated": cons,
        "unknown": unknown,
        "coverage": coverage,
    }


def period_total(d: dict, months: list) -> dict:
    """月次dictを指定期間で合計する。"""
    out = {k: 0.0 for k in PL_KEYS}
    for ym in months:
        row = d.get(ym)
        if not row:
            continue
        for k in out:
            out[k] += row.get(k, 0.0)
    return out


def complete_months(result: dict) -> list:
    """3社すべてのデータが揃っている暦月だけを返す。"""
    return [ym for ym in result["months"] if result["coverage"][ym]["complete"]]


def coverage_note(result: dict) -> str:
    """連結の網羅状況を1行で説明する。"""
    comp = complete_months(result)
    if not comp:
        return "全社のデータが揃う月がありません。連結値は算出できません。"
    partial = [ym for ym in result["months"] if ym not in comp]
    msg = f"全社データが揃う月: {comp[0]}〜{comp[-1]}（{len(comp)}ヶ月）"
    if partial:
        msg += f" ／ 一部の会社のみ: {len(partial)}ヶ月（連結値には使えません）"
    return msg


# ============================================================
# 決算期ごとの集計
# ============================================================

def fiscal_period_end(yyyymm: str, fy_start: int = GROUP_FY_START) -> str:
    """その暦月が属する期の期末(YYYYMM)を返す。"""
    y, m = int(yyyymm[:4]), int(yyyymm[4:])
    end_month = fy_start - 1 if fy_start > 1 else 12
    if fy_start == 1:
        ey = y
    elif m >= fy_start:
        ey = y + 1
    else:
        ey = y
    return f"{ey}{end_month:02d}"


def period_label(period_end: str) -> str:
    """期末(YYYYMM)を「令和N年M月期」に変換する。"""
    y, m = int(period_end[:4]), int(period_end[4:])
    reiwa = y - 2018
    return f"令和{reiwa}年{m}月期" if reiwa > 0 else f"{y}年{m}月期"


def period_range_label(period_end: str, fy_start: int = GROUP_FY_START) -> str:
    """「2025年7月〜2026年6月」形式の期間表記を返す。"""
    ey, em = int(period_end[:4]), int(period_end[4:])
    sm = fy_start
    sy = ey - 1 if fy_start > 1 else ey
    return f"{sy}年{sm}月〜{ey}年{em}月"


def group_by_fiscal_period(result: dict, fy_start: int = GROUP_FY_START) -> list:
    """連結結果を決算期ごとにまとめる。

    エンタ・エスクリエイトは6月決算（期首7月）のため、7月から新しい期に入る。
    連結もその区切りで見られるように、暦月を期ごとに束ねる。

    戻り値: 期末の新しい順に
      [{
        "period_end": "202606",
        "label": "令和8年6月期",
        "range": "2025年7月〜2026年6月",
        "months": [...],            # 全社データが揃う暦月のみ
        "all_months": [...],        # データがある暦月（不完全な月を含む）
        "elapsed": 12,              # 揃っている月数
        "simple_sum": {...},
        "consolidated": {...},
        "elimination": 消去額,
        "unknown": [未確認の取引名],
        "by_company": {cid: {...}},
      }, ...]
    """
    buckets = {}
    for ym in result["months"]:
        buckets.setdefault(fiscal_period_end(ym, fy_start), []).append(ym)

    out = []
    for pe in sorted(buckets, reverse=True):
        all_months = sorted(buckets[pe])
        months = [m for m in all_months if result["coverage"][m]["complete"]]
        if not months:
            continue
        elim = sum(result["elimination"][m]["amount"] for m in months)
        unknown = sorted({n for m in months for n in result["unknown"][m]})
        by_co = {}
        for cid, pl in result["by_company"].items():
            t = period_total(pl, months)
            if any(t.values()):
                by_co[cid] = t
        out.append({
            "period_end": pe,
            "label": period_label(pe),
            "range": period_range_label(pe, fy_start),
            "months": months,
            "all_months": all_months,
            "elapsed": len(months),
            "simple_sum": period_total(result["simple_sum"], months),
            "consolidated": period_total(result["consolidated"], months),
            "elimination": elim,
            "unknown": unknown,
            "by_company": by_co,
        })
    return out


def company_by_fiscal_period(result: dict, company_id: str, fy_start: int) -> list:
    """1社を、その会社自身の決算期で束ねて返す。

    グループの連結はエンタ・エスクリエイトの期首（7月）で区切るため、
    3月決算のライフワークは自社の期をまたいで割り当てられる。
    LWの決算書と突き合わせたい場合に、自社の期（4月〜3月）で見るための関数。
    """
    pl = result["by_company"].get(company_id, {})
    if not pl:
        return []
    buckets = {}
    for ym in pl:
        buckets.setdefault(fiscal_period_end(ym, fy_start), []).append(ym)

    out = []
    for pe in sorted(buckets, reverse=True):
        months = sorted(buckets[pe])
        out.append({
            "period_end": pe,
            "label": period_label(pe),
            "range": period_range_label(pe, fy_start),
            "months": months,
            "elapsed": len(months),
            "total": period_total(pl, months),
        })
    return out
