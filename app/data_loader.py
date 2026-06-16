"""月次Excelファイルからキーメトリクスを抽出するローダー

税理士ミーティング用に、3社の月次報告Excelから主要数値を取得する。
Excelの構造は税理士事務所のフォーマット（BS/PL/CF推移）に準拠。
"""
import pandas as pd
import yaml
import glob
import re
from pathlib import Path

CONFIG_DIR = Path(__file__).parent / "config"
COMPANIES_FILE = CONFIG_DIR / "companies.yaml"


def _load_config():
    with open(COMPANIES_FILE, "r") as f:
        return yaml.safe_load(f)


def get_monthly_dir() -> Path:
    config = _load_config()
    rel = config.get("data", {}).get("monthly_dir", "../data/monthly")
    return (Path(__file__).parent / rel).resolve()


def find_latest_file(company_id: str, yyyymm: str = None) -> Path:
    """指定法人の最新（または指定月）のExcelファイルを検索する。"""
    config = _load_config()
    company = next((c for c in config["companies"] if c["id"] == company_id), None)
    if not company:
        raise ValueError(f"Unknown company: {company_id}")

    monthly_dir = get_monthly_dir()
    prefix = company["file_prefix"]

    if yyyymm:
        pattern = f"{prefix}-{yyyymm}.xlsx"
        matches = list(monthly_dir.glob(pattern))
        if not matches:
            pattern = f"{prefix}{yyyymm}.xlsx"
            matches = list(monthly_dir.glob(pattern))
    else:
        matches = sorted(monthly_dir.glob(f"{prefix}*.xlsx"), reverse=True)

    if not matches:
        return None
    return matches[0]


def list_available_files() -> dict:
    """データディレクトリにある全ファイルを法人別に整理して返す。"""
    config = _load_config()
    monthly_dir = get_monthly_dir()
    result = {}
    for company in config["companies"]:
        prefix = company["file_prefix"]
        files = sorted(monthly_dir.glob(f"{prefix}*.xlsx"))
        result[company["id"]] = {
            "name": company["name"],
            "files": [f.name for f in files],
        }
    return result


def extract_pl_summary(file_path: Path, company_config: dict) -> dict:
    """PLシートから売上・原価・粗利・販管費・営業利益の累計を抽出する。"""
    sheet_name = company_config["sheets"].get("pl", "月次PL")
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    except Exception as e:
        return {"error": str(e)}

    result = {}
    targets = {
        "売上高合計": "売上高",
        "売上原価": "売上原価",
        "売上総利益": "粗利益",
        "販売費及び一般管理費": "販管費",
    }

    for row_idx in range(len(df)):
        cell = str(df.iloc[row_idx, 0]) if pd.notna(df.iloc[row_idx, 0]) else ""
        for key, label in targets.items():
            if key in cell:
                cumulative_col = df.shape[1] - 1
                for c in range(df.shape[1] - 1, 0, -1):
                    val = df.iloc[row_idx, c]
                    if pd.notna(val) and val != 0:
                        cumulative_col = c
                        break
                cum_val = df.iloc[row_idx, cumulative_col]
                if pd.notna(cum_val):
                    result[label] = int(cum_val) if isinstance(cum_val, (int, float)) else cum_val

    return result


def extract_bs_snapshot(file_path: Path, company_config: dict) -> dict:
    """BS千円シートから最新月の主要残高を抽出する。"""
    sheet_name = company_config["sheets"].get("bs_thousand", "BS千円")
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    except Exception as e:
        return {"error": str(e)}

    result = {}
    targets = ["現金及び預金合計", "流動資産合計", "固定資産合計", "流動負債合計", "固定負債合計", "純資産合計"]

    for row_idx in range(len(df)):
        cell = str(df.iloc[row_idx, 0]) if pd.notna(df.iloc[row_idx, 0]) else ""
        for target in targets:
            if target in cell:
                for c in range(df.shape[1] - 1, 1, -1):
                    val = df.iloc[row_idx, c]
                    if pd.notna(val) and val != 0:
                        result[target] = round(float(val), 1) if isinstance(val, (int, float)) else val
                        break

    return result


def extract_cf_summary(file_path: Path, company_config: dict) -> dict:
    """CF推移シートから各CFと月次CF合計を抽出する。"""
    sheet_name = company_config["sheets"].get("cf", "CF推移")
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    except Exception as e:
        return {"error": str(e)}

    result = {}
    targets = ["１．営業活動CF", "２．投資活動CF", "３．財務活動CF", "４．月次CF"]

    for row_idx in range(len(df)):
        for col_idx in range(df.shape[1]):
            cell = str(df.iloc[row_idx, col_idx]) if pd.notna(df.iloc[row_idx, col_idx]) else ""
            for target in targets:
                if target in cell:
                    label = target.split("．")[1] if "．" in target else target
                    monthly_vals = []
                    for c in range(col_idx + 2, df.shape[1]):
                        val = df.iloc[row_idx, c]
                        if pd.notna(val) and isinstance(val, (int, float)):
                            monthly_vals.append(round(float(val), 1))
                    if monthly_vals:
                        result[label] = {
                            "latest": monthly_vals[-1] if monthly_vals else None,
                            "cumulative": round(sum(monthly_vals), 1),
                            "months": len(monthly_vals),
                        }

    return result


def extract_company_summary(company_id: str, yyyymm: str = None) -> dict:
    """指定法人の月次データサマリーを抽出する。"""
    config = _load_config()
    company = next((c for c in config["companies"] if c["id"] == company_id), None)
    if not company:
        return {"error": f"Unknown company: {company_id}"}

    file_path = find_latest_file(company_id, yyyymm)
    if not file_path:
        return {"error": f"No data file found for {company['name']}"}

    return {
        "company": company["name"],
        "legal_entity": company["legal_entity"],
        "business_type": company["business_type"],
        "file": file_path.name,
        "pl": extract_pl_summary(file_path, company),
        "bs": extract_bs_snapshot(file_path, company),
        "cf": extract_cf_summary(file_path, company),
    }


def extract_all_summaries(yyyymm: str = None) -> list:
    """全法人のサマリーを一括取得する。"""
    config = _load_config()
    return [extract_company_summary(c["id"], yyyymm) for c in config["companies"]]


if __name__ == "__main__":
    import json
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == "files":
        files = list_available_files()
        for cid, info in files.items():
            print(f"\n{info['name']} ({cid}):")
            for f in info["files"]:
                print(f"  - {f}")
        sys.exit(0)

    yyyymm = sys.argv[1] if len(sys.argv) > 1 else None
    summaries = extract_all_summaries(yyyymm)
    for s in summaries:
        print(f"\n{'='*60}")
        print(f"{s.get('company', '?')} ({s.get('legal_entity', '')})")
        print(f"{'='*60}")
        if "error" in s:
            print(f"  ❌ {s['error']}")
            continue
        print(f"  ファイル: {s['file']}")
        print(f"  業種: {s['business_type']}")

        pl = s.get("pl", {})
        if pl and "error" not in pl:
            print(f"\n  [PL累計]")
            for k, v in pl.items():
                print(f"    {k}: {v:,.0f}" if isinstance(v, (int, float)) else f"    {k}: {v}")

        bs = s.get("bs", {})
        if bs and "error" not in bs:
            print(f"\n  [BS残高（千円）]")
            for k, v in bs.items():
                print(f"    {k}: {v:,.1f}" if isinstance(v, (int, float)) else f"    {k}: {v}")

        cf = s.get("cf", {})
        if cf and "error" not in cf:
            print(f"\n  [CF推移（千円）]")
            for k, v in cf.items():
                if isinstance(v, dict):
                    print(f"    {k}: 直近月 {v.get('latest', '?'):,.1f} / 累計 {v.get('cumulative', '?'):,.1f} ({v.get('months', '?')}ヶ月)")
