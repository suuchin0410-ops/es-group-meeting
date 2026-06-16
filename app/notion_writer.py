"""税理士ミーティング用 Notion書き込みユーティリティ

bokashiプロジェクトのnotion_writer.pyをベースに、
親ページIDを設定ファイルから読み込む形に拡張。

主な機能:
  - create_meeting_page(): 税理士ミーティングページを自動作成
  - append_blocks(): ブロック追記
  - get_blocks() / delete_blocks(): ブロック操作
  - create_transcript_toggle(): 文字起こし用トグル作成
  - ブロック生成ヘルパー: heading(), text_block(), bulleted(), callout(), table(), divider(), toggle()
"""
import requests
import json
import re
import yaml
from pathlib import Path
from datetime import date

CONFIG_DIR = Path(__file__).parent / "config"
TOKEN_FILE = CONFIG_DIR / ".notion_token"
COMPANIES_FILE = CONFIG_DIR / "companies.yaml"
NOTION_VERSION = "2022-06-28"


def _load_config():
    with open(COMPANIES_FILE, "r") as f:
        return yaml.safe_load(f)


def _get_parent_page_id():
    config = _load_config()
    pid = config.get("notion", {}).get("parent_page_id", "")
    if not pid:
        raise ValueError(
            "Notion親ページIDが未設定です。app/config/companies.yaml の "
            "notion.parent_page_id を設定してください。"
        )
    return pid


def _get_token():
    return TOKEN_FILE.read_text().strip()


def _headers():
    return {
        "Authorization": f"Bearer {_get_token()}",
        "Content-Type": "application/json",
        "Notion-Version": NOTION_VERSION,
    }


def extract_page_id(url_or_id: str) -> str:
    url_or_id = url_or_id.strip().rstrip("/")
    uuid_pattern = re.compile(
        r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$', re.I
    )
    if uuid_pattern.match(url_or_id):
        return url_or_id
    clean = url_or_id.split("?")[0]
    raw = clean.split("-")[-1] if "-" in clean.split("/")[-1] else clean.split("/")[-1]
    raw = raw.replace("-", "")
    if len(raw) == 32:
        return f"{raw[:8]}-{raw[8:12]}-{raw[12:16]}-{raw[16:20]}-{raw[20:]}"
    return raw


# ============================================================
# ブロック生成ヘルパー
# ============================================================

def text_block(text: str, block_type: str = "paragraph", **kwargs) -> dict:
    block = {
        "object": "block",
        "type": block_type,
        block_type: {
            "rich_text": [{"type": "text", "text": {"content": text}}]
        }
    }
    if "color" in kwargs:
        block[block_type]["color"] = kwargs["color"]
    return block


def heading(text: str, level: int = 2, toggleable: bool = False) -> dict:
    htype = f"heading_{level}"
    return {
        "object": "block",
        "type": htype,
        htype: {
            "rich_text": [{"type": "text", "text": {"content": text}}],
            "is_toggleable": toggleable,
        }
    }


def divider() -> dict:
    return {"object": "block", "type": "divider", "divider": {}}


def bulleted(text: str, bold_prefix: str = None) -> dict:
    if bold_prefix:
        return {
            "object": "block",
            "type": "bulleted_list_item",
            "bulleted_list_item": {
                "rich_text": [
                    {"type": "text", "text": {"content": bold_prefix}, "annotations": {"bold": True}},
                    {"type": "text", "text": {"content": f" {text}"}},
                ]
            }
        }
    return {
        "object": "block",
        "type": "bulleted_list_item",
        "bulleted_list_item": {
            "rich_text": [{"type": "text", "text": {"content": text}}]
        }
    }


def callout(text: str, icon: str = "📌") -> dict:
    return {
        "object": "block",
        "type": "callout",
        "callout": {
            "rich_text": [{"type": "text", "text": {"content": text}}],
            "icon": {"type": "emoji", "emoji": icon},
        }
    }


def table_row(cells: list) -> dict:
    return {
        "object": "block",
        "type": "table_row",
        "table_row": {
            "cells": [[{"type": "text", "text": {"content": str(c)}}] for c in cells]
        }
    }


def table(rows: list, has_header: bool = True) -> dict:
    width = len(rows[0]) if rows else 0
    return {
        "object": "block",
        "type": "table",
        "table": {
            "table_width": width,
            "has_column_header": has_header,
            "has_row_header": False,
            "children": [table_row(row) for row in rows],
        }
    }


def toggle(title: str, children: list = None) -> dict:
    block = {
        "object": "block",
        "type": "toggle",
        "toggle": {
            "rich_text": [{"type": "text", "text": {"content": title}}],
        }
    }
    if children:
        block["toggle"]["children"] = children[:100]
    return block


# ============================================================
# ページ操作API
# ============================================================

def create_meeting_page(meeting_date: str = None, title_suffix: str = "税理士ミーティング") -> dict:
    """税理士ミーティング用のNotionページを自動作成する。

    親ページから権限が自動継承されるため、手動接続は不要。
    """
    parent_id = _get_parent_page_id()

    if meeting_date is None:
        meeting_date = date.today().isoformat()

    page_title = f"{meeting_date} {title_suffix}"

    r = requests.post(
        "https://api.notion.com/v1/pages",
        headers=_headers(),
        json={
            "parent": {"page_id": parent_id},
            "properties": {
                "title": {
                    "title": [{"text": {"content": page_title}}]
                }
            },
        },
    )

    if r.status_code != 200:
        return {"error": r.status_code, "detail": r.json()}

    page = r.json()
    page_id = page["id"]
    url = page.get("url", f"https://notion.so/{page_id.replace('-', '')}")

    config = _load_config()
    company_names = [c["name"] for c in config.get("companies", [])]

    init_blocks = [
        callout(f"{meeting_date} 税理士ミーティング（Claude Code自動作成）", "📋"),
        divider(),
        heading("対象法人・事業", 2),
        bulleted("、".join(company_names) if company_names else "（companies.yamlで設定）"),
        divider(),
        heading("文字起こし全文", 2),
        callout(
            "会議後、文字起こしテキストをこの下に貼り付けてください",
            "👇"
        ),
    ]

    append_result = append_blocks(page_id, init_blocks)
    if append_result.get("error"):
        return {
            "error": "page_created_but_init_failed",
            "page_id": page_id, "url": url,
            "detail": append_result,
        }

    return {"ok": True, "page_id": page_id, "url": url, "title": page_title}


def append_blocks(page_url_or_id: str, blocks: list) -> dict:
    page_id = extract_page_id(page_url_or_id)
    url = f"https://api.notion.com/v1/blocks/{page_id}/children"
    results = []
    for i in range(0, len(blocks), 100):
        chunk = blocks[i:i + 100]
        r = requests.patch(url, headers=_headers(), json={"children": chunk})
        if r.status_code != 200:
            return {"error": r.status_code, "detail": r.json(), "chunk_index": i}
        results.append(r.json())
    return {"ok": True, "chunks_sent": len(results), "total_blocks": len(blocks)}


def insert_blocks_after(page_url_or_id: str, after_block_id: str, blocks: list) -> dict:
    page_id = extract_page_id(page_url_or_id)
    url = f"https://api.notion.com/v1/blocks/{page_id}/children"
    results = []
    current_after = after_block_id
    for i in range(0, len(blocks), 100):
        chunk = blocks[i:i + 100]
        payload = {"children": chunk, "after": current_after}
        r = requests.patch(url, headers=_headers(), json=payload)
        if r.status_code != 200:
            return {"error": r.status_code, "detail": r.json(), "chunk_index": i}
        data = r.json()
        results.append(data)
        inserted = data.get("results", [])
        if inserted:
            current_after = inserted[-1]["id"]
    return {"ok": True, "chunks_sent": len(results), "total_blocks": len(blocks)}


def get_blocks(page_url_or_id: str, include_children: bool = False) -> list:
    page_id = extract_page_id(page_url_or_id)
    blocks = _fetch_block_children(page_id)
    if include_children:
        enriched = []
        for b in blocks:
            enriched.append(b)
            if b.get("has_children") and b.get("type") != "child_page":
                children = _fetch_block_children(b["id"])
                for c in children:
                    c["_parent_block_id"] = b["id"]
                    enriched.append(c)
        return enriched
    return blocks


def delete_blocks(block_ids: list) -> dict:
    results = []
    for bid in block_ids:
        r = requests.delete(
            f"https://api.notion.com/v1/blocks/{bid}",
            headers=_headers(),
        )
        results.append({"id": bid, "status": r.status_code})
    failed = [r for r in results if r["status"] != 200]
    return {"ok": len(failed) == 0, "deleted": len(results) - len(failed), "failed": failed}


def update_block_text(block_id: str, new_text: str, block_type: str = "paragraph") -> dict:
    r = requests.patch(
        f"https://api.notion.com/v1/blocks/{block_id}",
        headers=_headers(),
        json={
            block_type: {
                "rich_text": [{"type": "text", "text": {"content": new_text}}]
            }
        },
    )
    if r.status_code == 200:
        return {"ok": True}
    return {"error": r.status_code, "detail": r.json()}


def create_transcript_toggle(page_url_or_id: str, transcript_text: str,
                              toggle_title: str = "文字起こし全文（クリックで展開）") -> dict:
    page_id = extract_page_id(page_url_or_id)
    chunks = []
    for i in range(0, len(transcript_text), 2000):
        chunks.append(transcript_text[i:i + 2000])

    first_batch = chunks[:100]
    children = [text_block(chunk) for chunk in first_batch]
    toggle_block = toggle(toggle_title, children)

    result = append_blocks(page_id, [toggle_block])
    if result.get("error"):
        return result

    if len(chunks) > 100:
        blocks = get_blocks(page_id)
        toggle_id = None
        for b in reversed(blocks):
            if b.get("type") == "toggle":
                toggle_id = b["id"]
                break
        if toggle_id:
            for i in range(100, len(chunks), 100):
                batch = chunks[i:i + 100]
                extra_children = [text_block(chunk) for chunk in batch]
                extra_result = append_blocks(toggle_id, extra_children)
                if extra_result.get("error"):
                    return extra_result

    return {"ok": True, "chunks": len(chunks), "toggle_title": toggle_title}


def find_block_by_text(page_url_or_id: str, search_text: str) -> str:
    blocks = get_blocks(page_url_or_id)
    for b in blocks:
        btype = b.get("type", "")
        content = b.get(btype, {})
        rich_text = content.get("rich_text", [])
        block_text = "".join(rt.get("plain_text", "") for rt in rich_text)
        if search_text in block_text:
            return b["id"]
    return None


# ============================================================
# 各社データベース読み込み
# ============================================================

def get_companies() -> list:
    config = _load_config()
    return config.get("companies", [])


def get_company(company_id: str) -> dict:
    for c in get_companies():
        if c["id"] == company_id:
            return c
    return None


# ============================================================
# 内部ヘルパー
# ============================================================

def _fetch_block_children(block_id: str) -> list:
    blocks = []
    cursor = None
    while True:
        url = f"https://api.notion.com/v1/blocks/{block_id}/children?page_size=100"
        if cursor:
            url += f"&start_cursor={cursor}"
        r = requests.get(url, headers=_headers())
        if r.status_code != 200:
            break
        data = r.json()
        blocks.extend(data.get("results", []))
        if not data.get("has_more"):
            break
        cursor = data.get("next_cursor")
    return blocks


# ============================================================
# CLI
# ============================================================

if __name__ == "__main__":
    import sys

    usage = """Usage:
  python notion_writer.py test                         # 接続テスト
  python notion_writer.py create [YYYY-MM-DD]          # 会議ページを新規作成
  python notion_writer.py append <page_url> <text>     # テキストを追記
  python notion_writer.py blocks <page_url>            # ブロック一覧を表示
  python notion_writer.py companies                    # 登録法人一覧を表示
"""

    if len(sys.argv) < 2:
        print(usage)
        sys.exit(1)

    cmd = sys.argv[1]

    if cmd == "test":
        try:
            pid = _get_parent_page_id()
            r = requests.get(
                f"https://api.notion.com/v1/pages/{pid}",
                headers=_headers(),
            )
            if r.status_code == 200:
                title_parts = r.json().get("properties", {}).get("title", {}).get("title", [])
                title = "".join(t.get("plain_text", "") for t in title_parts)
                print(f"✅ 接続成功: 「{title}」(ID: {pid})")
            else:
                print(f"❌ 接続失敗 (HTTP {r.status_code}): {r.json().get('message', '')}")
                print("   → Notionで親ページに claude-code integration を接続してください")
        except ValueError as e:
            print(f"❌ 設定エラー: {e}")

    elif cmd == "create":
        meeting_date = sys.argv[2] if len(sys.argv) > 2 else None
        result = create_meeting_page(meeting_date)
        print(json.dumps(result, indent=2, ensure_ascii=False))
        if result.get("ok"):
            print(f"\n✅ ページ作成完了: {result['title']}")
            print(f"   URL: {result['url']}")
        else:
            print(f"\n❌ エラー: {result}")

    elif cmd == "append":
        if len(sys.argv) < 4:
            print("Usage: python notion_writer.py append <page_url> <text>")
            sys.exit(1)
        blocks = [text_block(sys.argv[3])]
        result = append_blocks(sys.argv[2], blocks)
        print(json.dumps(result, indent=2, ensure_ascii=False))

    elif cmd == "blocks":
        if len(sys.argv) < 3:
            print("Usage: python notion_writer.py blocks <page_url>")
            sys.exit(1)
        blocks = get_blocks(sys.argv[2])
        for i, b in enumerate(blocks):
            btype = b.get("type", "")
            content = b.get(btype, {})
            rich_text = content.get("rich_text", [])
            text = "".join(rt.get("plain_text", "") for rt in rich_text)[:80]
            print(f"  [{i}] {btype} ({b['id']}): {text}")

    elif cmd == "companies":
        companies = get_companies()
        print(f"登録法人数: {len(companies)}")
        for c in companies:
            sources = len(c.get("data_sources", []))
            print(f"  - {c['name']} ({c['id']}) — データソース: {sources}件")

    else:
        print(usage)
