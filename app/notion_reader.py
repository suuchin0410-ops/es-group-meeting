"""税理士ミーティング用 Notion読み取りユーティリティ

Notionページのテキストコンテンツを取得する。
"""
import requests
import re
from pathlib import Path

TOKEN_FILE = Path(__file__).parent / "config" / ".notion_token"
NOTION_VERSION = "2022-06-28"


def _get_token():
    return TOKEN_FILE.read_text().strip()


def _headers():
    return {
        "Authorization": f"Bearer {_get_token()}",
        "Content-Type": "application/json",
        "Notion-Version": NOTION_VERSION,
    }


def _extract_page_id(url_or_id: str) -> str:
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


def _get_blocks(block_id: str) -> list:
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


def _extract_text(rich_text_array: list) -> str:
    return "".join(rt.get("plain_text", "") for rt in rich_text_array)


def _get_page_title(page_id: str) -> str:
    r = requests.get(
        f"https://api.notion.com/v1/pages/{page_id}",
        headers=_headers(),
    )
    if r.status_code != 200:
        return ""
    props = r.json().get("properties", {})
    title_prop = props.get("title", {})
    title_arr = title_prop.get("title", [])
    return _extract_text(title_arr)


def fetch_page(url_or_id: str) -> dict:
    """Notionページのテキストコンテンツを取得する。

    Returns:
        {"title": "...", "page_id": "...", "text": "...", "block_count": N}
    """
    page_id = _extract_page_id(url_or_id)
    title = _get_page_title(page_id)
    blocks = _get_blocks(page_id)

    lines = []
    for b in blocks:
        btype = b.get("type", "")
        content = b.get(btype, {})
        rich_text = content.get("rich_text", [])
        text = _extract_text(rich_text)
        if text:
            lines.append(text)

        if b.get("has_children") and btype != "child_page":
            children = _get_blocks(b["id"])
            for child in children:
                ctype = child.get("type", "")
                ccontent = child.get(ctype, {})
                crich = ccontent.get("rich_text", [])
                ctext = _extract_text(crich)
                if ctext:
                    lines.append(ctext)

    return {
        "title": title,
        "page_id": page_id,
        "text": "\n".join(lines),
        "block_count": len(blocks),
    }


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python notion_reader.py <page_url_or_id>")
        sys.exit(1)
    result = fetch_page(sys.argv[1])
    print(f"Title: {result['title']}")
    print(f"Blocks: {result['block_count']}")
    print(f"---\n{result['text']}")
