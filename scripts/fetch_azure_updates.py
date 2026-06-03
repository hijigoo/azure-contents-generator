"""Azure 최신 정보를 수집해 JSON으로 저장한다.

기본 소스:
- Azure Updates RSS (https://azurecomcdn.azureedge.net/en-us/updates/feed/)
- Microsoft Foundry / AI 관련 블로그 RSS
필요에 따라 소스 URL 목록을 환경 변수 AZURE_FEEDS 로 덮어쓸 수 있다.
"""
from __future__ import annotations

import argparse
import json
import os
from datetime import datetime, timezone
from pathlib import Path

import feedparser

DEFAULT_FEEDS = [
    "https://www.microsoft.com/releasecommunications/api/v2/azure/rss",
    "https://techcommunity.microsoft.com/plugins/custom/microsoft/o365/custom-blog-rss?tid=AI-AzureAIServicesBlog",
]


def fetch(feeds: list[str], limit_per_feed: int = 20) -> list[dict]:
    items: list[dict] = []
    for url in feeds:
        parsed = feedparser.parse(url)
        for entry in parsed.entries[:limit_per_feed]:
            items.append(
                {
                    "source": url,
                    "title": entry.get("title", ""),
                    "link": entry.get("link", ""),
                    "published": entry.get("published", ""),
                    "summary": entry.get("summary", ""),
                }
            )
    return items


def main() -> None:
    parser = argparse.ArgumentParser(description="Azure 최신 정보 수집")
    parser.add_argument("--out", required=True, help="결과 JSON 출력 경로")
    parser.add_argument("--limit", type=int, default=20, help="피드당 최대 항목 수")
    args = parser.parse_args()

    feeds_env = os.environ.get("AZURE_FEEDS")
    feeds = [u.strip() for u in feeds_env.split(",")] if feeds_env else DEFAULT_FEEDS

    payload = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "feeds": feeds,
        "items": fetch(feeds, args.limit),
    }

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[fetch_azure_updates] {len(payload['items'])}건을 {out_path} 에 저장")


if __name__ == "__main__":
    main()
