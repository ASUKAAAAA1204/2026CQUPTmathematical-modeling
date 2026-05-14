#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""公共外部数据爬取脚本

目标:
- 公开论文页: 通过 Crossref 发现 lightning / remanent magnetization 相关公开论文, 再用 Scrapling 抓页面与 PDF 链接
- 公开天气页: 用 Scrapling 抓日本气象厅 JMA 日值页, 提取逐日温度和湿度

说明:
- 运行环境固定使用 D:\\programss\\p2\\.venv
- 依赖: scrapling, requests
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import time
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urljoin, urlparse

import requests
from scrapling.fetchers import Fetcher


ROOT = Path(r"D:\CQUPT\self_file\competition\gjsschool2026\B题\赛题主要文件夹\问题四\外部数据抓取")
RAW_DIR = ROOT / "raw"
PAPER_HTML_DIR = RAW_DIR / "papers_html"
PAPER_PDF_DIR = RAW_DIR / "papers_pdf"
WEATHER_HTML_DIR = RAW_DIR / "weather_html"
STRUCTURED_DIR = ROOT / "structured"


CROSSREF_QUERIES = [
    "lightning remanent magnetization",
    "lightning magnetic properties volcanic ash",
    "geomagnetic anomaly lightning observatory",
    "fossil of lightning current remanent magnetization",
    "propagation lightning current remanent magnetization",
    "lightning induced magnetic anomaly open access",
    "lightning induced isothermal remanent magnetization",
    "remanent magnetization lightning current",
    "magnetic observatory lightning protection",
    "geomagnetic observatory lightning remanent magnetization",
]

PAPER_TITLE_RE = re.compile(
    r"(lightning|remanent|remanence|volcanic ash|geomagnetic|peak currents|impulse experiments)",
    re.IGNORECASE,
)

SEED_PAPER_URLS = [
    "https://doi.org/10.1038/s41598-019-41265-3",
    "https://doi.org/10.1186/BF03352687",
    "https://doi.org/10.2183/pjab.78.1",
    "https://doi.org/10.1541/ieejpes.133.694",
    "https://doi.org/10.1029/2002GL015207",
    "https://doi.org/10.1093/gji/ggt230",
    "https://doi.org/10.22564/19cisbgf2025.565",
    "https://doi.org/10.1002/(SICI)1520-6416(199806)123:4%3C41::AID-EEJ6%3E3.0.CO;2-O",
    "https://doi.org/10.3133/b1083E",
    "https://doi.org/10.3133/b1203A",
    "https://doi.org/10.1126/science.170.3958.628",
    "https://doi.org/10.1111/J.1365-246X.1980.TB04862.X",
    "https://doi.org/10.1111/J.1365-246X.1982.TB04918.X",
    "https://doi.org/10.1111/J.1365-246X.1989.TB02294.X",
    "https://doi.org/10.1190/1.1442765",
    "https://doi.org/10.1029/jb092ib08p08077",
    "https://doi.org/10.1029/92JB01026",
    "https://doi.org/10.1029/91JB01975",
    "https://doi.org/10.1029/2023GL107105",
    "https://doi.org/10.1038/243027A0",
    "https://doi.org/10.1038/197444A0",
    "https://doi.org/10.1038/197476A0",
    "https://doi.org/10.1038/237274A0",
    "https://doi.org/10.1063/1.4863490",
    "https://doi.org/10.1190/1.1487055",
    "https://doi.org/10.1029/2022JB024151",
    "https://doi.org/10.22541/essoar.169945514.44942099/v1",
    "https://doi.org/10.22541/essoar.176005203.38057263/v1",
    "https://doi.org/10.22541/essoar.177316864.44350310/v1",
]


JMA_TARGETS = [
    {
        "tag": "jma_0402_48",
        "prec_no": 48,
        "block_no": "0402",
        "page": "daily_a1",
        "view": "a1s",
        "station_name": "station_0402_48",
        "years": [2024, 2025],
        "months": list(range(1, 13)),
    },
    {
        "tag": "jma_47638_49",
        "prec_no": 49,
        "block_no": "47638",
        "page": "daily_s1",
        "view": "a1",
        "station_name": "kofu_47638_49",
        "years": [2024, 2025],
        "months": list(range(1, 13)),
    }
]


SESSION = requests.Session()
SESSION.headers.update(
    {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.8,zh-CN;q=0.7,zh;q=0.6",
    }
)


def ensure_dirs() -> None:
    for path in [ROOT, RAW_DIR, PAPER_HTML_DIR, PAPER_PDF_DIR, WEATHER_HTML_DIR, STRUCTURED_DIR]:
        path.mkdir(parents=True, exist_ok=True)


def slugify(value: str, limit: int = 120) -> str:
    value = re.sub(r"[^\w\u4e00-\u9fff]+", "_", value, flags=re.UNICODE)
    value = re.sub(r"_+", "_", value).strip("_")
    if not value:
        value = "item"
    if len(value) > limit:
        value = value[:limit].rstrip("_")
    return value


def file_token(value: str) -> str:
    digest = hashlib.md5(value.encode("utf-8")).hexdigest()[:10]
    return digest


def normalize_text(value: str | None) -> str:
    if not value:
        return ""
    return re.sub(r"\s+", " ", value).strip()


def clean_abstract(value: str | None) -> str:
    if not value:
        return ""
    text = re.sub(r"<[^>]+>", " ", value)
    text = text.replace("\n", " ")
    return normalize_text(text)


def first_nonempty(values: Iterable[str]) -> str:
    for value in values:
        value = normalize_text(value)
        if value:
            return value
    return ""


def to_float(value: str | None) -> float | None:
    if value is None:
        return None
    text = normalize_text(value)
    if not text or text in {"///", "---", "×", "××"}:
        return None
    text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: str | None) -> int | None:
    if value is None:
        return None
    text = normalize_text(value)
    if not text:
        return None
    try:
        return int(float(text))
    except ValueError:
        return None


def crossref_search(query: str, rows: int = 12) -> list[dict[str, Any]]:
    url = "https://api.crossref.org/works"
    params = {
        "query.bibliographic": query,
        "rows": rows,
    }
    resp = SESSION.get(url, params=params, timeout=45)
    resp.raise_for_status()
    data = resp.json()["message"]["items"]
    items: list[dict[str, Any]] = []
    for item in data:
        title = first_nonempty(item.get("title", []))
        doi = normalize_text(item.get("DOI", ""))
        page_url = normalize_text(item.get("URL", ""))
        if not title or not doi:
            continue
        items.append(
            {
                "query": query,
                "title": title,
                "doi": doi,
                "url": page_url or f"https://doi.org/{doi}",
                "publisher": normalize_text(item.get("publisher", "")),
                "container_title": first_nonempty(item.get("container-title", [])),
                "published": (
                    "-".join(str(x) for x in item.get("issued", {}).get("date-parts", [[None]])[0] if x is not None)
                    if item.get("issued")
                    else ""
                ),
                "type": normalize_text(item.get("type", "")),
                "score": item.get("score"),
                "abstract": clean_abstract(item.get("abstract", "")),
                "raw": item,
            }
        )
    return items


def discover_papers(limit: int = 24) -> list[dict[str, Any]]:
    seen: set[str] = set()
    discovered: list[dict[str, Any]] = []
    for query in CROSSREF_QUERIES:
        try:
            items = crossref_search(query, rows=limit)
        except Exception as exc:
            discovered.append({"query": query, "error": str(exc), "source": "crossref"})
            continue
        for item in items:
            doi = item["doi"].lower()
            if doi in seen:
                continue
            title = item.get("title", "")
            if item.get("type") == "reference-entry":
                continue
            if not PAPER_TITLE_RE.search(title):
                continue
            seen.add(doi)
            item["source"] = "crossref"
            discovered.append(item)
    for url in SEED_PAPER_URLS:
        doi = url.replace("https://doi.org/", "").strip().lower()
        if doi in seen:
            continue
        seen.add(doi)
        discovered.append(
            {
                "query": "manual_seed",
                "title": "",
                "doi": doi,
                "url": url,
                "publisher": "",
                "container_title": "",
                "published": "",
                "type": "seed",
                "score": "",
                "abstract": "",
                "source": "seed",
                "raw": {},
            }
        )
    return discovered


def save_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def save_json(path: Path, obj: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def extract_meta(resp) -> dict[str, Any]:
    meta = {
        "title": first_nonempty(
            [
                resp.css('meta[name="citation_title"]::attr(content)').get(),
                resp.css("title::text").get(),
                resp.css('meta[property="og:title"]::attr(content)').get(),
            ]
        ),
        "doi": first_nonempty(
            [
                resp.css('meta[name="citation_doi"]::attr(content)').get(),
                resp.css('meta[name="dc.Identifier"]::attr(content)').get(),
            ]
        ),
        "abstract": clean_abstract(
            first_nonempty(
                [
                    resp.css('meta[name="citation_abstract"]::attr(content)').get(),
                    resp.css('meta[name="description"]::attr(content)').get(),
                ]
            )
        ),
        "journal": first_nonempty(
            [
                resp.css('meta[name="citation_journal_title"]::attr(content)').get(),
                resp.css('meta[name="prism.publicationName"]::attr(content)').get(),
            ]
        ),
        "publication_date": first_nonempty(
            [
                resp.css('meta[name="citation_publication_date"]::attr(content)').get(),
                resp.css('meta[name="citation_date"]::attr(content)').get(),
            ]
        ),
        "pdf_url": first_nonempty(
            [
                resp.css('meta[name="citation_pdf_url"]::attr(content)').get(),
                resp.css('a[href*="pdf"]::attr(href)').get(),
            ]
        ),
        "authors": [normalize_text(x) for x in resp.css('meta[name="citation_author"]::attr(content)').getall() if normalize_text(x)],
    }
    meta["paragraph_count"] = len(resp.css("p").getall())
    meta["word_count_est"] = len(re.findall(r"\w+", clean_abstract(meta["abstract"])))
    return meta


def crawl_paper_page(url: str, index: int, download_pdfs: bool = True) -> dict[str, Any]:
    resp = Fetcher.get(url, timeout=75)
    meta = extract_meta(resp)
    html = getattr(resp, "html_content", "")
    if not isinstance(html, str):
        html = html.decode("utf-8", errors="replace") if isinstance(html, bytes) else str(html)
    token = file_token(url)
    title_slug = slugify(meta.get("title") or urlparse(url).path)
    base_name = f"{title_slug}_{token}"
    html_path = PAPER_HTML_DIR / f"{base_name}.html"
    save_text(html_path, html)

    pdf_saved = ""
    pdf_url = normalize_text(meta.get("pdf_url", ""))
    if pdf_url and download_pdfs:
        pdf_full = urljoin(url, pdf_url)
        pdf_path = PAPER_PDF_DIR / f"{base_name}.pdf"
        if download_pdf(pdf_full, pdf_path):
            pdf_saved = str(pdf_path)

    extracted = {
        "index": index,
        "page_url": url,
        "saved_html": str(html_path),
        "saved_pdf": pdf_saved,
        **meta,
    }
    return extracted


def download_pdf(url: str, path: Path) -> bool:
    try:
        resp = SESSION.get(url, timeout=90, allow_redirects=True)
        if resp.status_code != 200:
            return False
        content = resp.content
        if not content:
            return False
        if content[:4] != b"%PDF" and not url.lower().endswith(".pdf"):
            return False
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(content)
        return True
    except Exception:
        return False


def crawl_weather_month(
    prec_no: int,
    block_no: str,
    year: int,
    month: int,
    tag: str,
    page: str = "daily_a1",
    view: str = "a1s",
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    url = (
        f"https://www.data.jma.go.jp/stats/etrn/view/{page}.php"
        f"?block_no={block_no}&day=&month={month:02d}&prec_no={prec_no}&view={view}&year={year}"
    )
    resp = Fetcher.get(url, timeout=75)
    html = getattr(resp, "html_content", "")
    if not isinstance(html, str):
        html = html.decode("utf-8", errors="replace") if isinstance(html, bytes) else str(html)
    token = file_token(url)
    html_path = WEATHER_HTML_DIR / f"{tag}_{year}_{month:02d}_{token}.html"
    save_text(html_path, html)

    rows = resp.css("tr.mtx")
    records: list[dict[str, Any]] = []
    for row in rows:
        day_text = normalize_text(row.css("td a::text").get())
        if not day_text or not day_text.isdigit():
            continue
        cells = [normalize_text(x) for x in row.css("td::text").getall()]
        day = int(day_text)
        if page == "daily_s1":
            if len(cells) < 18:
                continue
            records.append(
                {
                    "tag": tag,
                    "prec_no": prec_no,
                    "block_no": block_no,
                    "year": year,
                    "month": month,
                    "day": day,
                    "date": f"{year:04d}-{month:02d}-{day:02d}",
                    "pressure_local_hpa": to_float(cells[0]),
                    "pressure_sea_level_hpa": to_float(cells[1]),
                    "pressure_extra_hpa": to_float(cells[2]),
                    "pressure_extra_time": cells[3],
                    "precip_total_mm": to_float(cells[4]),
                    "precip_1h_max_mm": to_float(cells[5]),
                    "precip_1h_max_time": cells[6],
                    "precip_10m_max_mm": to_float(cells[7]),
                    "precip_10m_max_time": cells[8],
                    "temp_mean_c": to_float(cells[9]),
                    "temp_max_c": to_float(cells[10]),
                    "temp_max_time": cells[11],
                    "temp_min_c": to_float(cells[12]),
                    "temp_min_time": cells[13],
                    "vapor_pressure_hpa": to_float(cells[14]),
                    "humidity_mean_pct": to_float(cells[15]),
                    "humidity_min_pct": to_float(cells[16]),
                    "humidity_min_time": cells[17],
                    "page_url": url,
                    "saved_html": str(html_path),
                }
            )
        else:
            if len(cells) < 14:
                continue
            records.append(
                {
                    "tag": tag,
                    "prec_no": prec_no,
                    "block_no": block_no,
                    "year": year,
                    "month": month,
                    "day": day,
                    "date": f"{year:04d}-{month:02d}-{day:02d}",
                    "pressure_local_hpa": None,
                    "pressure_sea_level_hpa": None,
                    "pressure_extra_hpa": None,
                    "pressure_extra_time": "",
                    "precip_total_mm": to_float(cells[0]),
                    "precip_1h_max_mm": to_float(cells[1]),
                    "precip_1h_max_time": cells[2],
                    "precip_10m_max_mm": to_float(cells[3]),
                    "precip_10m_max_time": cells[4],
                    "temp_mean_c": to_float(cells[5]),
                    "temp_max_c": to_float(cells[6]),
                    "temp_max_time": cells[7],
                    "temp_min_c": to_float(cells[8]),
                    "temp_min_time": cells[9],
                    "vapor_pressure_hpa": to_float(cells[10]),
                    "humidity_mean_pct": to_float(cells[11]),
                    "humidity_min_pct": to_float(cells[12]),
                    "humidity_min_time": cells[13],
                    "page_url": url,
                    "saved_html": str(html_path),
                }
            )

    summary = {
        "tag": tag,
        "prec_no": prec_no,
        "block_no": block_no,
        "page": page,
        "view": view,
        "year": year,
        "month": month,
        "page_url": url,
        "saved_html": str(html_path),
        "row_count": len(records),
    }
    return summary, records


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    fieldnames: list[str] = []
    for row in rows:
        for key in row.keys():
            if key not in fieldnames:
                fieldnames.append(key)
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({k: json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v for k, v in row.items()})


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--max-articles-per-query", type=int, default=12)
    parser.add_argument("--skip-pdf", action="store_true")
    parser.add_argument("--skip-papers", action="store_true")
    parser.add_argument("--skip-weather", action="store_true")
    parser.add_argument("--weather-years", type=int, nargs="*", default=[2024, 2025])
    parser.add_argument("--weather-months", type=int, nargs="*", default=list(range(1, 13)))
    args = parser.parse_args()

    ensure_dirs()

    manifest: dict[str, Any] = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "article_queries": CROSSREF_QUERIES,
        "weather_targets": JMA_TARGETS,
        "articles": [],
        "weather_months": [],
    }

    # 1) Paper discovery + landing pages
    article_rows: list[dict[str, Any]] = []
    if not args.skip_papers:
        discovered = discover_papers(limit=args.max_articles_per_query)
        for i, item in enumerate(discovered, 1):
            if item.get("source") not in {"crossref", "seed"}:
                continue
            try:
                page_url = item["url"]
                extracted = crawl_paper_page(page_url, i, download_pdfs=not args.skip_pdf)
                extracted["crossref_query"] = item["query"]
                extracted["crossref_publisher"] = item.get("publisher", "")
                extracted["crossref_container_title"] = item.get("container_title", "")
                extracted["crossref_score"] = item.get("score")
                extracted["crossref_abstract"] = item.get("abstract", "")
                article_rows.append(extracted)
                manifest["articles"].append(
                    {
                        "title": extracted.get("title", ""),
                        "doi": extracted.get("doi", ""),
                        "page_url": extracted.get("page_url", ""),
                        "saved_html": extracted.get("saved_html", ""),
                        "saved_pdf": extracted.get("saved_pdf", ""),
                    }
                )
            except Exception as exc:
                article_rows.append(
                    {
                        "index": i,
                        "page_url": item.get("url", ""),
                        "error": str(exc),
                        "title": item.get("title", ""),
                        "doi": item.get("doi", ""),
                    }
                )

    # 2) JMA weather month pages
    weather_rows: list[dict[str, Any]] = []
    if not args.skip_weather:
        for target in JMA_TARGETS:
            tag = target["tag"]
            prec_no = int(target["prec_no"])
            block_no = str(target["block_no"])
            page = str(target.get("page", "daily_a1"))
            view = str(target.get("view", "a1s"))
            for year in args.weather_years:
                for month in args.weather_months:
                    try:
                        summary, records = crawl_weather_month(prec_no, block_no, year, month, tag, page=page, view=view)
                        manifest["weather_months"].append(summary)
                        weather_rows.extend(records)
                    except Exception as exc:
                        manifest["weather_months"].append(
                            {
                                "tag": tag,
                                "prec_no": prec_no,
                                "block_no": block_no,
                                "page": page,
                                "view": view,
                                "year": year,
                                "month": month,
                                "error": str(exc),
                            }
                        )

    if not args.skip_papers:
        write_csv(STRUCTURED_DIR / "articles_crossref.csv", article_rows)
        save_json(STRUCTURED_DIR / "articles_crossref.json", article_rows)
    if not args.skip_weather:
        write_csv(STRUCTURED_DIR / "weather_jma_daily.csv", weather_rows)
        save_json(STRUCTURED_DIR / "weather_jma_daily.json", weather_rows)
    save_json(STRUCTURED_DIR / "crawl_manifest.json", manifest)

    print(f"articles={len(article_rows)} weather_rows={len(weather_rows)}")
    print(f"output={ROOT}")


if __name__ == "__main__":
    main()
