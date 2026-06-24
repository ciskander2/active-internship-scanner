import argparse
import base64
import datetime as dt
import html
import json
import os
import re
import shutil
import time
import urllib.parse
import urllib.request
from urllib.error import HTTPError, URLError

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


PROVIDERS = [
    ("greenhouse", "https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=true"),
    ("lever", "https://api.lever.co/v0/postings/{slug}?mode=json"),
    ("ashby", "https://api.ashbyhq.com/posting-api/job-board/{slug}"),
]

SLUGS = {
    "Microsoft": ["microsoft"],
    "Bloomberg": ["bloomberg"],
    "Meta": ["meta"],
    "Figma": ["figma"],
    "Dropbox": ["dropbox"],
    "Snowflake": ["snowflake"],
    "Adobe": ["adobe"],
    "Intuit": ["intuit"],
    "Airbnb": ["airbnb"],
    "Spotify": ["spotify"],
    "Uber": ["uber"],
    "Salesforce": ["salesforce"],
    "Lyft": ["lyft"],
    "Amazon": ["amazon"],
    "MongoDb": ["mongodb"],
    "Coinbase": ["coinbase"],
    "Atlassian": ["atlassian"],
    "Shopify": ["shopify"],
    "Pinterest": ["pinterest"],
    "Point 72": ["point72"],
    "SeatGeek": ["seatgeek"],
    "TikTok": ["bytedance", "tiktok"],
    "Databricks": ["databricks"],
    "Duolingo": ["duolingo"],
    "Cloudflare": ["cloudflare"],
    "Tesla": ["tesla"],
    "SoFi": ["sofi"],
    "Reddit": ["reddit"],
    "Two Sigma": ["twosigma"],
    "Rivian": ["rivian"],
    "Millenium": ["millennium"],
    "Citadel": ["citadel", "citadel-securities"],
    "SpaceX": ["spacex"],
    "Gusto": ["gusto"],
    "Snap Inc": ["snapchat", "snap"],
    "Zendesk": ["zendesk"],
    "Netflix": ["netflix"],
    "Precisely": ["precisely"],
    "Confluent": ["confluent"],
    "Robinhood": ["robinhood"],
    "NVIDIA": ["nvidia"],
    "Palantir": ["palantir"],
    "Anduril Industries": ["andurilindustries", "anduril"],
    "Stripe": ["stripe"],
    "Ramp": ["ramp"],
    "ServiceNow": ["servicenow"],
    "Discord": ["discord"],
    "Notion": ["notion"],
    "xAI": ["xai"],
    "Mercury": ["mercury"],
    "DoorDash": ["doordash"],
    "Google": ["google"],
    "Volirdge": ["voloridgeinvestmentmanagement", "voloridge"],
    "Susquehanna Associates": ["sig", "susquehanna"],
    "Bridgewater Associates": ["bridgewater"],
    "DE Shaw": ["deshaw", "the-d-e-shaw-group", "d-e-shaw"],
    "Hudson River Trading": ["hudsonrivertrading", "hrt"],
    "Akuna Capital": ["akunacapital", "akuna"],
    "IMC Trading": ["imc", "imc-trading"],
    "Optiver": ["optiver"],
    "Jump Trading": ["jumptrading", "jump-trading"],
    "Tower Research Capital": ["towerresearchcapital", "tower-research", "towerresearch"],
    "Five Rings": ["fiverings", "five-rings"],
    "Belvedere Trading": ["belvederetrading", "belvedere-trading"],
    "Old Mission": ["oldmission", "old-mission"],
    "Flow Trader": ["flowtraders", "flow-traders"],
    "Geneva Trading": ["genevatrading", "geneva-trading"],
    "XR Trading": ["xrtrading", "xr-trading"],
    "AQR Capital Management": ["aqr", "aqr-capital-management"],
    "Man Group": ["mangroup", "man-group"],
    "Balyasny Asset Management": ["balyasny", "bamfunds"],
    "Schonfeld Strategic Advisors": ["schonfeld", "schonfeldstrategicadvisors"],
    "Datadog": ["datadog"],
    "Jane Street": ["janestreet", "jane-street"],
}

# Search-confirmed official postings that do not expose a simple ATS API.
# Keep this list small and audit-friendly; the live ATS scan remains the primary source.
MANUAL_OFFICIAL_POSTINGS = {
    "Tesla": [
        {
            "title": "Internship, Data Engineer, Fleet Analytics (Fall 2026)",
            "location": "",
            "url": "https://www.tesla.com/careers/search/job/internship-data-engineer-fleet-analytics-fall-2026-271696",
        },
        {
            "title": "Internship, Data Engineer, Fleet Data, Self Driving (Fall 2026)",
            "location": "",
            "url": "https://www.tesla.com/careers/search/job/269828",
        },
        {
            "title": "Internship, Software Engineer, AI Data Infrastructure (Fall 2026)",
            "location": "",
            "url": "https://www.tesla.com/careers/search/job/internship-software-engineer-ai-data-infrastructure-fall-2026-269829",
        },
        {
            "title": "AI Engineering Intern, Summer 2026",
            "location": "",
            "url": "https://www.tesla.com/careers/search/job/ai-engineering-intern-summer-2026-259784",
        },
        {
            "title": "Internship, Software Engineer, Service Engineering (Summer 2026)",
            "location": "",
            "url": "https://www.tesla.com/careers/search/job/internship-software-engineer-service-engineering-summer-2026-259221",
        },
    ],
    "Amazon": [
        {
            "title": "2027 Amazon Operations Finance Rotational Program Summer Internship",
            "location": "Seattle, WA / Arlington, VA",
            "url": "https://www.amazon.jobs/en/jobs/10435673/2027-amazon-operations-finance-rotational-program-summer-internship",
        }
    ],
    "Optiver": [
        {
            "title": "Expressions of Interest - Quantitative Research Internship, PhD (Summer 2027 - Shanghai)",
            "location": "Shanghai",
            "url": "https://www.optiver.com/working-at-optiver/career-opportunities/page/2/?level=internship&numberposts=10",
        }
    ],
    "Susquehanna Associates": [
        {
            "title": "Quantitative Trader Internship: Summer 2027",
            "location": "",
            "url": "https://careers.sig.com/quantitative-systematic-trading-quantitative-research/jobs/10717?lang=en-us",
        },
        {
            "title": "Equity Analyst Internship: Summer 2027",
            "location": "New York",
            "url": "https://careers.sig.com/new-york/jobs/10573?lang=en-us",
        },
        {
            "title": "Operations Internship: Summer 2027",
            "location": "Bala Cynwyd, PA",
            "url": "https://careers.sig.com/jobs/10916?lang=en-us",
        },
        {
            "title": "ETF Sales Internship: Summer 2027",
            "location": "Bala Cynwyd, PA",
            "url": "https://careers.sig.com/jobs/10944?lang=en-us",
        },
    ],
    "DE Shaw": [
        {
            "title": "Software Developer Intern (New York) - Summer 2027",
            "location": "New York",
            "url": "https://www.deshaw.com/careers/software-developer-intern-new-york-summer-2027-5894",
        },
        {
            "title": "Trader/Analyst Intern (London) - Summer 2027",
            "location": "London",
            "url": "https://www.deshaw.com/careers/trader-analyst-intern-london-summer-2027-5862",
        },
    ],
}


def normalize_company(company):
    return re.sub(r"[^a-z0-9]", "", company.lower().replace("&", "and"))


def variants(company):
    values = list(SLUGS.get(company, []))
    base = normalize_company(company)
    values.extend(
        [
            base,
            base.replace("inc", ""),
            base.replace("industries", ""),
            base.replace("capitalmanagement", ""),
            base.replace("trading", ""),
        ]
    )
    out = []
    for value in values:
        value = value.strip("-")
        if value and value not in out:
            out.append(value)
    return out[:8]


def fetch_json(url, timeout=15):
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0", "Accept": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8", "ignore"))


def text_of(value):
    if isinstance(value, dict):
        return " ".join(text_of(v) for v in value.values())
    if isinstance(value, list):
        return " ".join(text_of(v) for v in value)
    return "" if value is None else str(value)


def extract_jobs(provider, data):
    jobs = []
    if provider == "greenhouse":
        for job in data.get("jobs", []):
            locations = ", ".join(
                location.get("name", "") for location in job.get("locations", [])
            )
            if not locations:
                locations = (job.get("location") or {}).get("name", "")
            jobs.append(
                {
                    "title": job.get("title", ""),
                    "url": job.get("absolute_url", ""),
                    "location": locations,
                    "body": text_of(job),
                }
            )
    elif provider == "lever":
        for job in data:
            categories = job.get("categories", {}) or {}
            jobs.append(
                {
                    "title": job.get("text", ""),
                    "url": job.get("hostedUrl") or job.get("applyUrl") or "",
                    "location": categories.get("location", ""),
                    "body": text_of(job),
                }
            )
    elif provider == "ashby":
        for job in data.get("jobs", []):
            jobs.append(
                {
                    "title": job.get("title", ""),
                    "url": job.get("jobUrl") or job.get("applyUrl") or "",
                    "location": job.get("locationName", ""),
                    "body": text_of(job),
                }
            )
    return jobs


def is_active_internship(job):
    title = (job.get("title") or "").lower()
    internship = r"\binterns?\b|\binternships?\b|\bco-?op\b|\bsummer analyst\b|\bsummer associate\b"
    return bool(re.search(internship, title))


def scan_company(company, delay_seconds=0.1):
    checked = []
    board = None
    all_jobs = []
    for slug in variants(company):
        for provider, template in PROVIDERS:
            url = template.format(slug=slug)
            checked.append(url)
            try:
                data = fetch_json(url)
                jobs = extract_jobs(provider, data)
            except (HTTPError, URLError, TimeoutError, json.JSONDecodeError, OSError):
                jobs = []
            if jobs:
                board = {"provider": provider, "slug": slug, "url": url}
                all_jobs = jobs
                break
        if all_jobs:
            break
    time.sleep(delay_seconds)
    matches = []
    for job in all_jobs:
        if is_active_internship(job):
            if not any(existing.get("url") == job.get("url") for existing in matches):
                matches.append(job)
    for job in MANUAL_OFFICIAL_POSTINGS.get(company, []):
        if is_active_internship(job):
            if not any(existing.get("url") == job.get("url") for existing in matches):
                matches.append(job)
    return {
        "company": company,
        "board": board,
        "job_count": len(all_jobs),
        "matches": matches,
        "checked": checked,
    }


def read_companies(workbook_path, sheet_name):
    workbook = openpyxl.load_workbook(workbook_path, data_only=True)
    sheet = workbook[sheet_name]
    companies = []
    for row in range(2, sheet.max_row + 1):
        value = sheet.cell(row, 1).value
        if value and str(value).strip():
            companies.append((row, str(value).strip()))
    return companies


def style_header(cell):
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.font = Font(color="FFFFFF", bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(bottom=Side(style="thin", color="D9E2F3"))


def write_workbook(input_path, output_path, sheet_name, results, scan_date):
    shutil.copy2(input_path, output_path)
    workbook = openpyxl.load_workbook(output_path)
    sheet = workbook[sheet_name]
    by_company = {result["company"]: result for result in results}

    start_col = 9
    headers = [
        "Active Internship Found?",
        "Internship Role(s) Found",
        "Location(s)",
        "Official Source URL(s)",
        "Scan Notes",
        "Last Checked",
    ]
    for index, header in enumerate(headers, start_col):
        cell = sheet.cell(1, index, header)
        style_header(cell)

    for row in range(2, sheet.max_row + 1):
        value = sheet.cell(row, 1).value
        if not value or not str(value).strip():
            continue
        company = str(value).strip()
        result = by_company.get(company, {})
        matches = result.get("matches", [])
        if matches:
            status = "Yes"
            fill = "C6EFCE"
            roles = "; ".join(job.get("title", "") for job in matches[:10])
            locations = "; ".join(
                dict.fromkeys(job.get("location", "") for job in matches[:10] if job.get("location"))
            )
            urls = "; ".join(job.get("url", "") for job in matches[:10] if job.get("url"))
            note = f"Confirmed active official internship posting(s) found. {len(matches)} match(es)."
        else:
            status = "No confirmed active posting found"
            fill = "FFF2CC"
            roles = ""
            locations = ""
            urls = ""
            board = result.get("board")
            if board:
                urls = board.get("url", "")
                note = (
                    f"Live {board.get('provider')} board scanned "
                    f"({result.get('job_count', 0)} jobs); no job title matched internship/co-op/summer analyst."
                )
            else:
                note = (
                    "Could not verify via Greenhouse/Lever/Ashby automated board scan; "
                    "no confirmed official internship posting found."
                )
        values = [status, roles, locations, urls, note, scan_date]
        for col, output in enumerate(values, start_col):
            cell = sheet.cell(row, col, output)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if col == start_col:
                cell.fill = PatternFill("solid", fgColor=fill)

    summary_name = "Internship Scan Summary"
    if summary_name in workbook.sheetnames:
        del workbook[summary_name]
    summary = workbook.create_sheet(summary_name, 0)
    summary.append(
        ["Company", "Active Internship Found?", "Internship Role(s) Found", "Official Source URL(s)", "Scan Notes"]
    )
    for cell in summary[1]:
        style_header(cell)

    for row in range(2, sheet.max_row + 1):
        company = sheet.cell(row, 1).value
        if company and str(company).strip():
            summary.append(
                [
                    company,
                    sheet.cell(row, start_col).value,
                    sheet.cell(row, start_col + 1).value,
                    sheet.cell(row, start_col + 3).value,
                    sheet.cell(row, start_col + 4).value,
                ]
            )

    for row in summary.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        if row[1].value == "Yes":
            row[1].fill = PatternFill("solid", fgColor="C6EFCE")
        else:
            row[1].fill = PatternFill("solid", fgColor="FFF2CC")

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:N{sheet.max_row}"
    for col, width in [(9, 24), (10, 70), (11, 45), (12, 85), (13, 85), (14, 16)]:
        sheet.column_dimensions[get_column_letter(col)].width = width

    summary.freeze_panes = "A2"
    summary.auto_filter.ref = f"A1:E{summary.max_row}"
    for col, width in [(1, 28), (2, 24), (3, 70), (4, 85), (5, 85)]:
        summary.column_dimensions[get_column_letter(col)].width = width

    workbook.save(output_path)


def main():
    parser = argparse.ArgumentParser(description="Scan a job tracker for active internships.")
    parser.add_argument("--input", required=True, help="Path to the source .xlsx tracker.")
    parser.add_argument("--output-dir", required=True, help="Directory for timestamped scan outputs.")
    parser.add_argument("--sheet", default="Tracking Template", help="Worksheet containing company names in column A.")
    parser.add_argument("--delay", type=float, default=0.1, help="Delay between company scans.")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M")
    scan_date = dt.date.today().isoformat()
    output_path = os.path.join(args.output_dir, f"active_internship_scan_{stamp}.xlsx")
    raw_json_path = os.path.join(args.output_dir, f"active_internship_scan_{stamp}.json")

    companies = read_companies(args.input, args.sheet)
    results = []
    for index, (_, company) in enumerate(companies, 1):
        result = scan_company(company, args.delay)
        results.append(result)
        print(
            f"{index:02d}/{len(companies)} {company}: "
            f"{len(result['matches'])} internship match(es)"
        )

    write_workbook(args.input, output_path, args.sheet, results, scan_date)
    with open(raw_json_path, "w", encoding="utf-8") as handle:
        json.dump(results, handle, indent=2)

    yes_count = sum(1 for result in results if result.get("matches"))
    print("")
    print(f"Companies scanned: {len(results)}")
    print(f"Companies with confirmed internship postings: {yes_count}")
    print(f"Workbook: {output_path}")
    print(f"Raw scan JSON: {raw_json_path}")


if __name__ == "__main__":
    main()
