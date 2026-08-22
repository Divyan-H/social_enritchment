"""
enrich.py — reads your campaign Excel, visits every Instagram/YouTube post
link in the "Post Link" column, scrapes as much as the data source gives us,
and writes ONE output Excel file = your original data + new enr_* columns.

Instagram: if APIFY_API_TOKEN is set (or --apify-token is passed), uses
Apify's Instagram Scraper for both the post AND the owning account's
profile, pulling everything Apify exposes: engagement (likes/comments/video
views), hashtags/mentions/location/sponsorship, and full profile info
(followers, following, bio, verification, business account/category,
external link). Without a token it falls back to a much more limited,
less reliable free scrape (see scraper.py docstring) which only fills
genre/handle/followers/following/description.

YouTube: yt-dlp, unaffected by the Apify token.

Usage:
    python enrich.py INPUT.xlsx
    python enrich.py INPUT.xlsx --out enriched_output.xlsx
    python enrich.py INPUT.xlsx --link-col "Post Link" --platform-col "Platform"
    python enrich.py INPUT.xlsx --sleep 1.5 --max 100

    # Apify token (Instagram). Prefer the env var over the flag - it avoids
    # putting the key in shell history / process args / logs:
    export APIFY_API_TOKEN=apify_api_xxxxxxxx
    python enrich.py INPUT.xlsx
    #   ...or, if you must:
    python enrich.py INPUT.xlsx --apify-token apify_api_xxxxxxxx

    # Only run specific rows (1-indexed, matching the row numbers you'd see
    # in Excel if row 1 is the header - i.e. "2" is the first data row):
    python enrich.py INPUT.xlsx --rows 2,5,10-15
    python enrich.py INPUT.xlsx --rows 3-3          # single row via range syntax also works
"""
import argparse
import os
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd

from scraper import scrape_post, scrape_instagram_urls_batch, detect_platform

# ScrapedPost attribute -> output column name. Single source of truth so the
# "preserve untouched rows" and "write output" logic can't drift apart.
OUTPUT_FIELDS = [
    ("genre", "enr_genre"),
    ("genre_confidence", "enr_genre_confidence"),
    ("handle", "enr_account_handle"),
    ("full_name", "enr_full_name"),
    ("followers", "enr_followers"),
    ("following", "enr_following"),
    ("subscribers", "enr_subscribers"),
    ("posts_count", "enr_posts_count"),
    ("biography", "enr_biography"),
    ("is_verified", "enr_is_verified"),
    ("is_private", "enr_is_private"),
    ("is_business_account", "enr_is_business_account"),
    ("business_category", "enr_business_category"),
    ("external_url", "enr_external_url"),
    ("description", "enr_post_description"),
    ("likes", "enr_likes"),
    ("comments", "enr_comments"),
    ("video_views", "enr_video_views"),
    ("post_type", "enr_post_type"),
    ("post_timestamp", "enr_post_timestamp"),
    ("hashtags", "enr_hashtags"),
    ("mentions", "enr_mentions"),
    ("location", "enr_location"),
    ("is_sponsored", "enr_is_sponsored"),
    ("source", "enr_source"),
]
STATUS_COL = "enr_status"
ALL_ENR_COLS = [col for _, col in OUTPUT_FIELDS] + [STATUS_COL]


def parse_rows_arg(rows_arg: str, n_data_rows: int) -> set:
    """Parse '2,5,10-15' (1-indexed, header=row1) into a 0-indexed set of
    dataframe row positions. Raises ValueError on bad input or out-of-range."""
    selected = set()
    for chunk in rows_arg.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "-" in chunk:
            start_s, end_s = chunk.split("-", 1)
            start, end = int(start_s), int(end_s)
        else:
            start = end = int(chunk)
        if start < 2 or end < start:
            raise ValueError(f"invalid row range '{chunk}' (rows are 1-indexed, row 1 is the header, so the first data row is 2)")
        for excel_row in range(start, end + 1):
            df_idx = excel_row - 2  # excel row 2 -> dataframe index 0
            if df_idx >= n_data_rows:
                raise ValueError(f"row {excel_row} is beyond the data (sheet has {n_data_rows} data rows, i.e. up to excel row {n_data_rows + 1})")
            selected.add(df_idx)
    return selected


def _row_values_from_result(result) -> dict:
    out = {}
    for attr, col in OUTPUT_FIELDS:
        val = getattr(result, attr, None)
        if isinstance(val, list):
            val = ", ".join(str(v) for v in val) if val else None
        elif val == "":
            val = None
        out[col] = val
    out[STATUS_COL] = "ok" if result.ok else f"failed: {result.error}"
    return out


def _empty_row(status: str) -> dict:
    out = {col: None for _, col in OUTPUT_FIELDS}
    out[STATUS_COL] = status
    return out


def main():
    ap = argparse.ArgumentParser(description="Scrape + enrich Instagram/YouTube post links in an Excel sheet.")
    ap.add_argument("input", help="Path to input .xlsx file")
    ap.add_argument("--out", default=None, help="Output .xlsx path (default: <input>_enriched.xlsx)")
    ap.add_argument("--sheet", default=0, help="Sheet name or index to read (default: first sheet)")
    ap.add_argument("--link-col", default="Post Link", help="Column name containing the post URL")
    ap.add_argument("--platform-col", default="Platform", help="Column name containing platform (Instagram/YouTube)")
    ap.add_argument("--sleep", type=float, default=0.0,
                     help="Delay between requests, only used for the free (no-Apify-token) "
                          "fallback path and YouTube rows. Ignored for batched Apify calls, "
                          "which don't need per-row throttling. Default 0.")
    ap.add_argument("--max", type=int, default=None, help="Cap number of rows to scrape (for testing)")
    ap.add_argument("--rows", default=None,
                     help="Only scrape these rows, 1-indexed with row 1 = header "
                          "(e.g. '2,5,10-15'). All other rows are left untouched "
                          "(existing enr_* values preserved if present) and marked "
                          "'skipped (not in --rows selection)' only if no prior value exists.")
    ap.add_argument("--apify-token", default=None,
                     help="Apify API token for Instagram scraping. Prefer setting the "
                          "APIFY_API_TOKEN environment variable instead of this flag - "
                          "env vars don't end up in shell history or process listings.")
    ap.add_argument("--youtube-workers", type=int, default=8,
                     help="How many YouTube rows to scrape in parallel via yt-dlp (default 8).")
    args = ap.parse_args()

    apify_token = args.apify_token or os.environ.get("APIFY_API_TOKEN")
    if not apify_token:
        print("NOTE: no Apify token found (set APIFY_API_TOKEN or pass --apify-token). "
              "Falling back to the free, less-reliable, per-row Instagram scrape - this "
              "will be much slower and less complete than the batched Apify path.")

    df = pd.read_excel(args.input, sheet_name=args.sheet)

    if args.link_col not in df.columns:
        print(f"ERROR: column '{args.link_col}' not found. Available columns: {list(df.columns)}")
        sys.exit(1)

    row_selection = None
    if args.rows:
        try:
            row_selection = parse_rows_arg(args.rows, len(df))
        except ValueError as e:
            print(f"ERROR: --rows problem: {e}")
            sys.exit(1)
        print(f"--rows given: only scraping {len(row_selection)} row(s), leaving the rest untouched.")

    n = len(df) if args.max is None else min(args.max, len(df))

    # Preserve any existing enr_* columns so a --rows re-run doesn't wipe out
    # previously scraped data for rows we're not touching this time.
    existing = {col: (df[col].tolist() if col in df.columns else [None] * len(df)) for col in ALL_ENR_COLS}

    results_by_row = {}  # df index -> dict of output columns

    # ── figure out which rows actually need scraping this run ──────────
    to_scrape = []  # list of (df_idx, url, platform_hint, platform)
    for i in range(len(df)):
        if row_selection is not None and i not in row_selection:
            prior = {col: existing[col][i] for col in ALL_ENR_COLS}
            if pd.isna(prior[STATUS_COL]):
                prior[STATUS_COL] = "skipped (not in --rows selection)"
            results_by_row[i] = prior
            continue
        if i >= n:
            results_by_row[i] = _empty_row("skipped (--max reached)")
            continue

        url = str(df.iloc[i].get(args.link_col, "") or "").strip()
        platform_hint = str(df.iloc[i].get(args.platform_col, "") or "").strip() if args.platform_col in df.columns else ""

        if not url or url.lower() == "nan":
            results_by_row[i] = _empty_row("no link")
            continue

        platform = detect_platform(url, platform_hint)
        to_scrape.append((i, url, platform_hint, platform))

    ig_rows = [(i, url) for i, url, _, platform in to_scrape if platform == "instagram"]
    yt_rows = [(i, url, ph) for i, url, ph, platform in to_scrape if platform == "youtube"]
    other_rows = [(i, url) for i, url, _, platform in to_scrape if platform not in ("instagram", "youtube")]

    print(f"Scraping {len(to_scrape)} row(s): {len(ig_rows)} Instagram, {len(yt_rows)} YouTube, "
          f"{len(other_rows)} unrecognized.")

    # ── Instagram: one batched pass (few Apify calls instead of one per row) ──
    if ig_rows:
        ig_urls = [u for _, u in ig_rows]
        if apify_token:
            print(f"Batch-scraping {len(set(ig_urls))} unique Instagram URL(s) via Apify...")
        scraped_by_url = scrape_instagram_urls_batch(ig_urls, apify_token) if apify_token else {
            u: scrape_post(u, "instagram") for u in dict.fromkeys(ig_urls)
        }
        for i, url in ig_rows:
            result = scraped_by_url.get(url)
            results_by_row[i] = _row_values_from_result(result) if result else _empty_row("failed: no result")

    # ── YouTube: parallelize with a small thread pool (yt-dlp is network-bound) ──
    if yt_rows:
        print(f"Scraping {len(yt_rows)} YouTube row(s) with {args.youtube_workers} parallel workers...")
        with ThreadPoolExecutor(max_workers=args.youtube_workers) as ex:
            future_to_row = {
                ex.submit(scrape_post, url, ph): i for i, url, ph in yt_rows
            }
            for fut in as_completed(future_to_row):
                i = future_to_row[fut]
                try:
                    result = fut.result()
                except Exception as e:
                    results_by_row[i] = _empty_row(f"failed: {type(e).__name__}: {e}")
                    continue
                results_by_row[i] = _row_values_from_result(result)

    # ── anything else (unrecognized platform) ──
    for i, url in other_rows:
        result = scrape_post(url, "")
        results_by_row[i] = _row_values_from_result(result)

    for col in ALL_ENR_COLS:
        df[col] = [results_by_row[i][col] for i in range(len(df))]

    out_path = args.out or args.input.rsplit(".", 1)[0] + "_enriched.xlsx"
    df.to_excel(out_path, index=False)
    n_scraped = len(to_scrape)
    print(f"\nDone. Wrote {out_path}  ({n_scraped} rows scraped, {len(df) - n_scraped} skipped/untouched)")


if __name__ == "__main__":
    main()
