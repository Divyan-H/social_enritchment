"""
enrich.py — reads your campaign Excel, visits every Instagram/YouTube post
link in the "Post Link" column, scrapes:

    - genre (content category the creator is posting)
    - following count       (Instagram)
    - followers count       (Instagram)
    - subscribers count     (YouTube)
    - description of the post

...and writes ONE output Excel file = your original data + these new columns.

Usage:
    python enrich.py INPUT.xlsx
    python enrich.py INPUT.xlsx --out enriched_output.xlsx
    python enrich.py INPUT.xlsx --link-col "Post Link" --platform-col "Platform"
    python enrich.py INPUT.xlsx --sleep 1.5 --max 100

    # Only run specific rows (1-indexed, matching the row numbers you'd see
    # in Excel if row 1 is the header - i.e. "2" is the first data row):
    python enrich.py INPUT.xlsx --rows 2,5,10-15
    python enrich.py INPUT.xlsx --rows 3-3          # single row via range syntax also works
"""
import argparse
import sys
import time

import pandas as pd

from scraper import scrape_post


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


def main():
    ap = argparse.ArgumentParser(description="Scrape + enrich Instagram/YouTube post links in an Excel sheet.")
    ap.add_argument("input", help="Path to input .xlsx file")
    ap.add_argument("--out", default=None, help="Output .xlsx path (default: <input>_enriched.xlsx)")
    ap.add_argument("--sheet", default=0, help="Sheet name or index to read (default: first sheet)")
    ap.add_argument("--link-col", default="Post Link", help="Column name containing the post URL")
    ap.add_argument("--platform-col", default="Platform", help="Column name containing platform (Instagram/YouTube)")
    ap.add_argument("--sleep", type=float, default=1.0, help="Delay between requests in seconds (politeness)")
    ap.add_argument("--max", type=int, default=None, help="Cap number of rows to scrape (for testing)")
    ap.add_argument("--rows", default=None,
                     help="Only scrape these rows, 1-indexed with row 1 = header "
                          "(e.g. '2,5,10-15'). All other rows are left untouched "
                          "(existing enr_* values preserved if present) and marked "
                          "'skipped (not in --rows selection)' only if no prior value exists.")
    args = ap.parse_args()

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

    # Preserve any existing enr_* columns so a --rows re-run doesn't wipe
    # out previously scraped data for rows we're not touching this time.
    enr_cols = {
        "enr_genre": "genres", "enr_account_handle": "handles",
        "enr_followers": "followers", "enr_following": "followings",
        "enr_subscribers": "subscribers", "enr_post_description": "descriptions",
        "enr_status": "statuses",
    }
    existing = {col: (df[col].tolist() if col in df.columns else [None] * len(df)) for col in enr_cols}

    genres, followings, followers, subscribers, descriptions = [], [], [], [], []
    statuses, handles = [], []

    for i in range(len(df)):
        if row_selection is not None and i not in row_selection:
            # not selected this run - keep whatever was there before
            genres.append(existing["enr_genre"][i])
            handles.append(existing["enr_account_handle"][i])
            followers.append(existing["enr_followers"][i])
            followings.append(existing["enr_following"][i])
            subscribers.append(existing["enr_subscribers"][i])
            descriptions.append(existing["enr_post_description"][i])
            prior_status = existing["enr_status"][i]
            statuses.append(prior_status if pd.notna(prior_status) else "skipped (not in --rows selection)")
            continue

        if i >= n:
            genres.append(None); followings.append(None); followers.append(None)
            subscribers.append(None); descriptions.append(None)
            statuses.append("skipped (--max reached)"); handles.append(None)
            continue

        url = str(df.iloc[i].get(args.link_col, "") or "").strip()
        platform_hint = str(df.iloc[i].get(args.platform_col, "") or "").strip() if args.platform_col in df.columns else ""

        if not url or url.lower() == "nan":
            genres.append(None); followings.append(None); followers.append(None)
            subscribers.append(None); descriptions.append(None)
            statuses.append("no link"); handles.append(None)
            continue

        excel_row = i + 2
        print(f"[row {excel_row}, {i+1}/{len(df)}] scraping: {url}")
        result = scrape_post(url, platform_hint)

        genres.append(result.genre or None)
        followings.append(result.following)
        followers.append(result.followers)
        subscribers.append(result.subscribers)
        descriptions.append(result.description or None)
        handles.append(result.handle or None)
        statuses.append("ok" if result.ok else f"failed: {result.error}")

        time.sleep(args.sleep)

    df["enr_genre"] = genres
    df["enr_account_handle"] = handles
    df["enr_followers"] = followers
    df["enr_following"] = followings
    df["enr_subscribers"] = subscribers
    df["enr_post_description"] = descriptions
    df["enr_status"] = statuses

    out_path = args.out or args.input.rsplit(".", 1)[0] + "_enriched.xlsx"
    df.to_excel(out_path, index=False)
    n_scraped = len(row_selection) if row_selection is not None else n
    print(f"\nDone. Wrote {out_path}  ({n_scraped} rows scraped, {len(df) - n_scraped} skipped/untouched)")


if __name__ == "__main__":
    main()
