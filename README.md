# Social Post Enrichment (Instagram + YouTube)

Reads your campaign Excel, visits every post link, and adds these columns:

| Column | Meaning |
|---|---|
| `enr_genre` | Content genre/category the creator is posting (comedy, educational, lifestyle, food, travel, fashion_beauty, fitness_health, music_dance, tech, gaming, promotional, review, etc.) — detected from the caption/description |
| `enr_account_handle` | The username/channel that posted it |
| `enr_followers` | Instagram followers (blank for YouTube) |
| `enr_following` | Instagram following (blank for YouTube) |
| `enr_subscribers` | YouTube subscribers (blank for Instagram) |
| `enr_post_description` | The post's caption / video description |
| `enr_status` | `ok`, `no link`, or `failed: <reason>` |

Nothing else is touched — your original columns are kept exactly as they are.

## Install
```bash
pip install -r requirements.txt
```

## Run
```bash
python enrich.py YOUR_FILE.xlsx
```
Produces `YOUR_FILE_enriched.xlsx` in the same folder — one file, with every
original column plus the `enr_*` enrichment columns above.

Options:
```bash
python enrich.py YOUR_FILE.xlsx --out result.xlsx     # custom output name
python enrich.py YOUR_FILE.xlsx --link-col "Post Link" --platform-col "Platform"
python enrich.py YOUR_FILE.xlsx --sleep 1.5            # slower = safer vs rate limits
python enrich.py YOUR_FILE.xlsx --max 20               # test on first 20 rows only
```

## How it works
- **YouTube** — `yt-dlp` pulls the description + subscriber count in one call.
- **Instagram** — `yt-dlp` is tried first; if a post/handle info is missing,
  it falls back to Instagram's public embed page for the caption, then a
  second request to `instagram.com/<username>/` to read followers/following
  off the public page. Each account is only looked up **once per run**, so
  a creator with many rows in your sheet doesn't cost extra requests.
- **Genre** — a fast offline keyword classifier reads the caption + handle.
  No LLM, no API key, no internet call needed for this step.

## Notes / limitations
- Private accounts, deleted posts, or heavy rate-limiting on Instagram's side
  will show up as `enr_status = failed: ...` for that row — it never stops
  the rest of the run.
- Only Instagram and YouTube links are supported (per your data).
- This needs to be run somewhere with real internet access to Instagram/YouTube.
