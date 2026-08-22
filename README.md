# Social Data Enrichment (Instagram & YouTube)

A high-performance pipeline for enriching campaign and influencer Excel spreadsheets with post engagement, creator profile statistics, and content genre classification for **Instagram** and **YouTube**.

---

## 🚀 Overview & Pipelines

This repository provides two specialized enrichment pipelines depending on the format of your input data:

| Pipeline Script | Best Used When | What It Enriches |
| :--- | :--- | :--- |
| **`enrich.py`** | Your spreadsheet contains **Post / Video URLs** in a dedicated column (e.g., `"Post Link"`). | Post metrics (likes, comments, views, caption, tags) + Owner profile data (followers, following, bio, verification, etc.) + Genre. |
| **`enrich_handles.py`** | Your spreadsheet contains **Hyperlinks, Profile Handles, or mixed Post Links** in a column (e.g., `"Social Handle"`). | Automatically resolves embedded Excel hyperlinks, handles creator handles/URLs, skips noise/headers (e.g., `'2022-23'`), and pulls full account & post metrics. |

---

## 📊 Enriched Columns Schema

Both pipelines output identical, standardized `enr_*` columns while preserving all your original sheet columns:

| Output Column | Description |
| :--- | :--- |
| `enr_genre` | Classified content genre (e.g., *comedy, fashion_beauty, food, tech, lifestyle, educational, fitness_health, review, promotional*) |
| `enr_genre_confidence` | Confidence score (0.0 to 1.0) of the genre classification |
| `enr_account_handle` | Instagram account username or YouTube channel handle |
| `enr_full_name` | Creator / Channel display name |
| `enr_followers` | Instagram follower count |
| `enr_following` | Instagram following count |
| `enr_subscribers` | YouTube subscriber count |
| `enr_posts_count` | Total posts / video count |
| `enr_biography` | Instagram profile bio or YouTube channel description |
| `enr_is_verified` | Verification badge status (`True` / `False`) |
| `enr_is_private` | Account privacy status (`True` / `False`) |
| `enr_is_business_account` | Whether the account is a registered business profile |
| `enr_business_category` | Instagram business/creator category |
| `enr_external_url` | Bio link / website URL(s) |
| `enr_post_description` | Post caption or video description |
| `enr_likes` | Post likes count |
| `enr_comments` | Post comments count |
| `enr_video_views` | Video / Reel view count |
| `enr_post_type` | Post format (`Image`, `Video`, `Sidecar`, `Reel`, etc.) |
| `enr_post_timestamp` | Publication timestamp |
| `enr_hashtags` | Extracted hashtags |
| `enr_mentions` | Extracted `@mentions` |
| `enr_location` | Geotagged location name |
| `enr_is_sponsored` | Sponsored / paid partnership flag |
| `enr_source` | Scraping provider used (`apify`, `ytdlp`, or `free_fallback`) |
| `enr_status` | Status of the row (`ok`, `no link`, `skipped`, or `failed: <reason>`) |

---

## ⚙️ Installation

1. **Clone the repository**:
   ```bash
   git clone <repository_url>
   cd social_enrichment
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Set your Apify API Token** (recommended for full metrics & fast batched processing):
   - **Linux / macOS**:
     ```bash
     export APIFY_API_TOKEN="apify_api_xxxxxxxxxxxxxxxxxxxxxxxx"
     ```
   - **Windows (PowerShell)**:
     ```powershell
     $env:APIFY_API_TOKEN="apify_api_xxxxxxxxxxxxxxxxxxxxxxxx"
     ```
   - **Windows (CMD)**:
     ```cmd
     set APIFY_API_TOKEN=apify_api_xxxxxxxxxxxxxxxxxxxxxxxx
     ```

---

## 🏃 Usage Examples

### 1. Enriching Hyperlinks & Social Handles (`enrich_handles.py`)

Handles embedded Excel hyperlink targets, mixed post/profile URLs, and usernames:

```bash
# Basic run (auto-detects 'Social Handle' column)
python enrich_handles.py input_sheet.xlsx --out result.xlsx

# Test on the first 20 rows
python enrich_handles.py input_sheet.xlsx --max 20

# Run specific row ranges (1-indexed matching Excel rows)
python enrich_handles.py input_sheet.xlsx --rows 2,5,10-25

# Specify custom handle/link column
python enrich_handles.py input_sheet.xlsx --handles-col "Social Handle"
```

### 2. Enriching Post Links (`enrich.py`)

Visits individual Instagram and YouTube post URLs:

```bash
# Basic run (defaults to 'Post Link' column)
python enrich.py campaign_posts.xlsx --out campaign_enriched.xlsx

# Custom column names and parallel YouTube workers
python enrich.py campaign_posts.xlsx --link-col "URL" --youtube-workers 8

# Run specific row ranges (1-indexed matching Excel rows, 2 = first data row)
python enrich.py campaign_posts.xlsx --rows 2,5,10-25
```

---

## ⚡ Architecture & Performance Features

- **Hyperlink Target Extraction**: Uses `openpyxl` to extract underlying URL targets from formatted display text, sheet hyperlinks, and `=HYPERLINK()` formulas.
- **High-Throughput Apify Batching**: Groups hundreds of URLs/handles into concurrent Apify actor runs to minimize cold-start latency and optimize credit usage.
- **Smart Handle & Header Filter**: Accurately differentiates social handles and URLs from plain-text headers (e.g. `'2022-23'`), categories, or empty rows.
- **Parallel & Non-Blocking YouTube Fetching**: Leverages multi-threaded `yt-dlp` execution with `extract_flat=True`, socket timeouts, and HTML fallback to prevent channel hangs.
- **In-Memory Caching & Deduplication**: Repeated handles and creators across large sheets are fetched only once per run.
- **Offline Genre Classifier**: Heuristic-based keyword classification across 17 categories without external API latency or cost.
