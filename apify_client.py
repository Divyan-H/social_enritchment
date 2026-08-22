"""
apify_client.py — thin wrapper around Apify's "Instagram Scraper" actor
(apify/instagram-scraper) for fetching Instagram post and profile data.

Auth: reads the token from the APIFY_API_TOKEN environment variable by
default (recommended - keeps it out of shell history / argv / logs), or
accepts an explicit token passed in by the caller (e.g. from --apify-token).

Docs: https://apify.com/apify/instagram-scraper (run-sync-get-dataset-items
endpoint used here for simplicity - one HTTP call in, JSON results out).
"""
import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Optional

import requests

APIFY_ACTOR = "apify~instagram-scraper"
APIFY_BASE = f"https://api.apify.com/v2/acts/{APIFY_ACTOR}/run-sync-get-dataset-items"

# Generous timeout: Apify runs the actual scrape on their infra per-request,
# which can take longer than a normal HTTP call, especially for profiles
# with a lot of recent posts.
_TIMEOUT = 90


def get_token(explicit_token: Optional[str] = None) -> Optional[str]:
    return explicit_token or os.environ.get("APIFY_API_TOKEN")


def _run_actor(payload: dict, token: str) -> Optional[list]:
    try:
        r = requests.post(
            APIFY_BASE,
            params={"token": token},
            json=payload,
            timeout=_TIMEOUT,
        )
    except requests.exceptions.RequestException as e:
        return None
    if r.status_code == 401:
        raise PermissionError("Apify rejected the API token (401) - check APIFY_API_TOKEN")
    if r.status_code == 429:
        # Apify-side rate limiting (concurrent run limits on your plan) -
        # one retry after a short backoff is usually enough.
        time.sleep(5)
        try:
            r = requests.post(APIFY_BASE, params={"token": token}, json=payload, timeout=_TIMEOUT)
        except requests.exceptions.RequestException:
            return None
    if r.status_code != 201 and r.status_code != 200:
        return None
    try:
        data = r.json()
    except ValueError:
        return None
    return data if isinstance(data, list) else None


def fetch_post(url: str, token: str) -> dict:
    """Full post-level data: caption, engagement, owner, media, etc."""
    payload = {
        "directUrls": [url],
        "resultsType": "details",
        "resultsLimit": 1,
        "addParentData": False,
    }
    items = _run_actor(payload, token)
    if not items:
        return {}
    return items[0] or {}


def fetch_profile(username_or_url: str, token: str) -> dict:
    """Full profile-level data: followers, bio, verification, category, etc."""
    url = username_or_url
    if not url.startswith("http"):
        url = f"https://www.instagram.com/{username_or_url}/"
    payload = {
        "directUrls": [url],
        "resultsType": "details",
        "resultsLimit": 1,
        "addParentData": False,
    }
    items = _run_actor(payload, token)
    if not items:
        return {}
    return items[0] or {}


# ── batch versions: pack many URLs into ONE actor run ───────────────────
# This is the actual speed lever. Each actor run has real cold-start
# overhead (Apify has to spin up a container on their infra) independent
# of how many URLs it processes - so 100 one-URL runs is far slower than
# ~2-4 runs of 25-50 URLs each. Chunk size is capped conservatively; Apify
# itself can usually handle larger batches, but a huge single run risks a
# very long-running sync HTTP call and makes partial failures harder to
# retry, so we keep chunks moderate and run chunks in parallel instead.

DEFAULT_CHUNK_SIZE = 40
DEFAULT_MAX_PARALLEL_CHUNKS = 4


def _chunk(items: list, size: int) -> list:
    return [items[i:i + size] for i in range(0, len(items), size)]


def _batch_timeout(n_urls: int) -> int:
    # Scale timeout with batch size; floor at the single-item timeout.
    return max(_TIMEOUT, 30 + n_urls * 4)


def _run_actor_batch(urls: list, token: str, timeout: int) -> list:
    """Generic batch runner for POST URLs (used by fetch_many / fetch_posts_batch)."""
    payload = {
        "directUrls": urls,
        "resultsType": "details",
        "resultsLimit": len(urls),
        "addParentData": False,
    }
    try:
        r = requests.post(APIFY_BASE, params={"token": token}, json=payload, timeout=timeout)
    except requests.exceptions.RequestException:
        return []
    if r.status_code == 401:
        raise PermissionError("Apify rejected the API token (401) - check APIFY_API_TOKEN")
    if r.status_code not in (200, 201):
        return []
    try:
        data = r.json()
    except ValueError:
        return []
    return data if isinstance(data, list) else []


def _run_profiles_batch(profile_urls: list, token: str, timeout: int) -> list:
    """Dedicated batch runner for Instagram *profile* URLs.

    Key differences from _run_actor_batch:
    - Does NOT pass resultsLimit — when scraping profile-page URLs the
      Instagram actor interprets resultsLimit as "number of posts to fetch
      from the timeline", so passing len(urls) (e.g. 9) means only 9 posts
      total are fetched, which can cut off after 1 profile and return nothing
      for the rest.  Omitting it lets the actor use its own default and
      return one profile-data item per URL as intended.
    - Uses a larger per-profile timeout floor (profiles are slower than posts).
    """
    payload = {
        "directUrls": profile_urls,
        "resultsType": "details",
        "addParentData": False,
    }
    try:
        r = requests.post(APIFY_BASE, params={"token": token}, json=payload, timeout=timeout)
    except requests.exceptions.RequestException:
        return []
    if r.status_code == 401:
        raise PermissionError("Apify rejected the API token (401) - check APIFY_API_TOKEN")
    if r.status_code not in (200, 201):
        return []
    try:
        data = r.json()
    except ValueError:
        return []
    return data if isinstance(data, list) else []


def fetch_many(urls: list, token: str,
                chunk_size: int = DEFAULT_CHUNK_SIZE,
                max_parallel: int = DEFAULT_MAX_PARALLEL_CHUNKS) -> dict:
    """Fetch post OR profile details for many URLs, keyed by input URL.

    Splits into chunks (one actor run per chunk) and runs chunks
    concurrently. Returns {input_url: raw_item_or_{}}. If Apify returns
    fewer items than URLs requested (some URLs failed inside the actor),
    unmatched URLs simply get {} - caller decides how to handle gaps
    (e.g. fall back to the free scrape for just those).
    """
    urls = list(dict.fromkeys(urls))  # dedupe, preserve order
    if not urls:
        return {}

    chunks = _chunk(urls, chunk_size)
    results: dict = {u: {} for u in urls}

    def run_one_chunk(chunk_urls):
        items = _run_actor_batch(chunk_urls, token, _batch_timeout(len(chunk_urls)))
        return chunk_urls, items

    with ThreadPoolExecutor(max_workers=min(max_parallel, len(chunks))) as ex:
        futures = [ex.submit(run_one_chunk, c) for c in chunks]
        for fut in as_completed(futures):
            chunk_urls, items = fut.result()
            # Match items back to input URLs. Apify echoes the requested
            # URL in a couple of possible fields depending on item type.
            remaining = set(chunk_urls)
            for item in items:
                candidate_urls = [
                    item.get("url"), item.get("inputUrl"),
                    item.get("postUrl"), item.get("permalink"),
                ]
                matched = None
                for cu in candidate_urls:
                    if cu and cu in remaining:
                        matched = cu
                        break
                if not matched:
                    # loose match: normalize trailing slash
                    for cu in candidate_urls:
                        if not cu:
                            continue
                        norm = cu.rstrip("/")
                        for ru in remaining:
                            if ru.rstrip("/") == norm:
                                matched = ru
                                break
                        if matched:
                            break
                if matched:
                    results[matched] = item
                    remaining.discard(matched)
            # anything left in `remaining` had no matching item -> stays {}
    return results


def fetch_posts_batch(urls: list, token: str, **kwargs) -> dict:
    """{post_url: raw_item}"""
    return fetch_many(urls, token, **kwargs)


# Per-profile timeout: profiles need more time than posts
_PROFILE_TIMEOUT_PER = 15   # seconds per profile in a batch
_PROFILE_TIMEOUT_FLOOR = 90


def _profile_batch_timeout(n: int) -> int:
    return max(_PROFILE_TIMEOUT_FLOOR, n * _PROFILE_TIMEOUT_PER)


def _match_items_to_usernames(all_items: list, usernames: list) -> dict:
    """Match Apify response items back to the original handle strings.

    Checks `username`, `ownerUsername`, `handle`, and the last path segment
    of any instagram.com URL field in the response item.  Case-insensitive.
    Returns {original_username: raw_item} for every item that matched.
    """
    result: dict[str, dict] = {}
    username_lower = {u.lower(): u for u in usernames}

    for item in all_items:
        candidates = [
            item.get("username"),
            item.get("ownerUsername"),
            item.get("handle"),
        ]
        for url_field in ("url", "profileUrl", "inputUrl", "permalink"):
            raw_url = item.get(url_field) or ""
            if "instagram.com" in raw_url:
                seg = raw_url.rstrip("/").rsplit("/", 1)[-1]
                if seg:
                    candidates.append(seg)

        for cand in candidates:
            if not cand:
                continue
            key = str(cand).lstrip("@").lower()
            if key in username_lower:
                orig = username_lower[key]
                if orig not in result:
                    result[orig] = item
                break

    return result


def fetch_profiles_batch(usernames: list, token: str,
                          chunk_size: int = DEFAULT_CHUNK_SIZE,
                          max_parallel: int = DEFAULT_MAX_PARALLEL_CHUNKS,
                          **kwargs) -> dict:
    """{username: raw_item} — accepts bare usernames or full profile URLs.

    Sends profile URLs to the Instagram actor in chunks via _run_profiles_batch
    (which omits resultsLimit so the actor returns one profile item per URL).
    Matches returned items to input handles by username field, not by URL.
    Any handles the batch missed are retried individually in parallel.
    """
    usernames = list(dict.fromkeys(usernames))  # dedupe, preserve order
    if not usernames:
        return {}

    # Build profile URLs
    url_for: dict[str, str] = {}   # full_url -> original_username
    urls: list[str] = []
    for u in usernames:
        full = u if u.startswith("http") else f"https://www.instagram.com/{u}/"
        url_for[full] = u
        urls.append(full)

    # ── Phase 1: batch fetch ──────────────────────────────────────────────
    chunks = _chunk(urls, chunk_size)
    all_items: list[dict] = []

    def run_one_chunk(chunk_urls):
        return _run_profiles_batch(chunk_urls, token, _profile_batch_timeout(len(chunk_urls)))

    with ThreadPoolExecutor(max_workers=min(max_parallel, len(chunks))) as ex:
        futures = [ex.submit(run_one_chunk, c) for c in chunks]
        for fut in as_completed(futures):
            all_items.extend(fut.result())

    result = _match_items_to_usernames(all_items, usernames)

    # ── Phase 2: per-handle retry for anything the batch missed ──────────
    # The Instagram actor occasionally returns fewer items than expected
    # when multiple profile URLs are batched.  Retry individually so we
    # don't silently lose data.
    missed = [u for u in usernames if u not in result]
    if missed:
        print(f"  [apify] Batch missed {len(missed)} profile(s), retrying individually...")

        def fetch_one(u):
            profile_url = u if u.startswith("http") else f"https://www.instagram.com/{u}/"
            items = _run_profiles_batch([profile_url], token, _profile_batch_timeout(1))
            return u, items

        with ThreadPoolExecutor(max_workers=min(8, len(missed))) as ex:
            futures2 = [ex.submit(fetch_one, u) for u in missed]
            for fut in as_completed(futures2):
                u, items = fut.result()
                if items:
                    matched = _match_items_to_usernames(items, [u])
                    result.update(matched)

    # Ensure every input username has an entry (empty dict = no data)
    for u in usernames:
        result.setdefault(u, {})

    return result


# ── field extraction helpers ────────────────────────────────────────────
# Apify's actor response schema has shifted field names across versions in
# the past, so we check a couple of likely keys for each value rather than
# assuming one exact shape.

def extract_post_fields(raw: dict) -> dict:
    if not raw:
        return {}
    return {
        "description": raw.get("caption") or "",
        "handle": raw.get("ownerUsername") or (raw.get("owner") or {}).get("username") or "",
        "likes": raw.get("likesCount"),
        "comments": raw.get("commentsCount"),
        "video_views": raw.get("videoViewCount") or raw.get("videoPlayCount"),
        "post_type": raw.get("type") or raw.get("productType"),
        "timestamp": raw.get("timestamp"),
        "hashtags": raw.get("hashtags") or [],
        "mentions": raw.get("mentions") or [],
        "location": (raw.get("locationName") or ""),
        "is_sponsored": raw.get("isSponsored"),
    }


def extract_profile_fields(raw: dict) -> dict:
    if not raw:
        return {}
    return {
        "followers": raw.get("followersCount"),
        "following": raw.get("followsCount") or raw.get("followingCount"),
        "posts_count": raw.get("postsCount"),
        "biography": raw.get("biography") or "",
        "full_name": raw.get("fullName") or "",
        "is_verified": raw.get("verified") or raw.get("isVerified"),
        "is_private": raw.get("private") or raw.get("isPrivate"),
        "is_business_account": raw.get("isBusinessAccount"),
        "business_category": raw.get("businessCategoryName") or "",
        "external_url": raw.get("externalUrl") or raw.get("externalUrls"),
        "profile_pic_url": raw.get("profilePicUrl") or raw.get("profilePicUrlHD"),
    }


# ── YouTube channel scraper (Apify actor: streamers/youtube-scraper) ────────
# Accepts channel handles (@name), channel URLs, or channel IDs.
# One actor run per chunk — batched + parallelised the same way as Instagram.

APIFY_YT_ACTOR = "streamers~youtube-scraper"
APIFY_YT_BASE = f"https://api.apify.com/v2/acts/{APIFY_YT_ACTOR}/run-sync-get-dataset-items"

# The YouTube actor can be slower than Instagram's, especially for large
# channels (it reads the channel page and About tab). Scale timeout generously.
_YT_TIMEOUT_PER_ITEM = 8   # seconds per channel URL in a batch
_YT_TIMEOUT_FLOOR = 120    # always wait at least this long


def _yt_batch_timeout(n: int) -> int:
    return max(_YT_TIMEOUT_FLOOR, n * _YT_TIMEOUT_PER_ITEM)


def _run_yt_actor_batch(start_urls: list, token: str, timeout: int) -> list:
    """POST a batch of YouTube channel URLs to the Apify YouTube actor.

    ``start_urls`` items follow the Apify RequestList format: a list of dicts
    with a ``url`` key (e.g. ``[{"url": "https://www.youtube.com/@NASA"}]``).
    Returns the raw list of dataset items, or [] on any error.
    """
    payload = {
        "startUrls": start_urls,
        # Fetch channel-level data only; we don't need individual videos
        "maxResults": 0,
        "maxResultShorts": 0,
        "maxResultStreams": 0,
    }
    try:
        r = requests.post(
            APIFY_YT_BASE,
            params={"token": token},
            json=payload,
            timeout=timeout,
        )
    except requests.exceptions.RequestException:
        return []
    if r.status_code == 401:
        raise PermissionError("Apify rejected the API token (401) — check APIFY_API_TOKEN")
    if r.status_code not in (200, 201):
        return []
    try:
        data = r.json()
    except ValueError:
        return []
    return data if isinstance(data, list) else []


def _match_yt_items_to_handles(all_items: list, handles: list, url_for_handle: dict) -> dict:
    result: dict[str, dict] = {}
    handles_lower = {h.lower().lstrip("@"): h for h in handles}
    url_to_handle_norm = {u.rstrip("/").lower(): h for u, h in url_for_handle.items()}

    for item in all_items:
        matched_handle = None

        # 1. Match by URL candidates
        candidate_urls = [
            item.get("inputUrl"),
            item.get("channelUrl"),
            item.get("url"),
        ]
        for cu in candidate_urls:
            if not cu:
                continue
            norm = cu.rstrip("/").lower()
            if norm in url_to_handle_norm:
                matched_handle = url_to_handle_norm[norm]
                break

        # 2. Match by handle / username / title candidates
        if not matched_handle:
            candidates = [
                item.get("channelHandle"),
                item.get("handle"),
                item.get("customUrl"),
                item.get("channelName"),
                item.get("title"),
            ]
            for cu in candidate_urls:
                if cu and "youtube.com" in cu:
                    seg = cu.rstrip("/").rsplit("/", 1)[-1]
                    if seg:
                        candidates.append(seg)

            for cand in candidates:
                if not cand:
                    continue
                key = str(cand).strip().lstrip("@").lower()
                if key in handles_lower:
                    matched_handle = handles_lower[key]
                    break

        if matched_handle and matched_handle not in result:
            result[matched_handle] = item

    return result


def fetch_youtube_channels_batch(
    handles: list,
    token: str,
    chunk_size: int = DEFAULT_CHUNK_SIZE,
    max_parallel: int = DEFAULT_MAX_PARALLEL_CHUNKS,
) -> dict:
    """{handle: raw_item} for a list of YouTube channel handles/URLs.

    Accepts bare handles (with or without ``@``), full channel URLs, or
    channel IDs — anything the Apify YouTube actor accepts.
    Returns a dict keyed by the *original* handle string you passed in.
    Handles that Apify returned nothing for are absent from the result dict
    (caller decides how to treat gaps, e.g. fall back to yt-dlp).
    """
    handles = list(dict.fromkeys(handles))  # dedupe, preserve order
    if not handles:
        return {}

    # Normalise to full URLs so we can match them back after the run
    url_for_handle: dict[str, str] = {}   # normalised_url -> original handle
    start_url_objects: list[dict] = []
    for h in handles:
        if h.startswith("http"):
            url = h
        elif h.startswith("@"):
            url = f"https://www.youtube.com/{h}"
        else:
            url = f"https://www.youtube.com/@{h}"
        url_for_handle[url] = h
        start_url_objects.append({"url": url})

    chunks = _chunk(start_url_objects, chunk_size)
    all_items: list[dict] = []

    def run_one_chunk(chunk_objs):
        return _run_yt_actor_batch(chunk_objs, token, _yt_batch_timeout(len(chunk_objs)))

    with ThreadPoolExecutor(max_workers=min(max_parallel, len(chunks))) as ex:
        futures = [ex.submit(run_one_chunk, c) for c in chunks]
        for fut in as_completed(futures):
            all_items.extend(fut.result())

    results = _match_yt_items_to_handles(all_items, handles, url_for_handle)

    # Phase 2: per-channel retry if batch missed anything
    missed = [h for h in handles if h not in results]
    if missed:
        print(f"  [apify] YouTube batch missed {len(missed)} channel(s), retrying individually...")

        def fetch_one(h):
            u = h if h.startswith("http") else (f"https://www.youtube.com/{h}" if h.startswith("@") else f"https://www.youtube.com/@{h}")
            items = _run_yt_actor_batch([{"url": u}], token, _yt_batch_timeout(1))
            return h, items, u

        with ThreadPoolExecutor(max_workers=min(8, len(missed))) as ex:
            futures2 = [ex.submit(fetch_one, h) for h in missed]
            for fut in as_completed(futures2):
                h, items, u = fut.result()
                if items:
                    matched = _match_yt_items_to_handles(items, [h], {u: h})
                    results.update(matched)

    return results


def extract_youtube_channel_fields(raw: dict) -> dict:
    """Extract standardised channel fields from a raw Apify YouTube actor item.

    The actor has shifted field names across versions, so we check several
    candidate keys for each value (same defensive pattern as Instagram).
    """
    if not raw:
        return {}

    # Subscriber count: actor returns both a number and a text form
    subscribers = (
        raw.get("numberOfSubscribers")
        or raw.get("subscriberCount")
        or raw.get("subscribers")
    )
    if isinstance(subscribers, str):
        # e.g. "15M" — convert to int using the same helper pattern
        import re as _re
        m = _re.match(r"^([\d.]+)\s*([KMB]?)$", subscribers.replace(",", ""), _re.I)
        if m:
            mult = {"": 1, "K": 1_000, "M": 1_000_000, "B": 1_000_000_000}[m.group(2).upper()]
            subscribers = int(float(m.group(1)) * mult)
        else:
            subscribers = None

    video_count = (
        raw.get("numberOfVideos")
        or raw.get("videoCount")
        or raw.get("videosCount")
    )

    channel_name = (
        raw.get("channelName")
        or raw.get("title")
        or raw.get("name")
        or ""
    )

    handle = (
        raw.get("channelHandle")
        or raw.get("handle")
        or raw.get("customUrl")
        or ""
    )
    # Strip leading @ if present, for consistency
    if handle.startswith("@"):
        handle = handle[1:]

    description = (
        raw.get("channelDescription")
        or raw.get("description")
        or raw.get("about")
        or ""
    )

    return {
        "channel_name": channel_name,
        "handle": handle,
        "subscribers": subscribers,
        "video_count": video_count,
        "description": description,
    }
