"""
scraper.py — visits an Instagram or YouTube post link and pulls:
  - post description / caption
  - account handle
  - followers (Instagram) / subscribers (YouTube)
  - following (Instagram only)

Strategy:
  YouTube            -> yt-dlp (subscriber count comes free in the same call)
  Instagram (any post)-> yt-dlp first (works for many public reels/posts and
                          also returns follower count directly when available)
  Instagram fallback  -> embed page scrape for caption + handle, then a
                          second request to instagram.com/<user>/ to parse
                          followers/following/posts out of the public
                          og:description meta tag.

Every account is only looked up ONCE per run (in-memory cache), so a sheet
with many rows from the same creator doesn't cost extra requests.
"""
import re
import time
import threading
from urllib.parse import urlparse
from dataclasses import dataclass, field
from typing import Optional

import requests
from bs4 import BeautifulSoup

try:
    import yt_dlp
    _YTDLP = True
except ImportError:
    _YTDLP = False

BROWSER_UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": BROWSER_UA,
    "Accept": "text/html,application/xhtml+xml,application/json;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}


@dataclass
class ScrapedPost:
    url: str
    platform: str = ""          # "instagram" | "youtube" | ""
    ok: bool = False
    error: str = ""
    handle: str = ""            # account username / channel name
    description: str = ""       # post caption / video description
    followers: Optional[int] = None     # Instagram followers
    following: Optional[int] = None     # Instagram following
    subscribers: Optional[int] = None   # YouTube subscribers
    genre: str = ""
    genre_confidence: float = 0.0


# ── small helpers ────────────────────────────────────────────────────────

def detect_platform(url: str, platform_hint: str = "") -> str:
    hint = (platform_hint or "").strip().lower()
    if "insta" in hint:
        return "instagram"
    if "you" in hint:  # "youtube", "yt"
        return "youtube"
    host = urlparse(url).netloc.lower().lstrip("www.")
    if "instagram" in host or "instagr.am" in host:
        return "instagram"
    if "youtube" in host or "youtu.be" in host:
        return "youtube"
    return ""


def _parse_count(raw: str) -> Optional[int]:
    """'12.3K' / '1,234' / '4.5M' -> int"""
    if not raw:
        return None
    raw = raw.strip().replace(",", "")
    m = re.match(r"^([\d.]+)\s*([KMB]?)$", raw, re.I)
    if not m:
        return None
    num = float(m.group(1))
    mult = {"": 1, "K": 1_000, "M": 1_000_000, "B": 1_000_000_000}[m.group(2).upper()]
    return int(num * mult)


def _get(url: str, timeout: int = 12) -> Optional[requests.Response]:
    for attempt in range(3):
        try:
            r = requests.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
            # Instagram rate-limits/login-walls with 429 or 401/403. Back off and
            # retry once or twice instead of giving up immediately - a lot of
            # "failures" are transient throttling, not actually dead links.
            if r.status_code == 429 and attempt < 2:
                time.sleep(3 * (attempt + 1))
                continue
            return r
        except requests.exceptions.RequestException:
            if attempt == 2:
                return None
            time.sleep(1.5)
    return None


def _classify_http_failure(r: Optional[requests.Response]) -> str:
    """Turn a failed/blocked response into a specific, actionable reason."""
    if r is None:
        return "network error / timeout reaching Instagram"
    if r.status_code == 429:
        return "rate-limited (HTTP 429) - Instagram is throttling this IP"
    if r.status_code in (401, 403):
        return "login wall (HTTP %d) - Instagram requires auth to view this" % r.status_code
    if r.status_code == 404:
        return "not found (HTTP 404) - post/account likely removed or handle wrong"
    if r.status_code >= 500:
        return f"Instagram server error (HTTP {r.status_code})"
    if "login" in r.url.lower() or "accounts/login" in r.text[:2000].lower():
        return "redirected to login page - auth required for this content"
    return f"unexpected response (HTTP {r.status_code})"


class DomainThrottle:
    """Simple per-domain politeness delay."""
    def __init__(self, delay: float = 1.2):
        self._last: dict[str, float] = {}
        self._lock = threading.Lock()
        self.delay = delay

    def wait(self, url: str):
        domain = urlparse(url).netloc
        with self._lock:
            gap = time.time() - self._last.get(domain, 0)
            if gap < self.delay:
                time.sleep(self.delay - gap)
            self._last[domain] = time.time()


_throttle = DomainThrottle()

# in-memory cache: account handle -> {followers, following, subscribers}
_profile_cache: dict[str, dict] = {}
_profile_fail_cache: dict[str, str] = {}
_profile_cache_lock = threading.Lock()


# ── yt-dlp based fetch (used for both YouTube and Instagram) ───────────────

def _fetch_ytdlp(url: str) -> Optional[dict]:
    if not _YTDLP:
        return None
    opts = {
        "quiet": True,
        "no_warnings": True,
        "skip_download": True,
        "ignoreerrors": True,
        "logger": type("NullLogger", (), {
            "debug": lambda s, m: None, "info": lambda s, m: None,
            "warning": lambda s, m: None, "error": lambda s, m: None,
        })(),
    }
    try:
        with yt_dlp.YoutubeDL(opts) as ydl:
            info = ydl.extract_info(url, download=False)
        if not info:
            return None
        return {
            "description": (info.get("description") or "")[:1500],
            "handle": info.get("uploader_id") or info.get("uploader") or info.get("channel") or "",
            "followers_or_subs": info.get("channel_follower_count"),
        }
    except Exception:
        return None


# ── Instagram-specific fallback (embed page + profile page) ────────────────

def _extract_ig_username(html: str) -> str:
    m = re.search(r'class="[^"]*UsernameText[^"]*"[^>]*>([A-Za-z0-9_.]+)<', html)
    if m:
        return m.group(1)
    m = re.search(r'href="https://www\.instagram\.com/([A-Za-z0-9_.]+)/?"', html)
    if m and m.group(1).lower() not in ("p", "reel", "tv", "explore"):
        return m.group(1)
    return ""


def _fetch_instagram_embed(url: str) -> dict:
    out = {"description": "", "handle": "", "fail_reason": ""}
    shortcode = re.search(r"/(?:p|reel|tv)/([A-Za-z0-9_-]+)", url)
    if not shortcode:
        out["fail_reason"] = "URL doesn't look like a post/reel/tv link (no shortcode found)"
        return out
    embed_url = f"https://www.instagram.com/p/{shortcode.group(1)}/embed/captioned/"
    _throttle.wait(embed_url)
    r = _get(embed_url)
    if not r or r.status_code != 200:
        out["fail_reason"] = _classify_http_failure(r)
        return out
    html = r.text
    soup = BeautifulSoup(html, "html.parser")

    out["handle"] = _extract_ig_username(html)

    caption_el = (
        soup.find("div", class_="Caption")
        or soup.find("span", class_="Caption")
        or soup.find("div", class_="Caption--container")
    )
    if caption_el:
        out["description"] = caption_el.get_text(" ", strip=True)[:1500]
    else:
        candidates = sorted(
            [t.get_text(" ", strip=True) for t in soup.find_all(["p", "span", "div"])
             if len(t.get_text("", strip=True)) > 40],
            key=len, reverse=True,
        )
        if candidates:
            out["description"] = candidates[0][:1500]

    if not out["description"] and not out["handle"]:
        out["fail_reason"] = "got HTTP 200 but page had no caption/handle - likely a login-wall page served with a 200 status"
    return out


def _fetch_instagram_profile_stats(username: str) -> Optional[dict]:
    if not username:
        return None
    key = username.lower()
    with _profile_cache_lock:
        if key in _profile_cache:
            return _profile_cache[key]

    result = None
    fail_reason = ""
    profile_url = f"https://www.instagram.com/{username}/"
    _throttle.wait(profile_url)
    r = _get(profile_url)
    if r and r.status_code == 200:
        soup = BeautifulSoup(r.text, "html.parser")
        desc = ""
        for m in soup.find_all("meta"):
            name = (m.get("name") or m.get("property") or "").lower()
            if name in ("og:description", "description"):
                desc = m.get("content") or ""
                break
        m = re.search(
            r"([\d.,]+[KMB]?)\s*Followers,\s*([\d.,]+[KMB]?)\s*Following,\s*([\d.,]+[KMB]?)\s*Posts",
            desc, re.I,
        )
        if m:
            result = {
                "followers": _parse_count(m.group(1)),
                "following": _parse_count(m.group(2)),
            }
        else:
            fail_reason = "got HTTP 200 but couldn't find follower/following counts in og:description - likely a login-wall page"
    else:
        fail_reason = _classify_http_failure(r)

    with _profile_cache_lock:
        _profile_cache[key] = result
        _profile_fail_cache[key] = fail_reason
    return result


def get_profile_fail_reason(username: str) -> str:
    return _profile_fail_cache.get((username or "").lower(), "")


# ── public entry point ──────────────────────────────────────────────────────

def scrape_post(url: str, platform_hint: str = "") -> ScrapedPost:
    from genre import classify_genre  # local import to avoid circulars in tests

    sp = ScrapedPost(url=url)
    sp.platform = detect_platform(url, platform_hint)

    if not sp.platform:
        sp.error = "unsupported/unrecognized platform (not Instagram or YouTube)"
        return sp

    try:
        if sp.platform == "youtube":
            data = _fetch_ytdlp(url)
            if data:
                sp.ok = True
                sp.description = data["description"]
                sp.handle = data["handle"]
                sp.subscribers = data["followers_or_subs"]
            else:
                sp.error = "yt-dlp failed to fetch this YouTube link"

        elif sp.platform == "instagram":
            fail_reasons = []

            # Try yt-dlp first (works for many reels, sometimes gives follower count)
            data = _fetch_ytdlp(url)
            if data and data.get("description"):
                sp.ok = True
                sp.description = data["description"]
                sp.handle = data["handle"]
                if data.get("followers_or_subs") is not None:
                    sp.followers = data["followers_or_subs"]
            elif _YTDLP:
                fail_reasons.append("yt-dlp: no usable data returned")

            # Fill gaps / fallback via embed page
            if not sp.ok or not sp.handle:
                embed = _fetch_instagram_embed(url)
                if embed.get("description") and not sp.description:
                    sp.description = embed["description"]
                    sp.ok = True
                if embed.get("handle") and not sp.handle:
                    sp.handle = embed["handle"]
                if embed.get("fail_reason"):
                    fail_reasons.append(f"embed: {embed['fail_reason']}")

            # Followers/following from the account's public profile page
            if sp.handle and (sp.followers is None or sp.following is None):
                stats = _fetch_instagram_profile_stats(sp.handle)
                if stats:
                    if sp.followers is None:
                        sp.followers = stats.get("followers")
                    sp.following = stats.get("following")
                else:
                    reason = get_profile_fail_reason(sp.handle)
                    if reason:
                        fail_reasons.append(f"profile: {reason}")

            if not sp.ok:
                detail = "; ".join(fail_reasons) if fail_reasons else "no data source succeeded"
                sp.error = f"could not fetch this Instagram link ({detail})"

        # Genre classification (offline, keyword-based)
        if sp.description or sp.handle:
            genre, conf = classify_genre(sp.description, sp.handle)
            sp.genre = genre
            sp.genre_confidence = conf

    except Exception as e:
        sp.error = f"{type(e).__name__}: {e}"

    return sp
