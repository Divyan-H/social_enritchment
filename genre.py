"""
genre.py — classifies a post's genre/content-type from its caption/description
and account name. Pure keyword heuristic — no LLM, no API key, no internet
call needed for this step (works fully offline).

Taxonomy:
  comedy, educational, emotional, promotional, lifestyle,
  challenge, entertainment, review, music_dance, food, travel,
  fashion_beauty, fitness_health, tech, gaming, news_politics, other
"""
import re

SIGNALS: dict[str, list[str]] = {
    "comedy": ["funny", "humor", "joke", "lol", "lmao", "haha", "meme",
               "skit", "prank", "comedy", "hilarious", "rofl", "relatable",
               "😂", "🤣", "😆"],
    "educational": ["tips", "tricks", "tutorial", "learn", "how to", "howto",
                    "facts", "knowledge", "explained", "guide", "did you know",
                    "hack", "teach", "step by step", "📚", "💡", "🎓"],
    "emotional": ["inspiration", "motivation", "emotional story", "journey",
                  "struggling", "overcome", "blessed", "grateful", "tears",
                  "hope", "strength", "🥹", "🙏", "💪"],
    "promotional": ["sponsored", "paid partnership", "promo", "promotion",
                    "discount", "offer", "buy now", "sale", "collab",
                    "collaboration", "partnership", "gifted", "ambassador",
                    "use code", "link in bio", "shop now", "#ad", " ad ",
                    "🛍️", "🎁", "🔗"],
    "food": ["recipe", "cooking", "foodie", "restaurant", "cafe", "cuisine",
             "delicious", "yummy", "eat", "meal", "🍽️", "🍕", "🍔"],
    "travel": ["travel", "trip", "vacation", "explore", "wanderlust",
               "destination", "flight", "hotel", "✈️", "🌍", "🧳"],
    "fashion_beauty": ["fashion", "ootd", "outfit", "makeup", "skincare",
                       "beauty", "style", "aesthetic", "haul", "💄", "👗"],
    "fitness_health": ["fitness", "workout", "gym", "health", "diet",
                       "wellness", "yoga", "training", "🏋️", "🧘"],
    "music_dance": ["dance", "music", "song", "singing", "choreography",
                    "cover", "beat", "concert", "🎵", "🎶", "💃"],
    "tech": ["tech", "gadget", "software", "app", "ai ", "coding",
             "review of", "unboxing tech", "smartphone", "laptop"],
    "gaming": ["gaming", "gameplay", "gamer", "esports", "playthrough",
               "level up", "🎮"],
    "news_politics": ["news", "breaking", "election", "government",
                      "politics", "policy"],
    "challenge": ["challenge", "trend", "viral", "duet", "try this",
                  "trending", "fyp", "foryou", "🔥"],
    "entertainment": ["show", "talent", "performance", "acting", "art",
                      "entertainment", "artist", "performer", "🎭", "🎨"],
    "review": ["review", "unboxing", "rating", "honest review",
               "first impression", "worth it", "recommend", "⭐"],
    "lifestyle": ["daily", "vlog", "day in my life", "routine", "home",
                  "morning routine", "weekend", "✨", "🏠"],
}


def classify_genre(caption: str = "", account_name: str = "") -> tuple[str, float]:
    """
    Returns (genre, confidence 0-1).
    Falls back to 'other' if no signal at all matches.
    """
    combined = f"{caption or ''} {account_name or ''}".lower()
    if not combined.strip():
        return "other", 0.0

    scores: dict[str, float] = {}
    for genre, words in SIGNALS.items():
        score = 0.0
        for w in words:
            if w in combined:
                score += 2.0 if len(w) > 6 else 1.0
        if score:
            scores[genre] = score

    if not scores:
        return "other", 0.0

    best = max(scores, key=scores.get)
    total = sum(scores.values())
    conf = round(min(0.95, scores[best] / total), 2) if total else 0.0
    return best, conf


def _extract_hashtags(text: str) -> list:
    return list(dict.fromkeys(re.findall(r"#(\w+)", text or "")))
