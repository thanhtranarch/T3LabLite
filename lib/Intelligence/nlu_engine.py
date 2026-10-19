# -*- coding: utf-8 -*-
"""NLU Engine: bilingual VI/EN intent parsing for Revit commands.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

from __future__ import unicode_literals, division

__author__  = "Tran Tien Thanh"
__title__   = "NLU Engine"

import random
import re
import unicodedata


# ─── Diacritics helper ────────────────────────────────────────────────────────

def _strip_diacritics(text):
    """Remove combining diacritics. NFD does not decompose đ/Đ, so they are
    folded to d/D explicitly (else "được" -> "đuoc" and regex cleaning eats the đ)."""
    try:
        nfd = unicodedata.normalize('NFD', text)
        out = ''.join(c for c in nfd if unicodedata.category(c) != 'Mn')
        return out.replace(u'đ', 'd').replace(u'Đ', 'D')
    except Exception:
        return text


def _norm(text):
    """Strip diacritics + lowercase."""
    return _strip_diacritics(text).lower()


# ─── Abbreviation / synonym map ───────────────────────────────────────────────
# Applied after _norm(), in order: longer keys first.

_ABBREVS = [
    # ── Tool name normalisations ───────────────────────────────────────────────
    # BatchOut
    ("batch out",       "batchout"),
    (" batcho ",        " batchout "),  # padded: "batcho" prefixes "batchout"
                                         # (unpadded -> "batchoutut")
    ("b.out",           "batchout"),
    # bare "bo " -> batchout is at the END of this list on purpose
    # ParaSync
    ("para sync",       "parasync"),
    ("dong bo tham so", "parasync"),
    ("dong bo thong so","parasync"),
    ("dong bo",         "parasync"),
    ("sync param",      "parasync"),
    ("parameter sync",  "parasync"),
    (" ps ",            " parasync "),
    # Load Family Cloud
    ("load family cloud","loadfamilycloud"),
    ("load fam cloud",  "loadfamilycloud"),
    ("tai family cloud","loadfamilycloud"),
    ("load cloud",      "loadfamilycloud"),
    (" lfc ",           " loadfamilycloud "),
    # Load Family
    ("load family",     "loadfamily"),
    (" load fam ",      " loadfamily "),  # padded: prefixes "load family"
                                           # (unpadded -> "loadfamilyily")
    ("tai family",      "loadfamily"),
    ("nap family",      "loadfamily"),
    ("keo family",      "loadfamily"),
    ("import family",   "loadfamily"),
    (" lf ",            " loadfamily "),
    # Project Name
    ("project name",    "projectname"),
    ("ten project",     "projectname"),
    ("ten du an",       "projectname"),
    ("dat ten project", "projectname"),
    ("thay ten",        "projectname"),
    (" pn ",            " projectname "),
    # Workset
    ("quan ly workset", "workset"),
    ("manage workset",  "workset"),
    (" ws ",            " workset "),
    # Upper Dim Text
    ("upper dim text",  "upperdimtext"),
    ("upper dim",       "upperdimtext"),
    (" udt ",           " upperdimtext "),
    # Dim Text
    ("dim text",        "dimtext"),
    ("dimension text",  "dimtext"),
    ("kich thuoc",      "dimtext"),
    ("sua dim",         "dimtext"),
    ("chinh dim",       "dimtext"),
    ("edit dim",        "dimtext"),
    ("edit dimension",  "dimtext"),
    (" dt ",            " dimtext "),
    # Reset Overrides
    ("reset graphic override","resetoverrides"),
    ("reset graphic",   "resetoverrides"),
    ("reset override",  "resetoverrides"),
    ("xoa override",    "resetoverrides"),
    ("bo override",     "resetoverrides"),
    ("xoa do ghi de",   "resetoverrides"),
    # Grids
    ("luoi truc",       "grids"),
    ("quan ly luoi",    "grids"),
    ("manage grid",     "grids"),

    # ── Intent-prefixed command patterns ──────────────────────────────────────
    # "muốn / cần / hãy / nhờ / làm ơn" before a verb → just keep the verb
    ("muon xuat",       "export"),
    ("muon in ",        "export "),
    ("muon mo ",        "open "),
    ("muon bat ",       "open "),
    ("muon chay ",      "open "),
    ("can xuat",        "export"),
    ("can in ",         "export "),
    ("can mo ",         "open "),
    ("hay xuat",        "export"),
    ("hay mo ",         "open "),
    ("lam on xuat",     "export"),
    ("lam on mo ",      "open "),
    ("lam on in ",      "export "),
    ("nho xuat",        "export"),
    ("nho mo ",         "open "),
    ("nho in ",         "export "),
    ("giup toi xuat",   "export"),
    ("giup toi mo ",    "open "),
    ("cho toi xuat",    "export"),
    ("cho toi mo ",     "open "),
    ("thu mo ",         "open "),
    ("thu xuat",        "export"),

    # ── Export verbs (specific bigrams first) ─────────────────────────────────
    ("in pdf",          "export pdf"),
    ("in dwg",          "export dwg"),
    ("in dwf",          "export dwf"),
    ("in dgn",          "export dgn"),
    ("in ifc",          "export ifc"),
    ("in nwd",          "export nwd"),
    ("in img",          "export img"),
    ("in sheet",        "export sheet"),
    ("in to ",          "export sheet "),  # "in tờ …"
    ("in het",          "export all"),
    ("in toan bo",      "export all"),
    ("in tat ca",       "export all"),
    ("in ra",           "export"),
    ("xuat ra",         "export"),
    ("xuat",            "export"),
    # Extra export synonyms
    ("ket xuat",        "export"),          # formal "render/export"
    ("save as pdf",     "export pdf"),
    ("save as dwg",     "export dwg"),
    ("convert to pdf",  "export pdf"),
    ("convert to dwg",  "export dwg"),
    ("chuyen sang pdf", "export pdf"),
    ("chuyen sang dwg", "export dwg"),
    ("luu thanh pdf",   "export pdf"),
    ("luu thanh dwg",   "export dwg"),
    ("xuat sang pdf",   "export pdf"),
    ("xuat sang dwg",   "export dwg"),
    ("out pdf",         "export pdf"),
    ("out dwg",         "export dwg"),

    # ── Open verbs ────────────────────────────────────────────────────────────
    ("mo len",          "open"),
    ("bat len",         "open"),
    ("chay len",        "open"),
    ("khoi dong",       "open"),
    ("khoi chay",       "open"),

    # ── Quantity / scope ──────────────────────────────────────────────────────
    ("tat ca sheet",    "all sheet"),
    ("toan bo sheet",   "all sheet"),
    ("all sheets",      "all sheet"),
    ("every sheet",     "all sheet"),
    ("all sheet",       "all sheet"),
    ("toan bo",         "all"),
    ("tat ca",          "all"),
    ("het sheets",      "all sheet"),
    ("het tat ca",      "all"),
    ("toan phan",       "all"),

    # ── Format synonyms ───────────────────────────────────────────────────────
    ("hinh anh",        "img"),
    ("image",           "img"),
    ("picture",         "img"),

    # ── "What can you do" capability queries ─────────────────────────────────
    # All stopwords once tokenised (score zero everywhere), so collapse the
    # phrase to one non-stopword marker.
    ("what can you do",         "capabilities query"),
    ("what can u do",           "capabilities query"),
    ("what do you do",          "capabilities query"),
    ("what could you do",      "capabilities query"),
    ("what are you capable of","capabilities query"),
    ("what can this do",        "capabilities query"),
    ("what can this tool do",   "capabilities query"),
    ("what features do you have","capabilities query"),
    ("what tools do you have",  "capabilities query"),
    ("how can you help me",     "capabilities query"),
    ("how can you help",        "capabilities query"),
    ("who are you",             "capabilities query"),
    ("what are you",            "capabilities query"),

    # ── Vietnamese → English capability verbs ────────────────────────────────
    # The tool catalog is English-named ("đổi tên view" found no tool). After
    # the specific phrases above ("quan ly workset"), which win. Verbs only:
    # folded nouns collide ("cua" = door / "của", "san" = floor / "sản").
    ("quan ly",         "manager"),
    ("quan li",         "manager"),
    ("doi ten",         "rename"),
    ("thay ten",        "rename"),
    ("dat ten",         "rename"),
    ("thu vien",        "library"),
    ("sao chep",        "copy"),
    ("di chuyen",       "move"),
    ("thong ke",        "schedule"),
    (" tao ",           " create "),
    (" chon ",          " select "),
    (" loc ",           " filter "),

    # ── Greeting shortcuts ────────────────────────────────────────────────────
    ("chao buoi sang",  "chao"),
    ("chao buoi chieu", "chao"),
    ("chao buoi toi",   "chao"),
    ("good morning",    "chao"),
    ("good afternoon",  "chao"),
    ("good evening",    "chao"),
    ("good night",      "chao"),
    ("xin chao",        "chao"),
    # Padded: unpadded "alo"/"a lo" mangled "kiem tra loi", "dialog", "catalog"
    (" alo ",           " chao "),
    (" a lo ",          " chao "),
    ("hi ban",          "chao"),
    ("hey ban",         "chao"),
    ("chao ban",        "chao"),

    # ── Thanks shortcuts ──────────────────────────────────────────────────────
    ("cam on nhieu",    "cam on"),
    ("cam on ban nhieu","cam on"),
    ("cam on ban rat nhieu", "cam on"),
    ("rat cam on",      "cam on"),
    ("tran trong",      "cam on"),
    ("biet on",         "cam on"),
    ("thank you so much","cam on"),
    ("thanks a lot",    "cam on"),
    ("thanks so much",  "cam on"),
    ("thank u",         "cam on"),
    (" tks ",           " cam on "),
    (" thks ",          " cam on "),
    (" thx ",           " cam on "),
    (" ty ",            " cam on "),

    # ── Acknowledgement shortcuts ─────────────────────────────────────────────
    ("ok roi",          "ok roi"),    # keep for trigger matching
    ("oke roi",         "ok roi"),
    ("duoc roi",        "ok roi"),
    ("chay roi",        "ok roi"),
    ("xong roi",        "ok roi"),
    ("hieu roi",        "hieu roi"),
    ("biet roi",        "hieu roi"),
    ("ra roi",          "hieu roi"),
    ("toi hieu",        "hieu roi"),
    ("ro rang",         "hieu roi"),
    ("dung roi",        "dung roi"),
    ("chinh xac",       "dung roi"),
    ("hop ly",          "dung roi"),

    # ── Vietnamese "open" at word boundary ───────────────────────────────────
    ("mo ",             "open "),
    ("mo\n",            "open\n"),

    # ── Bare "bo" -> BatchOut (LAST on purpose) ───────────────────────────────
    # "bo" collides with folded words ("toàn bộ" -> "toan bo"); every rule above
    # consumes its own "bo" first, so a leftover "bo" is the BatchOut shorthand.
    ("bo ",             "batchout "),
]


# Multi-word keys match on word boundaries ("what are you" ate into "what are
# your ..." -> "capabilities queryr"); single tokens stay substring because some
# are stems ("image" -> "images") or carry their own padding (" batcho ").
_ABBREV_RULES = []
for _src, _dst in _ABBREVS:
    if ' ' in _src.strip() and _src[-1:].isalnum() and _src[:1].isalnum():
        _ABBREV_RULES.append((re.compile(r'(?<![a-z0-9])' + re.escape(_src)
                                         + r'(?![a-z0-9])'), _dst, True))
    else:
        _ABBREV_RULES.append((_src, _dst, False))


def _expand(text):
    """Apply abbreviation / synonym substitutions to normalised text."""
    t = " " + text + " "
    for src, dst, is_re in _ABBREV_RULES:
        if is_re:
            # lambda replacement: keeps backslashes in `dst` literal
            t = src.sub(lambda m, _d=dst: _d, t)
        else:
            t = t.replace(src, dst)
    return t.strip()


# ─── Tokeniser ────────────────────────────────────────────────────────────────

_STOPWORDS = {
    # ── Vietnamese function / filler words ────────────────────────────────────
    "va", "de", "cho", "cai", "len", "ra", "vao", "di",
    "toi", "ban", "minh", "ho", "no", "cua", "voi", "trong",
    "la", "duoc", "co", "khong", "nhe", "nha", "nao", "gi",
    "a", "o", "roi", "se", "da", "dang", "rat", "khi",
    "neu", "sau", "truoc", "thi", "ma", "vay", "ay", "oi",
    "rang", "vi", "kia", "nhau", "hon",
    "mot", "hai", "ba", "bon", "nam",        # numbers (sheet count)
    "muon", "can", "hay", "thu", "giup",     # intent-prefix words
    "lam", "nho",                             # politeness words
    # "in" deliberately excluded: Vietnamese "print/export"
    # ── English function words ─────────────────────────────────────────────────
    "the", "an", "to", "for", "of", "at",
    "me", "my", "i", "you", "it", "is", "are", "was", "be",
    "and", "or", "not", "on", "up", "do", "as", "so",
    "can", "will", "would", "could", "should",
    "with", "from", "by", "all", "any", "some",
    "please", "just", "now", "here",
}


def _tokenise(text):
    """Return (unigrams, bigrams) as frozensets of normalised strings."""
    clean = re.sub(r'[^a-z0-9\s]', ' ', text)
    tokens = [w for w in clean.split() if len(w) >= 2 and w not in _STOPWORDS]
    unigrams = frozenset(tokens)
    bigrams  = frozenset(
        tokens[i] + " " + tokens[i + 1] for i in range(len(tokens) - 1)
    )
    return unigrams, bigrams


# ─── Intent trigger tables ────────────────────────────────────────────────────
# (feature, weight); unigrams and bigrams share the tables.

_TRIGGERS = {
    "open_batchout": [
        ("batchout",           20),
        ("open batchout",      30),
        ("mo batchout",        30),
        ("open",                3),
    ],

    "export_direct": [
        ("export",             15),
        ("export sheet",       20),
        ("export all",         20),
        ("export pdf",         20),  # bigram from "in pdf" expansion
        ("export dwg",         20),
        ("pdf",                 5),
        ("dwg",                 5),
        ("dwf",                 5),
        ("all sheet",          10),
        ("print",               8),
        ("in",                 10),  # Vietnamese print verb
        ("sheet",               4),
    ],

    "open_batchout_configured": [
        ("batchout",           15),
        ("open batchout",      20),
        ("mo batchout",        20),
        ("batchout pdf",       25),
        ("batchout dwg",       25),
        ("batchout dwf",       25),
        ("batchout sheet",     20),
        ("batchout all",       15),
        ("open",                3),
    ],

    "open_loadfamily": [
        ("loadfamily",         30),
        ("open loadfamily",    35),
        ("mo loadfamily",      35),
        ("family",              8),
        ("open",                2),
    ],

    # Only tools with a dedicated launcher belong here; auto-discovered tools go
    # through resolve_tool() (2026-07-28: stale open_workset etc. scored 35 and
    # beat it, so "mo workset" never reached ManaWorkset).

    "check_spelling": [
        # Proofread all Text Notes in the model
        ("spelling",           30),
        ("spellcheck",         35),
        ("spell check",        35),
        ("chinh ta",           35),
        ("loi chinh ta",       40),
        ("loi tieng anh",      40),
        ("tieng anh",          15),
        ("kiem tra loi",       22),
        ("check loi",          22),
        ("check english",      30),
        ("grammar",            28),
        ("ngu phap",           28),
        ("typo",               30),
        ("typos",              30),
        ("text note",          12),
        ("textnote",           12),
        ("review text",        15),
    ],

    "greet": [
        ("chao",               25),
        ("hello",              25),
        ("hey",                18),
        ("hi",                 20),
        ("howdy",              18),
        ("alo",                25),
        ("chao ban",           28),
        ("hi ban",             25),
        ("hey ban",            25),
        ("good morning",       20),
        # Bare time-of-day words must clear 18 alone (at 12 "morning" went to the
        # LLM, which hallucinated a tool intent that got mis-learned)
        ("morning",            20),
        ("afternoon",          20),
        ("evening",            20),
        ("yo",                 18),
        ("chao buoi",          25),
        # Farewell
        ("tam biet",           22),
        ("bye",                20),
        ("bai",                18),
        ("see you",            20),
        ("hen gap",            22),
        ("hen gap lai",        25),
        ("good bye",           22),
        ("goodbye",            22),
        ("tam biet ban",       25),
    ],

    "chat": [
        # ── Thanks ────────────────────────────────────────────────────────────
        ("cam on",             25),
        ("thank",              20),
        ("thanks",             20),
        # ── Simple affirmations ───────────────────────────────────────────────
        ("ok",                  8),
        ("oke",                 8),
        ("ok roi",             22),
        ("hieu roi",           22),
        ("dung roi",           22),
        ("got it",             18),
        ("vang",               20),
        ("yep",                15),
        ("yeah",               15),
        ("alright",            18),
        ("sure",               15),
        ("noted",              20),
        ("understood",         20),
        ("copy",               15),
        ("roger",              15),
        # ── Positive reactions ────────────────────────────────────────────────
        ("tuyet",              20),
        ("tuyet voi",          22),
        ("tot lam",            20),
        ("hay qua",            20),
        ("ngon",               18),
        ("xin",                12),
        ("perfect",            20),
        ("great",              18),
        ("nice",               15),
        ("awesome",            18),
        ("wow",                15),
        ("uu viet",            18),
        # ── State / emotion ───────────────────────────────────────────────────
        ("ban khoe",           22),
        ("khoe khong",         22),
        ("ban oi",             15),
        ("met qua",            20),
        ("buon qua",           20),
        ("chan qua",           18),
        ("stress qua",         18),
        ("kho qua",            18),
        ("sao vay",            18),
        ("sao the",            18),
        ("tai sao",            15),
        # ── Complaints / errors ───────────────────────────────────────────────
        ("loi roi",            20),
        ("bi loi",             20),
        ("gap loi",            20),
        ("khong chay duoc",    22),
        ("sao khong chay",     22),
        ("bi hong",            20),
        ("khong hoat dong",    22),
        ("sao vay ban",        22),
        ("help me",            15),
        # ── Polite filler that doesn't map elsewhere ──────────────────────────
        ("xin loi",            20),
        ("sorry",              15),
        ("pardon",             12),
        # ── Frustration at the assistant (de-escalating reply, _build_message) ─
        ("stupid",             22),
        ("dumb",               20),
        ("useless",            22),
        ("garbage",            20),
        ("trash",              18),
        ("suck",               18),
        ("sucks",              18),
        ("ngu",                20),
        ("vo dung",            22),
        ("te qua",             20),
        ("qua te",             20),
    ],

    "help": [
        # ── "What is X?" ──────────────────────────────────────────────────────
        ("la gi",              25),
        ("la cai gi",          25),
        ("nghia la gi",        25),
        ("what is",            25),
        ("what are",           22),
        ("what does",          22),
        ("batchout la gi",     30),
        ("parasync la gi",     28),
        ("loadfamily la gi",   28),
        # ── "How to do X?" ────────────────────────────────────────────────────
        ("lam gi",             22),
        ("lam nhu the nao",    25),
        ("lam the nao",        25),
        ("lam sao",            22),
        ("lam sao de",         25),
        ("bang cach nao",      25),
        ("cach nao",           20),
        ("how to",             22),
        ("how do",             20),
        ("how does",           20),
        ("how can",            18),
        ("xuat bang cach nao", 28),
        ("lam sao de xuat",    28),
        ("lam sao de mo",      28),
        # ── "What can you do?" ────────────────────────────────────────────────
        ("capabilities query", 30),   # from _ABBREVS
        ("lam duoc gi",        25),
        ("dung duoc gi",       25),
        ("ho tro gi",          25),
        ("co tinh nang gi",    25),
        ("ban co the",         20),
        ("ban biet lam gi",    25),
        ("ban giup duoc gi",   25),
        ("ban lam gi",         22),
        ("ban la ai",          25),
        ("ban ten gi",         25),
        ("may la ai",          22),
        ("tool nay",           18),
        # ── "Help/guide me" ───────────────────────────────────────────────────
        ("huong dan",          22),
        ("chi dan",            22),
        ("chi toi",            20),
        ("chi cach",           22),
        ("chi cho toi",        25),
        ("giai thich",         22),
        ("explain",            20),
        ("guide",              18),
        ("huong dan su dung",  28),
        ("cach su dung",       25),
        ("cach dung",          22),
        ("muon biet",          20),
        ("can biet",           20),
        ("cho toi biet",       25),
        ("noi cho toi",        22),
        ("tell me",            18),
        ("giup toi voi",       22),
        ("giup voi",           18),
        # ── Features / documentation ──────────────────────────────────────────
        ("tinh nang",          20),
        ("chuc nang",          20),
        ("su dung",            15),
        ("ho tro",             18),
        ("khai niem",          18),
        ("mo ta",              15),
        ("tai lieu",           18),
        ("document",           12),
        ("help",               12),
        ("info",               12),
    ],
}

_PENALTIES = {
    "export_direct":             [("open", -15), ("mo", -15), ("batchout", -5)],
    "open_batchout":             [("export", -8)],
    "open_batchout_configured":  [("export", -8)],
    # Create/delete verbs: the user wants to WRITE notes, not proofread
    "check_spelling":            [("tao", -25), ("create", -25), ("them", -20),
                                  ("add", -20), ("xoa", -25), ("delete", -25),
                                  ("draw", -20), ("ve", -20)],
}

_THRESHOLDS = {
    "open_batchout":             18,
    "export_direct":             18,
    "open_batchout_configured":  25,   # needs both open+batchout+params
    "open_loadfamily":           18,
    # ≥ 28 so "text note"(12) + "tieng anh"(15) alone can't fire it
    "check_spelling":            28,
    "greet":                     18,
    "chat":                      18,
    "help":                      18,
}


# ─── Conversational-input gate ────────────────────────────────────────────────
# Safety gate: small talk is never learned, matched against learned patterns,
# or allowed to launch a tool, whatever an LLM hallucinated.

# Any of these words present → a command, never small talk.
_COMMAND_WORDS = {
    "batchout", "parasync", "loadfamily", "loadfamilycloud", "projectname",
    "workset", "dimtext", "upperdimtext", "resetoverrides", "grids", "grid",
    "export", "open", "print", "sheet", "sheets", "family", "luoi", "truc",
    "pdf", "dwg", "dwf", "dgn", "ifc", "nwd", "img", "in",
    "capabilities", "query",   # capability-question marker
    "spelling", "spellcheck", "grammar", "typo", "typos", "textnote",
}


def _build_conversational_words():
    """Small-talk lexicon, derived from the greet/chat triggers plus extras."""
    words = set()
    for intent in ("greet", "chat"):
        for feat, _w in _TRIGGERS[intent]:
            words.update(feat.split())
    words.update({
        "morning", "afternoon", "evening", "night", "yo", "sup",
        "gm", "gn", "haha", "hihi", "lol", "ban", "buoi",
        "please", "welcome", "greetings",
    })
    return words - _COMMAND_WORDS


_CONVERSATIONAL_WORDS = _build_conversational_words()


def is_conversational(user_input):
    """True for pure small talk ("morning", "cảm ơn nhé"). Conservative: any
    command word, unknown word, or > 6 tokens → False."""
    if not user_input or not user_input.strip():
        return False
    expanded = _expand(_norm(user_input))
    clean = re.sub(r'[^a-z0-9\s]', ' ', expanded)
    tokens = clean.split()
    if not tokens or len(tokens) > 6:
        return False
    for t in tokens:
        if t in _COMMAND_WORDS:
            return False
    meaningful = [t for t in tokens if t not in _STOPWORDS and len(t) >= 2]
    if not meaningful:
        return True
    return all(t in _CONVERSATIONAL_WORDS for t in meaningful)


# ─── Slot extraction ──────────────────────────────────────────────────────────

_FORMATS = ["dwg", "dwf", "dgn", "ifc", "nwd", "img", "pdf"]  # pdf last = default

_FORMAT_LETTERS = {"PDF", "DWG", "DWF", "DGN", "IFC", "NWD", "IMG"}


def _extract_slots(raw):
    """Return {format, filter, combine} from the original (unicode) text."""
    normed = _norm(raw)

    fmt = "pdf"
    for f in _FORMATS:
        if f in normed.split() or (" " + f) in normed or (f + " ") in normed:
            fmt = f
            break

    # Sheet-prefix filter (one letter, not a format). 1: "G sheet" / "tờ G"
    sheet_filt = ""
    m = re.search(
        r'\b([A-Z])\s*[-–]?\s*(?:sheet|to|to\s|ban\s*ve|sheets)\b',
        raw, re.IGNORECASE
    )
    if m and m.group(1).upper() not in _FORMAT_LETTERS:
        sheet_filt = m.group(1).upper()

    if not sheet_filt:
        m = re.search(r'(?:sheet|to|sheets)\s+([A-Z])\b', raw, re.IGNORECASE)
        if m and m.group(1).upper() not in _FORMAT_LETTERS:
            sheet_filt = m.group(1).upper()

    # 2: a lone single-letter token
    if not sheet_filt:
        for token in raw.split():
            clean_tok = "".join(c for c in token if c.isalnum())
            if len(clean_tok) == 1 and clean_tok.isalpha():
                tok = clean_tok.upper()
                if tok not in _FORMAT_LETTERS:
                    sheet_filt = tok
                    break

    combine_kws = ["combine", "merge", "gop", "ghep", "1 file", "mot file"]
    combine = any(k in normed for k in combine_kws)

    return {"format": fmt, "filter": sheet_filt, "combine": combine}


# ─── Context / pronoun resolution ─────────────────────────────────────────────

# Pronouns for the last-mentioned tool, split by shape (see _is_pronoun_query):
# a word is anaphoric only when no noun follows it; a phrase is on its own.
_PRONOUN_WORDS   = {"no", "it", "that", "this"}
_PRONOUN_PHRASES = ("cai nay", "cai do", "cai ay", "no ay",
                    "tool nay", "tool do", "tool ay")

# Intent → label for dedicated launchers only; the rest come from the live
# registry via _tool_label() (a hardcoded list once named deleted tools).
_TOOL_LABELS = {
    "open_batchout":            "BatchOut",
    "export_direct":            "BatchOut",
    "open_batchout_configured": "BatchOut",
    "open_loadfamily":          "Family Loader",
}


def _tool_label(intent):
    """Display label for an intent — fixed map first, then the live registry."""
    if intent in _TOOL_LABELS:
        return _TOOL_LABELS[intent]
    for t in _tool_catalog():
        if t.get('intent') == intent:
            return t.get('title') or intent
    return intent


# ─── Deterministic tool resolver ─────────────────────────────────────────────
# Ranks the text against EVERY known tool and answers only when one clearly wins
# (a first-substring match once sent "mở mcp control" to CAD to Elements).

# Tools with a dedicated launcher (_SPECIAL_LAUNCHERS), not auto-discovered:
#   (intent, title, joined-name, [aliases: folder, XAML, old names], description)
# Everything else comes from tool_discovery in _tool_catalog(); add nothing here
# without a special launcher (2026-07-28: deleted tools kept being "opened").
_BUILTIN_TOOLS = [
    ("open_batchout", "BatchOut", "batchout",
     ["Batch Out", "Export Manager"],
     u"Xuất sheet hàng loạt sang PDF / DWG / DWF / IFC (batch export sheets)"),
    ("open_loadfamily", "Family Loader", "loadfamily",
     ["Load Family", "Family Manager", "ManaFami"],
     u"Tải family từ thư viện vào project (load family from library)"),
]

# Open verbs (post-_expand). Verbs _ABBREVS can't rewrite globally ("chạy" would
# break "không chạy được") live here. "dung" is absent on purpose: folded it is
# also "dựng" (build), which hijacked every modeling request.
_OPEN_VERBS = {"open", "launch", "start", "run", "show",
               "chay", "xai"}


def _singularise(w):
    """Fold trivial English plurals so 'views' matches 'view' etc."""
    if len(w) >= 4 and w.endswith("s") and not w.endswith("ss"):
        return w[:-1]
    return w


def _camel_split(text):
    """'DWGManagement' → 'DWG Management' — word boundaries for camel names."""
    return re.sub(r'(?<=[a-z0-9])(?=[A-Z])', ' ', text)


def _name_variants(text):
    """(joined_names, word_set) for ONE name over raw, camel-split and _expand()
    forms ("ImageToDrafting" also matches "img to drafting")."""
    joined, words = set(), set()
    for base in (text, _camel_split(text)):
        for form in (_norm(base), _expand(_norm(base))):
            j = re.sub(r'[^a-z0-9]', '', form)
            if j:
                joined.add(j)
            for w in re.findall(r'[a-z0-9]+', form):
                if w not in _STOPWORDS and len(w) >= 2:
                    words.add(_singularise(w))
    return joined, words


# Implementation words in developer docstrings ("pyRevit WPFWindow modeless")
# that must not match as capability topics.
_DESC_NOISE = {
    "py", "pyrevit", "ironpython", "wpf", "wpfwindow", "window", "modeless",
    "modal", "script", "dialog", "form", "gui", "ui", "xaml", "class",
    "module", "version", "tool", "revit", "button", "pushbutton", "panel",
    "dqt", "t3lab", "todo", "wip", "deprecated",
}

# Prose glue not in _STOPWORDS (Vietnamese "in" = print) but noise in descriptions.
_DESC_PREPS = {"in", "into", "onto", "over", "under", "per", "via", "within"}


def _desc_words(desc):
    """Word set of a description (capability Q&A), minus jargon and prose glue."""
    out = set()
    for w in re.findall(r'[a-z0-9]+', _norm(_camel_split(desc or ''))):
        if (len(w) < 2 or w in _STOPWORDS or w in _DESC_NOISE
                or w in _DESC_PREPS or re.match(r'^v?\d+$', w)):
            continue
        out.add(_singularise(w))
    return out


def _tool_entry(intent, title, names, desc='', panel='', extra_words=None):
    """One catalog entry. `names` = every name the tool goes by (title, folder,
    XAML, aliases): per-name variants for exact matching, a word union for fuzzy."""
    joined_all, variants, union = set(), [], set()
    for name in names:
        if not name:
            continue
        j, w = _name_variants(name)
        joined_all |= j
        if w and frozenset(w) not in variants:
            variants.append(frozenset(w))
        union |= w
    # Evidence tiers: identity (names + registry keywords) identifies the tool;
    # topic (description words) is only weak evidence.
    return {'intent': intent, 'title': title, 'desc': (desc or '').strip(),
            'panel': panel, 'joined': joined_all, 'variants': variants,
            'words': union, 'kw_words': set(extra_words or ()),
            'desc_words': _desc_words(desc)}


# Too costly to rebuild several times per turn; keyed on a registry signature.
_CATALOG_CACHE = {"sig": None, "value": None}


def _tool_catalog():
    """Every known tool (builtin + auto-discovered), all names unified per tool."""
    try:
        from Services.tool_discovery import get_registered_tools
        tools = get_registered_tools()
    except Exception:
        tools = []

    sig = tuple(sorted(u"{}|{}".format(t.get('intent') or '',
                                       t.get('title') or '') for t in tools))
    if _CATALOG_CACHE["value"] is not None and _CATALOG_CACHE["sig"] == sig:
        return _CATALOG_CACHE["value"]

    catalog = []
    for intent, title, joined, aliases, desc in _BUILTIN_TOOLS:
        e = _tool_entry(intent, title, [title] + list(aliases), desc,
                        panel=u"Core")
        e['joined'].add(joined)
        catalog.append(e)
    for t in tools:
        # Ampersand-stripped form is for matching only (displayed "Visual   Styles")
        raw_title = (t.get('title') or '').strip()
        title = raw_title.replace('&amp;', ' ').replace('&', ' ').strip()
        btn   = re.sub(r'\.(push|smart|url)button$', '', t.get('button') or '')
        # `aliases` carries the ribbon label when it differs from __title__
        names = ([raw_title, title, btn] + list(t.get('xaml') or [])
                 + list(t.get('aliases') or []))
        kw_words = set()
        for kw in (t.get('keywords') or []):
            for w in re.findall(r'[a-z0-9]+', _norm(kw)):
                if w not in _STOPWORDS and len(w) >= 2:
                    kw_words.add(_singularise(w))
        e = _tool_entry(t.get('intent'), raw_title or btn, names,
                        t.get('doc') or '',
                        panel=(t.get('panel') or '').replace('.panel', ''),
                        extra_words=kw_words)
        if e['words'] or e['joined']:
            catalog.append(e)

    _CATALOG_CACHE["sig"] = sig
    _CATALOG_CACHE["value"] = catalog
    return catalog


# Mutually exclusive words, one group per axis: query and tool naming different
# members is a contradiction (fuzzy overlap offered "Create Plan Views" for
# "create 3d view").
_CONTRADICTION_GROUPS = [
    frozenset(['3d', 'plan', 'section', 'elevation', 'schedule', 'legend',
               'sheet']),
    frozenset(['wall', 'floor', 'ceiling', 'roof', 'column', 'beam', 'room',
               'door', 'window']),
    frozenset(['pdf', 'dwg', 'ifc', 'nwc', 'image']),
]


def _contradicts(qset, tool_words):
    """True when query and tool name DIFFERENT members of one group."""
    for group in _CONTRADICTION_GROUPS:
        q = qset & group
        t = tool_words & group
        if q and t and not (q & t):
            return True
    return False


def resolve_tool(user_input, exact_only=False):
    """Resolve a tool-open request against the full catalog.

    Returns (match, candidates): match when one tool clearly wins, else up to 3
    candidates. Exact = all tokens minus open verbs, KEEPING stopwords (so "mcp
    control la gi" is not exact and goes to help), join to a tool's joined name
    or equal a name's word set in any order. Fuzzy needs ≥ 0.75 and a ≥ 0.2 lead.
    """
    if not user_input or not user_input.strip():
        return None, []
    expanded = _expand(_norm(user_input))
    clean = re.sub(r'[^a-z0-9\s]', ' ', expanded)
    tokens_all = [w for w in clean.split() if w not in _OPEN_VERBS]
    if not tokens_all or len(tokens_all) > 8:
        return None, []
    qjoined_all = "".join(tokens_all)
    qset_all    = set(_singularise(w) for w in tokens_all)
    qset = set(_singularise(w) for w in tokens_all if w not in _STOPWORDS)

    scored = []
    for tool in _tool_catalog():
        exact = (qjoined_all in tool['joined']
                 or any(qset_all == v for v in tool['variants'])
                 or (len(tokens_all) == 1 and tokens_all[0] in tool['joined']))
        if exact:
            scored.append((1.0, True, tool))
            continue
        if _contradicts(qset, tool['words']):
            continue
        inter = qset & tool['words']
        if inter:
            cov_tool  = len(inter) / float(len(tool['words']))
            cov_query = len(inter) / float(len(qset))
            # Fuzzy can reach 1.0 ("batchout là gì") but is still not exact
            scored.append((0.6 * cov_tool + 0.4 * cov_query, False, tool))

    if not scored:
        return None, []
    scored.sort(key=lambda x: (-x[0], not x[1]))
    top_score, top_exact, top = scored[0]
    runner_up = scored[1][0] if len(scored) > 1 else 0.0

    if top_exact:
        # Two tools both exact (shared XAML alias) → ambiguous
        exact_intents = set(t['intent'] for s, e, t in scored if e)
        if len(exact_intents) > 1:
            return None, [t for s, e, t in scored if e][:3]
        return top, []
    if exact_only:
        return None, []
    if top_score >= 0.75 and (top_score - runner_up) >= 0.2:
        return top, []
    candidates = [t for s, e, t in scored[:3] if s >= 0.45]
    return None, candidates


def _has_open_verb(expanded):
    """True if the expanded text contains an explicit open verb."""
    return bool(set(expanded.split()) & _OPEN_VERBS)


# ─── Capability questions ─────────────────────────────────────────────────────
# "Có tool nào để X không?" is answered from the REAL catalog, never by the LLM.

_CAP_RES = [
    re.compile(r'\bco\s+(?:tool|cong cu|chuc nang|tinh nang|lenh)\b'),
    re.compile(r'\b(?:tool|cong cu|lenh)\s+nao\b'),
    re.compile(r'\b(?:tool|cong cu)\s+(?:de|giup|ho tro|cho)\b'),
    re.compile(r'\btool\s+(?:for|to)\b'),
    re.compile(r'\bdo\s+you\s+have\b'),
    re.compile(r'\bis\s+there\s+(?:a|an|any)\b'),
    re.compile(r'\b(?:which|what)\s+tool\b'),
    re.compile(r'\bany\s+tool\b'),
    re.compile(r'\bhave\s+a\s+tool\b'),
    # Generic "what can you do" (EN forms collapsed by _ABBREVS)
    re.compile(r'\bcapabilities query\b'),
    re.compile(r'\blam duoc gi\b'),
    re.compile(r'\bdung duoc gi\b'),
    re.compile(r'\bho tro gi\b'),
    re.compile(r'\bbiet lam gi\b'),
    re.compile(r'\bgiup duoc gi\b'),
    # Asking for the LIST: missed, the LLM answers from memory with a truncated list
    re.compile(r'\b(?:liet ke|danh sach|list)\s+(?:cac\s+|nhung\s+|the\s+)?'
               r'(?:tool|cong cu|lenh|chuc nang|tinh nang)\b'),
    re.compile(r'\b(?:chuc nang|tinh nang)\s+(?:cua|of)\s+'
               r'(?:assistant|t3lab|extension|ban|you)\b'),
    re.compile(r'\bgioi thieu\s+(?:cac\s+|nhung\s+)?(?:tool|chuc nang|tinh nang)\b'),
    re.compile(r'\b(?:extension|t3lab|assistant)\s+(?:nay\s+)?co\s+gi\b'),
    # Forms _ABBREVS can't cover ("what ELSE can you do"), anchored on the verb:
    # not "what can you tell me about walls" / "what do you think"
    re.compile(r'\b(?:what|which)\s+(?:\w+\s+){0,2}(?:can|could)'
               r'\s+(?:you|u|t3lab|this|it)\s+do\b'),
    re.compile(r'\bwhat\s+are\s+your\s+(?:capabilit|feature|function|skill)'),
    re.compile(r'\b(?:list|show)\s+(?:me\s+)?(?:(?:your|all|the|my)\s+){1,2}'
               r'(?:tool|feature|capabilit|function|skill)'),
    # Vietnamese, subject-anchored so "tôi phải làm gì bây giờ" stays out
    re.compile(r'\bco\s+the\s+lam\s+(?:duoc\s+)?gi\b'),
    re.compile(r'\b(?:ban|t3lab|assistant)\s+(?:co\s+the\s+)?'
               r'(?:lam|ho tro|giup|xu ly)\s+(?:duoc\s+)?(?:nhung\s+)?gi\b'),
]

# Question boilerplate stripped before matching the FUNCTION words
_CAP_BOILERPLATE = {
    "tool", "cong", "cu", "chuc", "nang", "tinh", "lenh", "nao",
    "have", "there", "which", "what", "any", "capabilities", "query",
    "biet", "ung", "dung", "ban", "assistant", "t3lab", "co", "khong",
    "gi", "giup", "ho", "tro", "lam", "duoc", "thuc", "hien", "does",
    "function", "feature", "help", "the",
    # Plural / "things" / "else" particles: alone they became the predicate of
    # "bạn có thể làm được những gì" → "no dedicated tool"
    "nhung", "cac", "viec", "khac", "nua", "nhi", "con",
}

# ─── Semantic roles inside a question ────────────────────────────────────────
# Capability question = FRAME("có tool nào để … không") PREDICATE("xuất pdf")
# [SCOPE("with this project")]. Only the PREDICATE may select a tool ("what can
# you do with this project" once answered "Family Loader"). Scope is stripped on
# the QUERY side only; for tools "model"/"project" stay meaningful.
_SCOPE_PHRASES = (
    # Stripped before tokenising ("tại" in "hiện tại" vs "tải" = load)
    "hien tai", "hien nay", "luc nay", "bay gio", "o day", "trong nay",
    "du an", "right now", "at the moment",
)

_SCOPE_WORDS = {
    # environment nouns
    "project", "model", "file", "revit", "document", "doc",
    # demonstratives / deixis
    "this", "these", "that", "those", "here", "current",
    "nay", "day", "kia",
}

# Greeting / politeness openers (a lone "hi" became the predicate of "hi. what
# can you do?"). Narrow on purpose: only words that never name an operation
# ("copy" stays a predicate).
_SOCIAL_WORDS = {
    "hi", "hii", "hey", "helo", "hello", "hallo", "halo", "howdy", "yo", "sup",
    "hiya", "greetings", "morning", "afternoon", "evening", "night",
    "chao", "alo", "bye", "goodbye", "thanks", "thank", "thanx", "thx",
    "please", "pls", "sorry", "excuse", "oi", "nhe", "nha", "a", "oke", "okay",
}

# A pronoun followed by one of these (or a scope noun) is a determiner, not anaphora.
_DETERMINERS = {
    "the", "a", "an", "this", "that", "these", "those",
    "my", "our", "your", "its", "cac", "nhung",
}


def _predicate_words(expanded):
    """Content words naming the FUNCTION asked about (frame, open verbs,
    stopwords, greetings and scope stripped from `_expand`ed text). Empty = the
    generic "what can you do?"."""
    padded = u" " + u" ".join(
        re.sub(r'[^a-z0-9\s]', ' ', expanded).split()) + u" "
    for phrase in _SCOPE_PHRASES:
        padded = padded.replace(u" " + phrase + u" ", u" ")
    tokens = padded.split()

    out = set()
    for i, w in enumerate(tokens):
        if (len(w) < 2 or w in _STOPWORDS or w in _CAP_BOILERPLATE
                or w in _SCOPE_WORDS or w in _OPEN_VERBS
                or w in _SOCIAL_WORDS):
            continue
        # "in" = Vietnamese print, unless a determiner/scope noun follows
        # (English preposition: "in this model")
        if w == "in":
            nxt = tokens[i + 1] if i + 1 < len(tokens) else None
            if nxt and (nxt in _DETERMINERS or nxt in _SCOPE_WORDS):
                continue
        out.add(_singularise(w))
    return out


def is_capability_question(expanded):
    """True if the (expanded) input asks whether a tool/feature exists."""
    clean = re.sub(r'[^a-z0-9\s]', ' ', expanded)
    padded = u" " + u" ".join(clean.split()) + u" "
    return any(r.search(padded) for r in _CAP_RES)


# ─── Practical capability overview (grouped by workflow, verified live) ───────
# "What can you do?" = a CURATED map of MCP tools (core/server.py), validated
# against the live registry at render time: a line whose tools are all gone is
# dropped (drift trap, Intelligence/t3lab_agent.py); unmapped tools show under
# "Khác/Other". Notes were checked against the core/server.py schemas.

def _live_tool_names():
    """Registered MCP tool names; empty set when unreachable (offline tests)."""
    try:
        from Intelligence.t3lab_agent import _get_mcp_tools
        return set(t.get("name", "") for t in _get_mcp_tools())
    except Exception:
        return set()


def _live_skills():
    """[(name, description)] of enabled assistant skills, or [] if none/error."""
    try:
        from Intelligence.skills_engine import get_skills_engine
        return [(s.get("name") or s.get("id"), s.get("description") or u"")
                for s in get_skills_engine().get_catalog(enabled_only=True)]
    except Exception:
        return []


def _user_display_name():
    """User's set display name, or '' (no fallback — used only to personalise)."""
    try:
        from config.user_profile import get_profile
        return (get_profile().get_name(fallback=False) or u"").strip()
    except Exception:
        return u""


# Plumbing tools kept out of the overview and its "Other" catch-all.
_CAP_HIDE = frozenset([
    u"say_hello", u"show_assistant_pane", u"file_watcher_status",
])

# Sections: (emoji, vi_title, en_title, lines); line: (names, vi_md, en_md).
# A line renders offline or when one of its names is registered; every name
# counts as "covered" for the anti-drift catch-all.
_CAP_MAP = [
    (u"📖", u"ĐỌC / QUERY", u"READ / QUERY", [
        ((u"ai_element_filter",),
         u"`ai_element_filter` — lọc element theo category, group, workset, tầng, type, tên, parameter (gộp trong 1 lần gọi); `group_by` để đếm/thống kê, `store_as` để thao tác hàng loạt không cần chuyển id",
         u"`ai_element_filter` — filter elements by category, group, workset, level, type, name, parameter (all in one call); `group_by` for counts, `store_as` for bulk actions without ferrying ids"),
        ((u"get_all_parameters", u"get_parameter", u"revit_get_element_info"),
         u"`get_all_parameters` / `get_parameter` / `revit_get_element_info` — đọc mọi param & thông tin của 1 element",
         u"`get_all_parameters` / `get_parameter` / `revit_get_element_info` — read every parameter & info of an element"),
        ((u"get_elements_by_level", u"get_current_view_elements", u"get_current_view_info", u"revit_get_active_view", u"get_element_bounding_box"),
         u"`get_elements_by_level`, `get_current_view_elements`, `get_current_view_info` / `revit_get_active_view`, `get_element_bounding_box`",
         u"`get_elements_by_level`, `get_current_view_elements`, `get_current_view_info` / `revit_get_active_view`, `get_element_bounding_box`"),
        ((u"revit_get_selected_elements", u"select_elements", u"get_available_family_types"),
         u"`revit_get_selected_elements`, `select_elements` (chọn theo category+filter, zoom tới được), `get_available_family_types`",
         u"`revit_get_selected_elements`, `select_elements` (select by category+filter, can zoom to them), `get_available_family_types`"),
        ((u"get_schedule_data", u"get_material_quantities"),
         u"`get_schedule_data` — đọc schedule ra JSON (row_count + column_totals chính xác), `get_material_quantities`",
         u"`get_schedule_data` — read a schedule as JSON (exact row_count + column_totals), `get_material_quantities`"),
        ((u"list_levels", u"revit_list_views", u"revit_list_sheets", u"list_worksets"),
         u"`list_levels`, `revit_list_views`, `revit_list_sheets`, `list_worksets`",
         u"`list_levels`, `revit_list_views`, `revit_list_sheets`, `list_worksets`"),
        ((u"get_revit_context", u"revit_get_project_info", u"analyze_model_statistics"),
         u"`get_revit_context`, `revit_get_project_info`, `analyze_model_statistics` — tổng quan model",
         u"`get_revit_context`, `revit_get_project_info`, `analyze_model_statistics` — model overview"),
    ]),
    (u"✏️", u"TẠO / SỬA", u"CREATE / MODIFY", [
        ((u"place_wall", u"create_surface_based_element"),
         u"**Walls / Floors / Ceilings / Roofs** — `place_wall`, `create_surface_based_element`",
         u"**Walls / Floors / Ceilings / Roofs** — `place_wall`, `create_surface_based_element`"),
        ((u"create_point_based_element",),
         u"**Doors / Windows / Furniture** — `create_point_based_element` (host được vào wall)",
         u"**Doors / Windows / Furniture** — `create_point_based_element` (can be hosted into a wall)"),
        ((u"create_line_based_element", u"create_structural_framing_system"),
         u"**Beams / Pipes / Ducts** — `create_line_based_element`, `create_structural_framing_system`",
         u"**Beams / Pipes / Ducts** — `create_line_based_element`, `create_structural_framing_system`"),
        ((u"create_grid", u"create_level", u"create_room", u"load_family", u"room_to_floor"),
         u"**Grids / Levels / Rooms** — `create_grid` (tự đặt label A/B/C · 1/2/3), `create_level`, `create_room`, `load_family`, `room_to_floor`",
         u"**Grids / Levels / Rooms** — `create_grid` (auto labels A/B/C · 1/2/3), `create_level`, `create_room`, `load_family`, `room_to_floor`"),
        ((u"move_elements", u"copy_elements", u"rotate_element", u"split_element", u"split_curve"),
         u"`move_elements`, `copy_elements`, `rotate_element`, `split_element` / `split_curve` (giữ đúng arc/spline, không flatten)",
         u"`move_elements`, `copy_elements`, `rotate_element`, `split_element` / `split_curve` (keeps arcs/splines exact, no flattening)"),
        ((u"edit_elements",),
         u"`edit_elements` — mirror (lật đối xứng), đổi type hàng loạt, group / ungroup",
         u"`edit_elements` — mirror, bulk type swap (change_type), group / ungroup"),
        ((u"join_geometry", u"delete_element", u"operate_element"),
         u"`join_geometry` (join/unjoin), `delete_element` (có `dry_run` xem trước), `operate_element` (hide / isolate / bỏ isolate / ẩn cả category / pin / unpin / halftone / trong suốt / chọn tương tự / reset màu)",
         u"`join_geometry` (join/unjoin), `delete_element` (`dry_run` preview), `operate_element` (hide / isolate / reset temporary / hide a whole category / pin / unpin / halftone / transparency / select similar / reset color)"),
        ((u"bulk_set_parameter", u"set_parameter", u"rename_element"),
         u"`bulk_set_parameter` — set 1 param cho hàng trăm element cùng lúc (filter được); `set_parameter`, `rename_element` cho từng element",
         u"`bulk_set_parameter` — set one parameter across hundreds of elements at once (filterable); `set_parameter`, `rename_element` for a single element"),
        ((u"create_project_parameter", u"create_workset", u"set_element_workset"),
         u"`create_project_parameter`, `create_workset`, `set_element_workset`",
         u"`create_project_parameter`, `create_workset`, `set_element_workset`"),
    ]),
    (u"📁", u"VIEW / SHEET / ANNOTATION", u"VIEW / SHEET / ANNOTATION", [
        ((u"create_view", u"duplicate_view", u"apply_view_template", u"create_view_filter", u"set_active_view"),
         u"`create_view` (plan sàn / trần / kết cấu / area, 3D, section, elevation, drafting, legend), `duplicate_view` (plain / with detailing / dependent), `apply_view_template`, `create_view_filter`, `set_active_view`",
         u"`create_view` (floor / ceiling / structural / area plan, 3D, section, elevation, drafting, legend), `duplicate_view` (plain / with detailing / dependent), `apply_view_template`, `create_view_filter`, `set_active_view`"),
        ((u"manage_sheet", u"manage_revision"),
         u"`manage_sheet` — nhân bản sheet, đánh lại số sheet, xem print set; `manage_revision` — liệt kê / tạo revision, gán revision vào sheet, đánh dấu đã phát hành",
         u"`manage_sheet` — duplicate sheets, renumber, list print sets; `manage_revision` — list / create revisions, assign to sheets, mark issued"),
        ((u"manage_material", u"create_detail_annotation", u"manage_document"),
         u"`manage_material` — liệt kê vật liệu, xem vật liệu của element; `create_detail_annotation` — filled region / detail line; `manage_document` — save, save as, sync with central (sync mặc định TẮT)",
         u"`manage_material` — list materials, read an element's materials; `create_detail_annotation` — filled region / detail line; `manage_document` — save, save as, sync with central (sync is OFF by default)"),
        ((u"manage_links",),
         u"`manage_links` — liệt kê link Revit & CAD, reload / unload / xoá / pin",
         u"`manage_links` — list Revit & CAD links, reload / unload / delete / pin"),
        ((u"manage_view", u"manage_view_template"),
         u"`manage_view` — đổi scale / detail level / discipline / crop box; `manage_view_template` — liệt kê, đếm view đang dùng, đổi tên, nhân bản, xoá view template",
         u"`manage_view` — set scale / detail level / discipline / crop box; `manage_view_template` — list, usage count, rename, duplicate, delete view templates"),
        ((u"create_schedule", u"create_sheet", u"place_views_on_sheets", u"add_view_to_sheet"),
         u"`create_schedule`, `create_sheet`, `place_views_on_sheets`, `add_view_to_sheet`",
         u"`create_schedule`, `create_sheet`, `place_views_on_sheets`, `add_view_to_sheet`"),
        ((u"create_dimension", u"create_text_note", u"tag_elements", u"tag_all_rooms", u"tag_all_walls"),
         u"`create_dimension`, `create_text_note`, `tag_elements` / `tag_all_rooms` / `tag_all_walls`",
         u"`create_dimension`, `create_text_note`, `tag_elements` / `tag_all_rooms` / `tag_all_walls`"),
        ((u"color_elements", u"revit_override_color"),
         u"`color_elements` — tô màu theo giá trị parameter (rất hay để QA); `revit_override_color`",
         u"`color_elements` — color-code by parameter value (great for QA); `revit_override_color`"),
    ]),
    (u"📤", u"EXPORT", u"EXPORT", [
        ((u"export_sheets_pdf", u"export_dwg", u"export_image", u"export_room_data"),
         u"`export_sheets_pdf` (gộp thành 1 file được), `export_dwg`, `export_image` (PNG), `export_room_data`",
         u"`export_sheets_pdf` (can combine into one file), `export_dwg`, `export_image` (PNG), `export_room_data`"),
        ((u"export_model",),
         u"`export_model` — IFC, NWC (Navisworks), DWF, DGN",
         u"`export_model` — IFC, NWC (Navisworks), DWF, DGN"),
    ]),
    (u"🔍", u"QA / MODEL HEALTH", u"QA / MODEL HEALTH", [
        ((u"check_bad_geometry",),
         u"`check_bad_geometry` — tìm hình học lỗi (sliver, zero-area, singular point) làm crash export PDF/DWG",
         u"`check_bad_geometry` — find the degenerate geometry (slivers, zero-area faces, singular points) that crashes PDF/DWG export"),
        ((u"audit_model", u"get_model_warnings", u"get_model_health", u"analyze_model_statistics"),
         u"`audit_model`, `get_model_warnings`, `get_model_health`, `analyze_model_statistics`",
         u"`audit_model`, `get_model_warnings`, `get_model_health`, `analyze_model_statistics`"),
        ((u"purge_unused",),
         u"`purge_unused` — mặc định dry-run (chỉ báo cáo); set `dry_run=false` mới thật sự xoá",
         u"`purge_unused` — dry-run by default (report only); set `dry_run=false` to actually purge"),
    ]),
    (u"⚡", u"ADVANCED", u"ADVANCED", [
        ((u"send_code_to_revit",),
         u"`send_code_to_revit` — chạy **IronPython trực tiếp** trong Revit context (full API) → mạnh nhất, làm được cả những gì tool có sẵn chưa cover",
         u"`send_code_to_revit` — run **IronPython directly** in the Revit context (full API) → the most powerful path, does what the built-in tools don't cover yet"),
        ((u"store_project_data", u"store_room_data", u"query_stored_data"),
         u"Lưu / truy vấn dữ liệu bền vững trong model: `store_project_data`, `store_room_data`, `query_stored_data`",
         u"Persist / query data inside the model: `store_project_data`, `store_room_data`, `query_stored_data`"),
        ((u"switch_active_document", u"open_document", u"close_document", u"list_open_documents", u"list_recent_documents"),
         u"Làm việc trên nhiều model đang mở: `switch_active_document`, `open_document`, `close_document`, `list_open_documents`, `list_recent_documents`",
         u"Work across open models: `switch_active_document`, `open_document`, `close_document`, `list_open_documents`, `list_recent_documents`"),
    ]),
]


def _ribbon_tools_section(viet):
    """Every openable ribbon tool by panel, named in full (a bare count told the
    user nothing). Same catalog as the resolver, so the list matches what opens."""
    groups, order = {}, []
    for tool in _tool_catalog():
        panel = (tool.get('panel') or u"").strip() or (u"Khác" if viet else u"Other")
        title = (tool.get('title') or u"").strip()
        if not title:
            continue
        if panel not in groups:
            groups[panel] = []
            order.append(panel)
        if title not in groups[panel]:
            groups[panel].append(title)
    if not groups:
        return u""

    total = sum(len(v) for v in groups.values())
    head = (u"\n📂 **{} tool trên ribbon** mình mở được — gõ *mở <tên tool>*:"
            if viet else
            u"\n📂 **{} ribbon tools** I can open — type *open <tool name>*:")
    lines = [head.format(total)]
    # "Core" first, then panels alphabetically
    for panel in sorted(order, key=lambda p: (p != u"Core", p.lower())):
        lines.append(u"• **{}** — {}".format(
            panel, u", ".join(sorted(groups[panel], key=lambda s: s.lower()))))
    return u"\n".join(lines)


def _capabilities_overview(viet):
    """Workflow-grouped overview: _CAP_MAP validated live, skills, ribbon tools."""
    live = _live_tool_names()   # empty set when offline → show every line

    if viet:
        out = [u"Đây là những gì mình làm được **trực tiếp trên model Revit** "
               u"(qua T3Lab connector):"]
    else:
        out = [u"Here's what I can do **directly on the Revit model** "
               u"(via the T3Lab connector):"]

    covered = set()
    for emoji, vi_title, en_title, lines in _CAP_MAP:
        rendered = []
        for names, vi_line, en_line in lines:
            covered.update(names)
            if live and not (set(names) & live):
                continue
            rendered.append(u"• " + (vi_line if viet else en_line))
        if rendered:
            out.append(u"\n{} **{}**\n{}".format(
                emoji, vi_title if viet else en_title, u"\n".join(rendered)))

    # ── Anti-drift: registered tools the map never named ──
    if live:
        extra = sorted(live - covered - _CAP_HIDE)
        if extra:
            chips = u", ".join(u"`{}`".format(n) for n in extra)
            out.append((u"\n🧩 **Khác** (tool mới, chưa xếp nhóm): {}" if viet
                        else u"\n🧩 **Other** (newer tools, not yet grouped): {}"
                        ).format(chips))

    # ── Skills (SkillsEngine), personalised if a name is set ──
    skills = _live_skills()
    if skills:
        name = _user_display_name()
        if viet:
            hdr = (u"\n🎯 **SKILLS chuyên biệt đã cài cho {}**:".format(name)
                   if name else u"\n🎯 **SKILLS chuyên biệt đã cài**:")
        else:
            hdr = (u"\n🎯 **Custom skills set up for {}**:".format(name)
                   if name else u"\n🎯 **Custom skills installed**:")
        s_lines = [hdr]
        for sname, sdesc in skills[:12]:
            s_lines.append(u"• **{}** — {}".format(sname, sdesc) if sdesc
                           else u"• **{}**".format(sname))
        rest = skills[12:]
        if rest:
            names = u", ".join(sn for sn, _ in rest)
            s_lines.append((u"• …+{} skill nữa: {}" if viet
                            else u"• …+{} more: {}").format(len(rest), names))
        out.append(u"\n".join(s_lines))

    # ── Ribbon tools + quick affordances ──
    ribbon = _ribbon_tools_section(viet)
    if ribbon:
        out.append(ribbon)
    if viet:
        out.append(u"\nCũng có thể hỏi *có tool nào để … không?* hoặc xuất "
                   u"nhanh *xuất pdf G sheet*.")
    else:
        out.append(u"\nYou can also ask *is there a tool for …?* or quick-"
                   u"export *export pdf G sheets*.")

    return u"\n".join(out)


def answer_capability_question(user_input, viet):
    """Answer 'do you have a tool for X?' from the real catalog.

    Returns {intent, params, message, _nlu, _authoritative}; _authoritative
    means no LLM may override it.
    """
    expanded = _expand(_norm(user_input))
    pred = _predicate_words(expanded)

    # No predicate ("what can you do with this project") → full overview
    if not pred:
        msg = _capabilities_overview(viet)
        return {"intent": "help", "params": {"answer": msg}, "message": msg,
                "_nlu": True, "_authoritative": True}

    # ── Evidence-tiered matching ─────────────────────────────────────────────
    # Identity hits (name / keyword) identify a tool; topic hits (description)
    # are weak. Each hit is damped by how widely the word is shared ("manager"
    # is in 16 of 44 names); rarity may only damp evidence, never create it.
    catalog = _tool_catalog()
    df = {}
    for tool in catalog:
        for w in (tool['words'] | tool['kw_words'] | tool['desc_words']):
            df[w] = df.get(w, 0) + 1
    size = float(max(1, len(catalog)))

    def rarity(w):
        share = df.get(w, 1) / size
        if share <= 0.05:
            return 1.0
        if share <= 0.12:
            return 0.6
        if share <= 0.25:
            return 0.25
        return 0.05

    matches, near = [], []
    for tool in catalog:
        id_hits    = pred & (tool['words'] | tool['kw_words'])
        topic_hits = (pred & tool['desc_words']) - id_hits
        hits = id_hits | topic_hits
        if not hits:
            continue
        coverage = len(hits) / float(len(pred))
        evidence = (2.0 * sum(rarity(w) for w in id_hits)
                    + sum(rarity(w) for w in topic_hits))
        strong_enough = (evidence >= 0.5
                         and (id_hits or len(topic_hits) >= 2
                              or coverage >= 0.5))
        if strong_enough:
            matches.append((evidence + coverage, bool(id_hits), coverage, tool))
        else:
            near.append((coverage, tool))
    matches.sort(key=lambda x: -x[0])
    near.sort(key=lambda x: -x[0])

    # Real but too generic ("quản lý project") → the overview, not "no such tool"
    if not matches and pred and all(w in df and rarity(w) <= 0.05 for w in pred):
        msg = _capabilities_overview(viet)
        return {"intent": "help", "params": {"answer": msg}, "message": msg,
                "_nlu": True, "_authoritative": True}

    if matches:
        # "bạn có thể mở X không?": named outright + open verb → just open
        top_score, top_is_id, top_cov, top = matches[0]
        if _has_open_verb(expanded) and top_is_id and top_cov >= 0.999:
            msg = (u"Đang mở {}...".format(top['title']) if viet
                   else u"Opening {}...".format(top['title']))
            return {"intent": top['intent'], "params": {}, "message": msg,
                    "_nlu": True, "_authoritative": True}
        lines = []
        for s, _is_id, _cov, t in matches[:3]:
            d = (t.get('desc') or u'').strip()
            lines.append(u"• **{}**{}".format(t['title'],
                                              u" — " + d if d else u""))
        if viet:
            msg = (u"✅ Có! Tool phù hợp:\n{}\n\n"
                   u"Gõ 'mở <tên tool>' để mở nhé.").format(u"\n".join(lines))
        else:
            msg = (u"✅ Yes! Matching tools:\n{}\n\n"
                   u"Type 'open <tool name>' to launch.").format(u"\n".join(lines))
    else:
        near_txt = u", ".join(t['title'] for s, t in near[:3])
        if viet:
            msg = u"❌ Hiện T3Lab chưa có tool riêng cho chức năng đó."
            if near_txt:
                msg += u"\nGần nhất có thể là: {}.".format(near_txt)
            msg += u"\nGõ 'bạn làm được gì' để xem toàn bộ danh sách tool."
        else:
            msg = u"❌ T3Lab doesn't have a dedicated tool for that yet."
            if near_txt:
                msg += u"\nClosest options: {}.".format(near_txt)
            msg += u"\nType 'what can you do' to see the full tool list."
    return {"intent": "help", "params": {"answer": msg}, "message": msg,
            "_nlu": True, "_authoritative": True}


# The assistant's own "Opening X..." line: the strongest referent for "nó"/"it".
# Matched on _norm()-only text (_expand would mangle the title).
_OPENED_LINE_RE = re.compile(r'(?:opening|dang mo)\s+([a-z0-9][a-z0-9 ]*)')


def _tool_by_title(text):
    """Catalog entry whose TITLE is exactly `text` (normalised), else None.
    Resolves "Opening <title>..." even when two tools share an alias."""
    key = u" ".join(_norm(text).split())
    if not key:
        return None
    for tool in _tool_catalog():
        if u" ".join(_norm(tool['title']).split()) == key:
            return tool
    return None


def _last_tool_from_history(history):
    """Intent of the tool the conversation last referred to, newest first: the
    assistant's "Opening X..." line, or a message that IS a tool request (exact
    matcher). The whole message must be about the tool (a substring scan bound
    "thanks for the feedback" to Send Feedback).
    """
    if not history:
        return None
    for entry in reversed(history[-6:]):
        raw = entry.get("content", "") or ""
        content = _norm(raw)

        # Only the assistant's own "Opening X..." lines count, not a user's
        if entry.get("role") == "assistant":
            m = _OPENED_LINE_RE.search(content)
            if m:
                named = m.group(1).strip()
                tool = _tool_by_title(named)
                if tool:
                    return tool['intent']
                tool, _ = resolve_tool(named, exact_only=True)
                if tool:
                    return tool['intent']

        tool, _ = resolve_tool(raw, exact_only=True)
        if tool:
            return tool['intent']
    return None


def _is_pronoun_query(normed_expanded):
    """True when the input refers back to the last tool ("mở nó", "what is it").

    A pronoun word counts only in place of a noun, not as a determiner ("this
    project" would hijack the request); and only in short inputs (≤ 6 tokens).
    """
    tokens = re.sub(r'[^a-z0-9\s]', ' ', normed_expanded).split()
    if not tokens or len(tokens) > 6:
        return False

    padded = u" " + u" ".join(tokens) + u" "
    if any((u" " + p + u" ") in padded for p in _PRONOUN_PHRASES):
        return True

    for i, w in enumerate(tokens):
        if w not in _PRONOUN_WORDS:
            continue
        nxt = tokens[i + 1] if i + 1 < len(tokens) else None
        # Followed by a content word → determiner
        if nxt and len(nxt) >= 2 and nxt not in _STOPWORDS \
                and nxt not in _CAP_BOILERPLATE:
            continue
        return True
    return False


# ─── Scoring ──────────────────────────────────────────────────────────────────

def _score(intent, unigrams, bigrams, padded_expanded):
    """Weighted match score for one intent.

    Single words match the stopword-filtered unigrams; multi-word features match
    as literal phrases in the padded text, since many ("lam duoc gi") are all
    stopwords and could never form a bigram.
    """
    score = 0
    all_features = unigrams | bigrams
    for feature, weight in _TRIGGERS.get(intent, []):
        if " " in feature:
            if (" " + feature + " ") in padded_expanded:
                score += weight
        elif feature in all_features:
            score += weight
    for feature, penalty in _PENALTIES.get(intent, []):
        if " " in feature:
            matched = (" " + feature + " ") in padded_expanded
        else:
            matched = feature in all_features
        if matched:
            score += penalty   # already negative
    return score


# ─── Disambiguation rules ─────────────────────────────────────────────────────

def _disambiguate(scores, unigrams, bigrams, slots):
    """Apply domain rules to ambiguous scores; return the intent or None."""
    all_features = unigrams | bigrams

    # ── open_batchout_configured needs batchout + at least one config param ──
    has_batchout = "batchout" in unigrams
    has_config   = bool(slots["filter"]) or (slots["format"] != "pdf")

    if has_batchout and "open" in all_features and has_config:
        scores["open_batchout_configured"] = max(
            scores.get("open_batchout_configured", 0),
            scores.get("open_batchout", 0) + 10
        )
        scores["open_batchout"] -= 10

    # ── export + batchout without open → export_direct ──
    has_open   = "open" in all_features or "mo" in all_features
    has_export = "export" in all_features

    if has_export and not has_open and has_batchout:
        scores["export_direct"] = max(
            scores.get("export_direct", 0),
            scores.get("open_batchout_configured", 0) + 5
        )

    # ── Filter letter + export word → export_direct ("in G sheet" scores 14) ──
    _export_words = {"export", "in", "print", "pdf", "dwg", "dwf", "dgn",
                     "ifc", "nwd", "img", "sheet"}
    if slots.get("filter") and (_export_words & (unigrams | bigrams)):
        scores["export_direct"] = max(
            scores.get("export_direct", 0),
            _THRESHOLDS["export_direct"]
        )

    if not scores:
        return None
    best_intent = max(scores, key=lambda k: scores[k])
    best_score  = scores[best_intent]

    threshold = _THRESHOLDS.get(best_intent, 18)
    if best_score < threshold:
        return None
    return best_intent


# ─── Message builder ──────────────────────────────────────────────────────────

# Canned messages for intents without a registry entry; auto-discovered tools
# get "Opening <title>..." in _build_message().
_MESSAGES_VI = {
    "open_batchout":          u"Đang mở BatchOut...",
    "open_batchout_configured": u"Mở BatchOut đã cấu hình...",
    "open_loadfamily":        u"Đang mở Family Loader...",
    "check_spelling":         u"Đang quét Text Note trong model để kiểm tra chính tả tiếng Anh...",
    "greet":  u"Xin chào! Tôi là T3Lab Assistant.\nBạn muốn làm gì hôm nay?",
    "farewell": u"Tạm biệt! Gặp lại bạn sau nhé.",
    "chat":   u"Không có gì! Cần gì cứ hỏi tôi nhé.",
    "help":   (u"Tôi có thể giúp bạn:\n"
               u"• Xuất sheet: 'xuất pdf G sheet', 'in tất cả sang dwg'\n"
               u"• Mở tool: 'mở batchout', 'load family', 'mở manaviews'\n"
               u"• Cấu hình nhanh: 'mở batchout G sheet pdf'\n"
               u"Gõ tên tool hoặc mô tả điều bạn muốn làm!"),
}

_MESSAGES_EN = {
    "open_batchout":          "Opening BatchOut...",
    "open_batchout_configured": "Opening BatchOut (pre-configured)...",
    "open_loadfamily":        "Opening Family Loader...",
    "check_spelling":         "Scanning model Text Notes for English spelling errors...",
    "greet":   "Hello! I'm T3Lab Assistant.\nWhat would you like to do today?",
    "farewell": "Goodbye! See you later.",
    "chat":    "You're welcome! Let me know if you need anything.",
    "help":    ("I can help you:\n"
                "• Export sheets: 'export pdf G sheet', 'print all to dwg'\n"
                "• Open tools: 'open batchout', 'load family', 'open manaviews'\n"
                "• Quick config: 'open batchout G sheet pdf'\n"
                "Type a tool name or describe what you want to do!"),
}


def _is_viet(raw):
    viet_chars = (u"àáâãèéêìíòóôõùúýăđơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợ"
                  u"ụủứừửữựỳỵỷỹ")
    return any(c in viet_chars for c in raw.lower())


def _kw_hit(raw_input, normed_raw, keywords):
    """Whole-word keyword test for the canned chat replies.

    Accent-exact when the user typed diacritics ("ngủ" = bedroom, "ngu" =
    insult); folded matching only for unaccented input ("may ngu the").
    Substring matching once made "phòng ngủ" an insult and "quantity" a thanks.
    """
    low     = (raw_input or u"").lower()
    accents = _is_viet(low)
    for kw in keywords:
        kw = (kw or u"").strip().lower()
        if not kw:
            continue
        pat = r'(?<!\w)' + re.escape(kw) + r'(?!\w)'
        try:
            if re.search(pat, low, re.UNICODE):
                return True
            if not accents and kw == _norm(kw) \
                    and re.search(pat, normed_raw, re.UNICODE):
                return True
        except Exception:
            if kw in normed_raw:
                return True
    return False


# Longer inputs are requests that merely contain a reaction word.
_REACTION_MAX_WORDS = 6


# ─── Small-talk keyword tables (shared by classify() and _build_message) ─────
_FAREWELL_KWS = ["tam biet", "tạm biệt", "bye", "bai", "bài",
                 "see you", "hen gap", "hẹn gặp", "goodbye"]
_INSULT_KWS   = ["stupid", "dumb", "useless", "garbage", "trash", "suck",
                 "ngu", "vo dung", "vô dụng", "te qua", "tệ quá",
                 "qua te", "quá tệ"]
_ERROR_KWS    = ["loi", "lỗi", "bi hong", "bị hỏng", "khong chay",
                 "không chạy", "khong hoat dong", "không hoạt động",
                 "error", "broken", "not working"]
_POSITIVE_KWS = ["tuyet", "tuyệt", "tot", "tốt", "ngon", "perfect",
                 "great", "awesome", "nice"]
_THANKS_KWS   = ["cam on", "cảm ơn", "thank", "thanks", "tks", "thks"]
# "How are you?" is instant; low mood goes to the LLM (a canned reply is tone-deaf).
_STATE_HOWAREYOU_KWS = ["khoe", "khỏe"]
_STATE_LOWMOOD_KWS   = ["met", "mệt", "buon", "buồn", "chan", "chán",
                        "sao vay", "sao vậy", "stress"]


# ─── English small-talk the trigger tables miss ───────────────────────────────
# Instant-safe; emotional / complaint phrasing stays with the LLM.
_EN_SMALLTALK_PHRASES = (
    # how are you
    "how are you", "how are u", "how r u", "how are ya", "how are things",
    "how is it going", "hows it going", "how it going", "hows things",
    "how do you do", "how you doing", "how are you doing", "how ya doing",
    "hows your day", "how is your day", "how have you been", "how you been",
    # what's up (apostrophe stripped → "what s up")
    "whats up", "what s up", "what is up", "wassup", "wazzup", "sup",
    "whats good", "whats new", "long time no see",
    # presence / wellbeing checks
    "you there", "are you there", "u there", "you awake", "still there",
    "you good", "you ok", "you okay", "u good", "u ok", "are you ok",
    "are you okay", "you alright", "you all right",
)
# Trailing fillers / names peeled before the exact match ("how are you today").
_EN_SMALLTALK_TAIL = frozenset((
    "today", "now", "then", "mate", "man", "bro", "buddy", "friend", "pal",
    "sir", "dear", "sun", "ban", "a", "ha", "huh", "hey", "hi", "there",
))


def _fold_plain(text):
    """Diacritic-folded, punctuation-free, whitespace-collapsed lowercase."""
    plain = re.sub(r'[^a-z0-9\s]', ' ', _norm(text or ""))
    return " ".join(plain.split())


def _match_english_smalltalk(raw):
    """'howareyou' when the WHOLE message (minus ≤ 2 filler words) is English
    small talk, else None — a command opening with a greeting is not swallowed."""
    f = _fold_plain(raw)
    if not f:
        return None
    words = f.split()
    if len(words) > 7:
        return None
    # Unpeeled first, so "you there" survives the peeling below
    if f in _EN_SMALLTALK_PHRASES:
        return "howareyou"
    for _ in range(2):
        if len(words) >= 2 and words[-1] in _EN_SMALLTALK_TAIL:
            words = words[:-1]
        else:
            break
    if " ".join(words) in _EN_SMALLTALK_PHRASES:
        return "howareyou"
    return None


# Canned even with a provider online; other `chat` kinds go to the LLM (Hybrid).
_INSTANT_CHAT_KINDS = frozenset(("farewell", "thanks", "positive", "howareyou"))


def _chat_smalltalk_kind(raw_input, normed_raw):
    """Kind of a `chat` message, or None. Same tables as _build_message; strong
    kinds first so "how are you" cannot mask a complaint."""
    _short = len((raw_input or u"").split()) <= _REACTION_MAX_WORDS
    if _short and _kw_hit(raw_input, normed_raw, _FAREWELL_KWS):
        return "farewell"
    if _short and _kw_hit(raw_input, normed_raw, _INSULT_KWS):
        return "insult"
    if _kw_hit(raw_input, normed_raw, _ERROR_KWS):
        return "error"
    if _short and _kw_hit(raw_input, normed_raw, _POSITIVE_KWS):
        return "positive"
    if _kw_hit(raw_input, normed_raw, _THANKS_KWS):
        return "thanks"
    if _short and _kw_hit(raw_input, normed_raw, _STATE_HOWAREYOU_KWS):
        return "howareyou"
    if _short and _kw_hit(raw_input, normed_raw, _STATE_LOWMOOD_KWS):
        return "lowmood"
    return None


_HOWAREYOU_REPLIES_EN = (
    "Doing great, thanks for asking! What can I help you with?",
    "All good here! What would you like to work on?",
    "I'm doing well, thank you! How can I help?",
)
_HOWAREYOU_REPLIES_VI = (
    u"Cảm ơn bạn hỏi thăm! Tôi ổn. Bạn cần tôi giúp gì không?",
    u"Tôi vẫn khỏe đây! Bạn muốn làm gì hôm nay?",
    u"Mọi thứ ổn cả! Bạn cần tôi hỗ trợ gì nào?",
)


def _howareyou_reply(viet):
    return random.choice(_HOWAREYOU_REPLIES_VI if viet
                         else _HOWAREYOU_REPLIES_EN)


def _build_message(intent, slots, viet, raw_input=""):
    """Build a friendly message for the given intent and extracted slots."""
    normed_raw = _norm(raw_input)

    if intent == "export_direct":
        fmt  = slots.get("format", "pdf").upper()
        filt = slots.get("filter", "")
        if viet:
            part = u" {} sheet".format(filt) if filt else u" tất cả sheet"
            return u"Đang xuất{} sang {}...".format(part, fmt)
        else:
            part = " {} sheet".format(filt) if filt else " all sheets"
            return "Exporting{} to {}...".format(part, fmt)

    if intent == "open_batchout_configured":
        fmt  = slots.get("format", "pdf").upper()
        filt = slots.get("filter", "")
        if viet:
            part = u" {} sheet".format(filt) if filt else u""
            return u"Mở BatchOut{} ({})...".format(part, fmt)
        else:
            part = " {} sheet".format(filt) if filt else ""
            return "Opening BatchOut{} ({})...".format(part, fmt)

    # ── Chat replies: wording only, _chat_smalltalk_kind() categorises ──
    if intent == "chat":
        if _match_english_smalltalk(raw_input):
            return _howareyou_reply(viet)
        kind = _chat_smalltalk_kind(raw_input, normed_raw)
        if kind == "farewell":
            return (_MESSAGES_VI if viet else _MESSAGES_EN).get("farewell",
                    u"Tạm biệt!" if viet else "Goodbye!")
        if kind == "insult":
            if viet:
                return (u"Xin lỗi vì trải nghiệm chưa tốt! Ở chế độ offline khả năng "
                        u"của tôi hạn chế — kết nối AI trong phần Cài đặt để trả "
                        u"lời tự nhiên hơn.")
            return ("Sorry that reply wasn't good enough! Offline mode is limited — "
                    "connect an AI provider in Settings for smarter answers.")
        if kind == "error":
            if viet:
                return (u"Xin lỗi bạn gặp vấn đề! Bạn có thể thử:\n"
                        u"• Đóng và mở lại tool\n"
                        u"• Kiểm tra Revit console để xem lỗi chi tiết")
            return ("Sorry you're having issues! You can try:\n"
                    "• Close and reopen the tool\n"
                    "• Check the Revit console for error details")
        if kind == "positive":
            return u"Cảm ơn bạn! Cần gì cứ hỏi nhé." if viet else "Thank you! Let me know if you need anything."
        if kind == "thanks":
            return u"Không có gì! Cần gì cứ hỏi tôi nhé." if viet else "You're welcome! Let me know if you need anything."
        if kind == "howareyou":
            return _howareyou_reply(viet)
        # Low mood: light empathy offline; online it goes to the LLM (not _instant)
        if kind == "lowmood":
            if viet:
                return (u"Nghe bạn nói vậy tôi cũng lo. Nghỉ một chút nhé — "
                        u"khi nào cần tôi vẫn ở đây giúp bạn.")
            return ("Sorry to hear that. Take a short break if you can — "
                    "I'm here whenever you need a hand.")

    table = _MESSAGES_VI if viet else _MESSAGES_EN
    if intent in table:
        return table[intent]

    # Auto-discovered tools: message from the live registry title
    if intent and intent.startswith("open_"):
        label = _tool_label(intent)
        return (u"Đang mở {}...".format(label) if viet
                else u"Opening {}...".format(label))

    return u"Đang xử lý..." if viet else "Processing..."


# ─── Public API ───────────────────────────────────────────────────────────────

def classify(user_input, history=None):
    """Classify user_input → {intent, params, message, _nlu: True}, or None if
    below threshold. `history` is a list of {role, content} dicts."""
    if not user_input or not user_input.strip():
        return None

    viet = _is_viet(user_input)

    normed   = _norm(user_input)
    expanded = _expand(normed)

    # ── Capability questions: BEFORE anaphora ("cái này làm được gì" has its
    # own subject; resolving the pronoun would open a tool instead) ──
    if is_capability_question(expanded):
        return answer_capability_question(user_input, viet)

    # ── English small-talk the trigger tables miss: _instant, no LLM ─────────
    if _match_english_smalltalk(user_input):
        return {"intent": "chat", "params": {},
                "message": _howareyou_reply(viet),
                "_nlu": True, "_instant": True}

    # ── Pronoun resolution ───────────────────────────────────────────────────
    if _is_pronoun_query(expanded) and history:
        last_tool = _last_tool_from_history(history)
        if last_tool:
            slots = _extract_slots(user_input)
            label = _tool_label(last_tool)
            msg = (u"Đang mở {}...".format(label) if viet
                   else "Opening {}...".format(label))
            return {"intent": last_tool, "params": {}, "message": msg,
                    "_nlu": True}

    # ── Exact tool-name hit opens exactly that tool. Extra words ("mở batchout
    # G sheet pdf", "mcp control là gì") fall through to scoring ──
    _tool, _ = resolve_tool(user_input, exact_only=True)
    if _tool:
        msg = (u"Đang mở {}...".format(_tool['title']) if viet
               else u"Opening {}...".format(_tool['title']))
        return {"intent": _tool['intent'], "params": {}, "message": msg,
                "_nlu": True}

    unigrams, bigrams = _tokenise(expanded)
    # Punctuation-stripped like _tokenise(), else "what can you do?" never matches
    _clean = re.sub(r'[^a-z0-9\s]', ' ', expanded)
    padded_expanded = u" " + u" ".join(_clean.split()) + u" "

    scores = {intent: _score(intent, unigrams, bigrams, padded_expanded)
              for intent in _TRIGGERS}

    slots = _extract_slots(user_input)
    best = _disambiguate(dict(scores), unigrams, bigrams, slots)

    # ── No trigger matched: ranked resolver; ambiguous "open X" asks to pick ──
    if best is None:
        _tool, _cands = resolve_tool(user_input)
        if _tool:
            msg = (u"Đang mở {}...".format(_tool['title']) if viet
                   else u"Opening {}...".format(_tool['title']))
            return {"intent": _tool['intent'], "params": {}, "message": msg,
                    "_nlu": True}
        if _cands and _has_open_verb(expanded):
            names = u"\n".join(u"• {}".format(c['title']) for c in _cands)
            if viet:
                msg = (u"Bạn muốn mở tool nào? Tôi tìm thấy các tool gần giống:\n"
                       u"{}\nGõ đúng tên tool để mở chính xác nhé!".format(names))
            else:
                msg = ("Which tool do you mean? Closest matches:\n"
                       "{}\nType the exact tool name to open it!".format(names))
            return {"intent": "chat", "params": {}, "message": msg,
                    "_nlu": True, "_authoritative": True}

    # ── Soft "didn't understand" for input without tool words ──
    if best is None:
        _tool_words = {
            "batchout", "parasync", "loadfamily", "projectname",
            "workset", "dimtext", "upperdimtext", "resetoverrides", "grids",
            "export", "open", "in", "print",
        }
        if not (unigrams & _tool_words):
            if viet:
                msg = (u"Xin lỗi, tôi chưa hiểu ý bạn. Bạn có thể thử:\n"
                       u"• 'mở batchout' / 'xuất pdf G sheet'\n"
                       u"• 'parasync', 'load family', 'workset'...")
            else:
                msg = ("Sorry, I didn't understand. You can try:\n"
                       "• 'open batchout' / 'export pdf G sheet'\n"
                       "• 'parasync', 'load family', 'workset'...")
            # _generic_fallback: placeholder, so script.py can show "the AI
            # didn't respond" instead when the LLM call fails
            return {"intent": "chat", "params": {}, "message": msg, "_nlu": True,
                    "_generic_fallback": True}
        return None

    params = {}
    if best in ("export_direct", "open_batchout_configured"):
        params = slots
    elif best == "open_batchout":
        params = {}

    message = _build_message(best, slots, viet, raw_input=user_input)
    result = {"intent": best, "params": params, "message": message,
              "_nlu": True}
    # _instant: the router skips the LLM (biggest latency win on slow local
    # models). `greet` always; `chat` only for _INSTANT_CHAT_KINDS.
    if best == "greet":
        result["_instant"] = True
    elif best == "chat":
        if _chat_smalltalk_kind(user_input, _norm(user_input)) \
                in _INSTANT_CHAT_KINDS:
            result["_instant"] = True
    return result
