"""
tds_matcher.py  -  content-based "best match" finder for TDS PDFs
===================================================================
Instead of requiring the filename to equal the model number, this reads the
TEXT inside every TDS in the repository (first pages + filename + folder
names) and ranks them against each BOQ line item (Description, Make, Model).

It is a free, offline ranking model (IDF-weighted keyword coverage + model
number detection) - no API key, no per-call cost. It reports a confidence
level for every match so doubtful ones can be reviewed instead of silently
used.

Scoring (per line item, per TDS):
  model signal  : model == filename (100) | model contained in filename (70)
                  | model found inside the PDF text (45) | all model words found (25)
  coverage      : share of the line item's keywords (weighted by rarity across the
                  repository, so common words like "cable" count little) that
                  appear in the TDS            -> up to 50 points
  make bonus    : manufacturer name in filename/folder (+8) or in text (+5)
Confidence: High / Medium / Low (see _confidence()).
"""

import difflib
import math
import re
from collections import Counter
from pathlib import Path

import pymupdf

TOKEN_RE = re.compile(r"[a-z0-9]+(?:\.[0-9]+)?")

STOP = {
    "a", "an", "the", "and", "or", "of", "in", "on", "at", "to", "for", "with", "by", "as",
    "is", "are", "be", "shall", "should", "will", "this", "that", "these", "those", "from",
    "per", "including", "include", "includes", "minimum", "maximum", "latest", "version",
    "not", "have", "has", "etc", "any", "all", "each", "such", "if", "than", "then", "into",
    "over", "under", "up", "out", "only", "also", "must", "may", "can", "it", "its",
    "supply", "supplying", "installation", "installing", "testing", "commissioning", "laying",
    "termination", "certified", "fixing", "complete", "work", "works", "tender", "drawings",
    "drawing", "technical", "specification", "specifications", "engineer", "issued", "pdf",
    "datasheet", "data", "sheet", "tds", "type", "floor", "na", "nos", "no",
}
NO_MAKE = {"", "na", "n/a", "-", "none", "nil", "tbd"}

CONF_RANK = {"High": 3, "Medium": 2, "Low": 1}


def stem(t):
    """Very light plural stemming: racks->rack, cables->cable, switches->switch, batteries->battery."""
    if len(t) > 4 and t.endswith("ies"):
        return t[:-3] + "y"
    if len(t) > 4 and t.endswith("es") and t[-3] in "sxz":
        return t[:-2]
    if len(t) > 3 and t.endswith("s") and not t.endswith("ss") and not t[-2].isdigit():
        return t[:-1]
    return t


def normalize_text(text):
    """Unify common electrical/BOQ spellings so look-alike descriptions compare equal."""
    t = (text or "").lower()
    t = re.sub(r"(\d+)\s*(?:c|core|cores)\s*[x\u00d7]\s*(\d+(?:\.\d+)?)", r" \1 core \2 sqmm ", t)   # 3Cx1.0
    t = re.sub(r"sq\.?\s*mm", " sqmm ", t)
    t = re.sub(r"\bd[\s\-]link\b", " dlink ", t)
    return t


def tokenize(text):
    """Unigrams (noise filtered) + number/word pairs such as '4_core', 'cat_6', '1.0_sqmm'
    so items that differ only by a number are told apart."""
    raw = [stem(t) for t in TOKEN_RE.findall(normalize_text(text))]
    out = []
    for t in raw:
        if t in STOP:
            continue
        if t.isdigit() or re.fullmatch(r"\d+\.\d+", t):
            if len(t) >= 3:
                out.append(t)
        elif len(t) >= 2:           # includes '4u', '24v', 'rack'
            out.append(t)
    for a, b in zip(raw, raw[1:]):
        a_num, b_num = a[0].isdigit(), b[0].isdigit()
        if a_num != b_num:
            word = b if a_num else a
            if len(word) >= 2 and word not in STOP:
                out.append(a + "_" + b)
    return out


def compact(text):
    return re.sub(r"[^a-z0-9]", "", (text or "").lower())


def extract_text(path, max_pages=8, max_chars=150_000):
    try:
        doc = pymupdf.open(str(path))
    except Exception:
        return ""
    parts, n = [], 0
    try:
        for page in doc[:max_pages]:
            t = page.get_text()
            parts.append(t)
            n += len(t)
            if n >= max_chars:
                break
    finally:
        doc.close()
    return "".join(parts)[:max_chars]


class Doc:
    __slots__ = ("path", "name", "stem", "folder", "name_set", "text_set",
                 "stem_c", "text_c", "has_text", "tf", "length", "title_set", "excerpt")

    def __init__(self, path, name, folder, text):
        self.path = Path(path)
        self.name = name
        self.stem = name.rsplit(".", 1)[0].strip()
        self.folder = folder
        self.name_set = set(tokenize(self.stem + " " + folder))
        toks = tokenize(text)
        self.text_set = set(toks)
        self.tf = Counter(toks)
        self.length = max(len(toks), 1)
        self.title_set = set(tokenize(text[:300]))
        self.excerpt = re.sub(r"\s+", " ", text[:1200]).strip()[:800]
        self.stem_c = compact(self.stem)
        self.text_c = compact(text)
        self.has_text = len(text.strip()) >= 50


class MatchResult:
    def __init__(self, doc, score, confidence, reason, alternatives):
        self.doc, self.score, self.confidence = doc, score, confidence
        self.reason, self.alternatives = reason, alternatives  # [(name, score), ...]

    @property
    def path(self):
        return self.doc.path if self.doc else None


class TDSLibrary:
    def __init__(self):
        self.docs = []
        self.idf = {}

    # ---- building ---------------------------------------------------------
    def add(self, path, name=None, folder=""):
        path = Path(path)
        self.docs.append(Doc(path, name or path.name, folder, extract_text(path)))

    @classmethod
    def from_dir(cls, root, progress_cb=None):
        lib = cls()
        root = Path(root)
        pdfs = [p for p in root.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf"]
        for i, p in enumerate(pdfs, 1):
            lib.add(p, p.name, str(p.parent.relative_to(root)).replace("\\", " ").replace("/", " "))
            if progress_cb and i % 10 == 0:
                progress_cb(f"Reading TDS files... {i}/{len(pdfs)}")
        lib.finalize()
        return lib

    def finalize(self):
        n = len(self.docs)
        df = Counter()
        for d in self.docs:
            df.update(d.name_set | d.text_set)
        self.idf = {t: math.log(1 + (n - c + 0.5) / (c + 0.5)) for t, c in df.items()}
        self._n = n
        self.avglen = (sum(d.length for d in self.docs) / n) if n else 1.0

    @property
    def count(self):
        return len(self.docs)

    # ---- matching ---------------------------------------------------------
    def _idf(self, tok):
        # unseen tokens are not in any doc, so they can never match; weight irrelevant
        return self.idf.get(tok, math.log(1 + self._n + 0.5))

    def _repair(self, tokens):
        """Fix typos ('switrch' -> 'switch') by snapping unknown words to a close word seen in the repository."""
        if not hasattr(self, "_vocab"):
            self._vocab = [t for t in self.idf if "_" not in t and len(t) >= 5 and not t[0].isdigit()]
        fixed = []
        for t in tokens:
            if t not in self.idf and "_" not in t and len(t) >= 5 and not t[0].isdigit():
                close = difflib.get_close_matches(t, self._vocab, n=1, cutoff=0.82)
                fixed.append(close[0] if close else t)
            else:
                fixed.append(t)
        return fixed

    def prepare(self, row):
        """Tokenize a BOQ line item once (not once per TDS)."""
        model = (row.get("Model no.") or row.get("Model") or "").strip()
        make = (row.get("Make") or "").strip()
        desc = row.get("Description") or ""
        weights, head = {}, set()
        # BOQ descriptions name the product first ("Supply ... of Duct type Temperature Sensor
        # with ...") and give specs afterwards, so the opening words count extra.
        head_words = " ".join(desc.split()[:14])
        head = set(self._repair(tokenize(head_words)))
        for t in self._repair(tokenize(desc)):
            w = 1.5 if "_" in t else 1.0
            if t in head:
                w *= 2.5
            weights[t] = max(weights.get(t, 0), w)
        for t in tokenize(model):
            weights[t] = max(weights.get(t, 0), 3.0)
        make_c = compact(make) if make.lower() not in NO_MAKE else ""
        has_make = make.lower() not in NO_MAKE
        make_tokens = set(tokenize(make)) if has_make else set()
        for t in make_tokens:
            weights[t] = max(weights.get(t, 0), 1.5)
        return {"model": model, "model_c": compact(model), "model_tokens": set(tokenize(model)),
                "weights": weights, "head": head, "make_tokens": make_tokens, "make_c": make_c}

    def score(self, doc, q):
        reasons = []
        model_c = q["model_c"]
        msig = 0
        if len(model_c) >= 3:
            why = ""
            if model_c == doc.stem_c:
                msig, why = 100, "model = filename"
            elif len(model_c) >= 4 and len(doc.stem_c) >= 4 and (model_c in doc.stem_c or doc.stem_c in model_c):
                msig, why = 70, "model ~ filename"
            elif len(model_c) >= 4 and model_c in doc.text_c:
                msig, why = 45, "model found in TDS text"
            elif len(q["model_tokens"]) >= 2 and q["model_tokens"] <= (doc.name_set | doc.text_set):
                msig, why = 25, "model words found in TDS"
            if why:
                reasons.append(why)

        total = hit = 0.0
        matched = 0
        len_norm = 0.25 + 0.75 * (doc.length / self.avglen)
        for t, w in q["weights"].items():
            if t not in self.idf:  # word appears in no TDS at all -> ignore, don't penalise
                continue
            idf = self.idf[t]
            if t in q["head"]:
                idf = max(idf, 1.0)   # the product name must count even if the word is common
            val = w * idf
            total += val
            tf = doc.tf.get(t, 0)
            if t in doc.title_set:
                tf += 3          # words in the datasheet heading are the strongest evidence
            if t in doc.name_set:
                tf += 3          # ...as are words in the filename / folder name
            if tf:
                hit += val * (tf / (tf + 1.2 * len_norm))
                matched += 1
        coverage = hit / total if total else 0.0
        if coverage:
            reasons.append(f"{coverage:.0%} keyword coverage")

        # product-name words (start of the description) that appear in this TDS's title/filename
        head_total = head_hit = 0.0
        title_all = doc.title_set | doc.name_set
        for t in q["head"]:
            if t in self.idf and "_" not in t:
                head_total += 1.0
                if t in title_all:
                    head_hit += 1.0
        title_cov = head_hit / head_total if head_total else 0.0
        if title_cov >= 0.5:
            reasons.append(f"title matches product name ({title_cov:.0%})")

        make_bonus = 0
        mk = q["make_tokens"]
        if mk:
            if mk & doc.name_set:
                make_bonus = 8
            elif mk & doc.text_set:
                make_bonus = 5
            if make_bonus:
                reasons.append("make matches")
        mc = q["make_c"]
        if len(mc) >= 3 and not make_bonus:
            if mc in doc.stem_c or mc in compact(doc.folder):
                make_bonus = 8
            elif mc in doc.text_c:
                make_bonus = 5
            if make_bonus:
                reasons.append("make matches")

        score = msig + 40 * coverage + 25 * title_cov + make_bonus
        return score, msig, (coverage, title_cov), matched, "; ".join(reasons)

    @staticmethod
    def _confidence(msig, cov, tcov, margin, n_matched):
        if msig >= 70:
            return "High"
        if msig >= 45:
            return "High" if cov >= 0.20 else "Medium"
        # Description/make only (no model evidence): need the product name to match the TDS title,
        # real keyword overlap, and a clear lead over the runner-up.
        if tcov >= 0.6 and cov >= 0.35 and margin >= 5 and n_matched >= 3:
            return "High"
        if (tcov >= 0.5 or cov >= 0.5) and margin >= 3 and n_matched >= 2:
            return "Medium"
        return "Low"

    def match(self, row, top_k=3):
        if not self.docs:
            return MatchResult(None, 0, "Low", "repository is empty", [])
        q = self.prepare(row)
        scored = []
        for d in self.docs:
            s, msig, (cov, tcov), nm, why = self.score(d, q)
            scored.append((s, msig, cov, tcov, nm, why, d))
        scored.sort(key=lambda x: x[0], reverse=True)
        best = scored[0]
        margin = best[0] - (scored[1][0] if len(scored) > 1 else 0.0)
        conf = self._confidence(best[1], best[2], best[3], margin, best[4])
        alts = [(x[6].name, round(x[0], 1)) for x in scored[1:top_k]]
        res = MatchResult(best[6], round(best[0], 1), conf, best[5], alts)
        res.margin, res.cov, res.tcov, res.n = round(margin, 1), round(best[2], 2), round(best[3], 2), best[4]
        res.candidates = [(x[6], x[0]) for x in scored[:8]]       # for the optional LLM re-ranker
        return res



# --------------------------------------------------------------------------
# Model / make read-back from the chosen TDS (used when the BOQ has none)
# --------------------------------------------------------------------------

_NOISE = re.compile(r"(?<![\w\-])(?:data\s*sheet|datasheet|tds|technical\s+data|technical\s+specifications?|"
                    r"specifications?|spec|rev\.?\s*\w*\d+|v\d+(?:\.\d+)*|final|copy)(?![\w\-])", re.I)


def infer_model(filename):
    """Best-effort model name from a TDS filename: 'NTD-Datasheet.pdf' -> 'NTD'."""
    stem = filename.rsplit(".", 1)[0]
    stem = re.sub(r"\(\d+\)", " ", stem)            # '(1)' download-copy markers
    stem = stem.replace("_", " ")
    stem = re.sub(r"[\s\-]+(?:data\s*sheet|datasheet|tds|specifications?)\s*$", "", stem.strip(), flags=re.I)
    stem = re.sub(r"^\s*(?:tds|data\s*sheet|datasheet)[\s\-]+", "", stem, flags=re.I)
    stem = _NOISE.sub(" ", stem)
    stem = re.sub(r"\s+", " ", stem).strip(" -_.")
    return stem or filename.rsplit(".", 1)[0]


def infer_make(folder):
    """Manufacturer from a folder name like 'TDS - Nextech' / 'Nextech TDS'. '' if unclear."""
    for part in reversed(re.split(r"\s{2,}|/|\\", folder or "")):
        m = re.search(r"tds\s*[-:]\s*(.+)$", part, re.I) or re.search(r"^(.+?)\s*[-:]?\s*tds$", part, re.I)
        if m and m.group(1).strip():
            return m.group(1).strip()
    m = re.search(r"tds\s*[-:]\s*([A-Za-z0-9 &.]+)$", folder or "", re.I)
    return m.group(1).strip() if m else ""


def option_labels(docs):
    """Unique human-readable label per TDS (adds the folder only when two files share a name)."""
    counts = Counter(d.name for d in docs)
    return {d: (d.name if counts[d.name] == 1 else f"{d.name} [{d.folder}]") for d in docs}


def _plan(self, items, judge=None, min_conf="Medium", progress_cb=None):
    """Decide a TDS for every BOQ line.
    items: [(sno, row_dict)].  judge: optional llm_judge.Judge (reads the shortlisted datasheets).
    Returns a list of dicts (one per line) that the UI can show/edit and run() can build from."""
    from concurrent.futures import ThreadPoolExecutor
    labels = option_labels(self.docs)
    min_rank = CONF_RANK[min_conf]
    entries = []
    for i, (sno, row) in enumerate(items, 1):
        res = self.match(row)
        boq_model = (row.get("Model no.") or row.get("Model") or "").strip()
        cands = [(labels[d], d.path, round(sc, 1)) for d, sc in getattr(res, "candidates", [])]
        entries.append({"sno": sno, "row": row, "res": res, "candidates": cands, "boq_model": boq_model})
        if progress_cb:
            progress_cb(i, len(items), f"Shortlisting S.No {sno}")

    # optional AI judge: only where the ranker has no hard model evidence
    verdicts = {}
    if judge is not None:
        todo = [e for e in entries if e["res"].doc is not None and not (e["boq_model"] and e["res"].confidence == "High")]
        def ask(e):
            cand_docs = [d for d, _ in e["res"].candidates]
            try:
                return e["sno"], judge.decide(e["row"], cand_docs)
            except Exception as ex:                      # never let one bad call stop the batch
                return e["sno"], {"error": str(ex)}
        done = 0
        with ThreadPoolExecutor(max_workers=3) as ex:
            for sno, v in ex.map(ask, todo):
                verdicts[sno] = v
                done += 1
                if progress_cb:
                    progress_cb(done, len(todo), f"AI reading datasheets ({done}/{len(todo)})")

    plan = []
    for e in entries:
        res, row = e["res"], e["row"]
        doc, conf, reason, by = res.doc, res.confidence, res.reason, "ranker"
        model, make = "", ""
        v = verdicts.get(e["sno"])
        if v and "error" not in v:
            by = "AI judge"
            pick = v.get("choice")
            cand_docs = [d for d, _ in res.candidates]
            if pick and 1 <= pick <= len(cand_docs):
                doc = cand_docs[pick - 1]
                conf = {"high": "High", "medium": "Medium"}.get(str(v.get("confidence", "")).lower(), "Low")
                reason = "AI: " + str(v.get("reason", ""))[:160]
                model, make = (v.get("model") or "").strip(), (v.get("make") or "").strip()
            else:
                doc, conf, reason = None, "Low", "AI: no suitable TDS in the repository - " + str(v.get("reason", ""))[:140]
        elif v and "error" in v:
            reason = f"{reason} (AI judge failed: {v['error'][:80]})"
        usable = doc is not None and CONF_RANK[conf] >= min_rank
        if doc is not None and not model:
            model = e["boq_model"] or infer_model(doc.name)
        if doc is not None and not make:
            make = (row.get("Make") or "").strip()
            if make.lower() in NO_MAKE:
                make = infer_make(doc.folder)
        label = option_labels(self.docs)[doc] if doc is not None else ""
        plan.append({
            "sno": e["sno"], "description": (row.get("Description") or "").replace("\n", " ")[:90],
            "candidates": e["candidates"], "match": label, "confidence": conf, "reason": reason, "by": by,
            "use": usable, "model": model if usable or doc is not None else "", "make": make,
            "path": doc.path if doc is not None else None,
        })
    return plan


TDSLibrary.plan = _plan


class AIMatchProvider:
    """Plugs into tds_automation.run(): .match_row(row) -> (Path|None, info dict)."""

    def __init__(self, library, min_confidence="Medium"):
        self.library = library
        self.min_rank = CONF_RANK[min_confidence]

    @property
    def count(self):
        return self.library.count

    def find(self, model_value):  # legacy exact lookup (not used when match_row exists)
        return None

    def match_row(self, row):
        m = self.library.match(row)
        usable = m.doc is not None and CONF_RANK[m.confidence] >= self.min_rank
        info = {
            "Matched TDS": m.doc.name if m.doc else "",
            "Confidence": m.confidence,
            "Score": m.score,
            "Used": "Yes" if usable else "No",
            "Why": m.reason,
            "Other Candidates": " | ".join(f"{n} ({s})" for n, s in m.alternatives),
            "TDS has text layer": "Yes" if (m.doc and m.doc.has_text) else "No",
        }
        return (m.path if usable else None), info
