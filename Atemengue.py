# -*- coding: utf-8 -*-
"""
CIVI-MATH ANALYTICS — Application web Python de collecte, de traitement et
d'analyse statistique des enquêtes liées au BTP et au génie civil.

Lancement :   streamlit run app.py --theme.base dark
Auto-tests :  python app.py --selftest

Technologies : Streamlit · Pandas / NumPy · Plotly · SQLite · ReportLab · openpyxl
Fichier unique : l'image de couverture est intégrée (base64) en fin de fichier.
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import html
import io
import json
import os
import re
import shutil
import sqlite3
import sys
import tempfile
import textwrap
import unicodedata
import uuid
import warnings
from contextlib import contextmanager
from datetime import datetime, timedelta
from xml.sax.saxutils import escape as xml_escape

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ════════════════════════════════════════════════════════════════════════════
# 1. CONSTANTES
# ════════════════════════════════════════════════════════════════════════════
APP_NAME = "CIVI-MATH ANALYTICS"
APP_VERSION = "1.0"
DB_PATH = os.environ.get("CIVIMATH_DB", "civimath.db")

PROFILES = ["Professionnel", "Client", "Observateur"]
PROFILE_LABELS = {
    "Professionnel": "Professionnel du BTP",
    "Client": "Particulier / Client ayant déjà fait construire, rénover ou géré un projet de bâtiment",
    "Observateur": "Observateur / Proche connaissant des ouvriers ou observant régulièrement des chantiers",
}
PROFILE_CODE = "PROFIL"
PROFILE_QUESTION_LABEL = "Niveau de proximité avec le domaine du BTP / Génie Civil"

QTYPES = {
    "unique": "Choix unique",
    "multiple": "Choix multiples",
    "ordinale": "Échelle ordinale",
    "nombre": "Nombre",
    "texte": "Texte libre",
}
QTYPE_BY_LABEL = {v: k for k, v in QTYPES.items()}
CATEGORICAL = ("unique", "multiple", "ordinale")

# Palette accessible (Okabe-Ito) : lisible en cas de déficience de perception des couleurs.
OKABE = ["#56B4E9", "#F0E442", "#CC79A7", "#009E73", "#E69F00", "#0072B2", "#D55E00", "#B8C4D6"]
PROFILE_COLORS = {"Professionnel": "#56B4E9", "Client": "#F0E442", "Observateur": "#CC79A7"}
CYAN = "#4CC9F0"
ALERT = "#FF6B6B"

PAGES = ["Vue d'ensemble", "Collecte", "Import", "Données", "Analyses", "Comparaison", "Rapports", "Paramètres"]
PAGE_ICONS = {
    "Vue d'ensemble": "◈", "Collecte": "✎", "Import": "⇪", "Données": "▦",
    "Analyses": "∑", "Comparaison": "⇄", "Rapports": "▤", "Paramètres": "⚙",
}

A_IGNORE = "— Ignorer cette colonne —"
A_TIME = "⏱ Horodatage (date de soumission)"
A_NEW = "➕ Nouvelle question"

PLOT_CONFIG = {"displaylogo": False, "toImageButtonOptions": {"format": "png", "scale": 2}}

# ════════════════════════════════════════════════════════════════════════════
# 2. OUTILS GÉNÉRAUX
# ════════════════════════════════════════════════════════════════════════════
def esc(x) -> str:
    return html.escape(str(x))


def fr(x, d: int = 1) -> str:
    """Nombre au format français (virgule décimale)."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    return f"{x:.{d}f}".replace(".", ",")


def norm_key(s) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def clean_spaces(s) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()


def normalize_profile(text):
    """Rattache un texte libre ou une option du formulaire à l'un des trois profils."""
    k = norm_key(text)
    if not k:
        return None
    if "professionnel" in k:
        return "Professionnel"
    if "observateur" in k or "proche" in k:
        return "Observateur"
    if "client" in k or "particulier" in k:
        return "Client"
    return None


def wrap_label(s, width: int = 30) -> str:
    s = str(s)
    return "<br>".join(textwrap.wrap(s, width)) or s


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def parse_ts(s):
    s = clean_spaces(s)
    if not s:
        return None
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        t = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if pd.isna(t):
        return None
    return t.strftime("%Y-%m-%d %H:%M:%S")


def sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Neutralise les cellules texte susceptibles d'être interprétées comme formules (Excel)."""
    out = df.copy()
    for c in out.columns:
        if out[c].dtype == object:
            out[c] = out[c].map(
                lambda v: "'" + v if isinstance(v, str) and v[:1] in ("=", "+", "-", "@") and not re.fullmatch(r"-?\d+([.,]\d+)?", v) else v
            )
    return out


# ════════════════════════════════════════════════════════════════════════════
# 3. BASE DE DONNÉES (SQLite)
# ════════════════════════════════════════════════════════════════════════════
SCHEMA = """
CREATE TABLE IF NOT EXISTS questions(
    question_id INTEGER PRIMARY KEY AUTOINCREMENT,
    code        TEXT NOT NULL,
    label       TEXT NOT NULL,
    qtype       TEXT NOT NULL,
    options     TEXT NOT NULL DEFAULT '[]',
    profiles    TEXT NOT NULL DEFAULT '[]',
    required    INTEGER NOT NULL DEFAULT 0,
    position    INTEGER NOT NULL DEFAULT 0,
    active      INTEGER NOT NULL DEFAULT 1,
    is_demo     INTEGER NOT NULL DEFAULT 0,
    UNIQUE(code, is_demo)
);
CREATE TABLE IF NOT EXISTS import_batches(
    import_batch_id TEXT PRIMARY KEY,
    filename    TEXT,
    source      TEXT,
    imported_at TEXT NOT NULL,
    n_rows      INTEGER NOT NULL DEFAULT 0,
    n_ok        INTEGER NOT NULL DEFAULT 0,
    n_rejected  INTEGER NOT NULL DEFAULT 0,
    n_excluded  INTEGER NOT NULL DEFAULT 0,
    is_demo     INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS responses(
    response_id     TEXT PRIMARY KEY,
    source          TEXT NOT NULL,
    submitted_at    TEXT,
    profile_group   TEXT NOT NULL,
    import_batch_id TEXT,
    is_demo         INTEGER NOT NULL DEFAULT 0,
    archived        INTEGER NOT NULL DEFAULT 0,
    fingerprint     TEXT,
    created_at      TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS answers(
    response_id TEXT NOT NULL,
    question_id INTEGER NOT NULL,
    raw_value   TEXT,
    clean_value TEXT,
    PRIMARY KEY(response_id, question_id),
    FOREIGN KEY(response_id) REFERENCES responses(response_id) ON DELETE CASCADE,
    FOREIGN KEY(question_id) REFERENCES questions(question_id)
);
CREATE TABLE IF NOT EXISTS column_map(
    original_column TEXT PRIMARY KEY,
    question_id     INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS audit_log(
    id      INTEGER PRIMARY KEY AUTOINCREMENT,
    ts      TEXT NOT NULL,
    action  TEXT NOT NULL,
    details TEXT
);
CREATE INDEX IF NOT EXISTS idx_resp_scope ON responses(is_demo, archived);
CREATE INDEX IF NOT EXISTS idx_ans_q ON answers(question_id);
"""


@contextmanager
def db():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    with db() as c:
        c.executescript(SCHEMA)
        ensure_profile_question(c, 0)


def ensure_profile_question(c, is_demo: int) -> int:
    row = c.execute("SELECT question_id FROM questions WHERE code=? AND is_demo=?", (PROFILE_CODE, is_demo)).fetchone()
    if row:
        return row[0]
    cur = c.execute(
        "INSERT INTO questions(code,label,qtype,options,profiles,required,position,active,is_demo) "
        "VALUES(?,?,?,?,?,1,0,1,?)",
        (PROFILE_CODE, PROFILE_QUESTION_LABEL, "unique",
         json.dumps(list(PROFILE_LABELS.values()), ensure_ascii=False), "[]", is_demo),
    )
    return cur.lastrowid


def get_profile_qid(scope: int) -> int:
    with db() as c:
        return ensure_profile_question(c, scope)


def audit(action: str, details: str = ""):
    try:
        with db() as c:
            c.execute("INSERT INTO audit_log(ts,action,details) VALUES(?,?,?)", (now_str(), action, details[:2000]))
    except Exception:
        pass


# ── Chargement ───────────────────────────────────────────────────────────────
def load_questions(scope: int, active_only: bool = True) -> pd.DataFrame:
    sql = "SELECT * FROM questions WHERE is_demo=?" + (" AND active=1" if active_only else "") + " ORDER BY position, question_id"
    with db() as c:
        df = pd.read_sql_query(sql, c, params=(scope,))
    df["options_l"] = df["options"].map(lambda s: json.loads(s or "[]"))
    df["profiles_l"] = df["profiles"].map(lambda s: json.loads(s or "[]"))
    return df


def load_responses(scope: int) -> pd.DataFrame:
    with db() as c:
        df = pd.read_sql_query("SELECT * FROM responses WHERE is_demo=? ORDER BY COALESCE(submitted_at, created_at)", c, params=(scope,))
    df["submitted_dt"] = pd.to_datetime(df["submitted_at"], errors="coerce")
    return df


def _to_vals(clean, qtype):
    if clean is None or (isinstance(clean, float) and np.isnan(clean)) or str(clean).strip() == "":
        return []
    if qtype == "multiple":
        try:
            v = json.loads(clean)
            return [str(x) for x in v if str(x).strip()]
        except Exception:
            return [str(clean)]
    return [str(clean)]


def load_answers(scope: int, qall: pd.DataFrame) -> pd.DataFrame:
    with db() as c:
        df = pd.read_sql_query(
            "SELECT a.response_id, a.question_id, a.raw_value, a.clean_value FROM answers a "
            "JOIN responses r ON r.response_id=a.response_id WHERE r.is_demo=?", c, params=(scope,))
    qtype = dict(zip(qall["question_id"], qall["qtype"]))
    df["vals"] = [_to_vals(cv, qtype.get(qid)) for cv, qid in zip(df["clean_value"], df["question_id"])]
    return df


def load_context(scope: int) -> dict:
    qall = load_questions(scope, active_only=False)
    qact = qall[qall["active"] == 1]
    resp_all = load_responses(scope)
    return {
        "scope": scope,
        "qall": qall,
        "qlist": qact.to_dict("records"),
        "qall_list": qall.to_dict("records"),
        "resp_all": resp_all,
        "resp": resp_all[resp_all["archived"] == 0],
        "ans": load_answers(scope, qall),
    }


def load_colmap() -> dict:
    with db() as c:
        return {r["original_column"]: r["question_id"] for r in c.execute("SELECT * FROM column_map")}


def load_existing_fps(scope: int) -> dict:
    out: dict = {}
    with db() as c:
        for r in c.execute("SELECT fingerprint, submitted_at FROM responses WHERE is_demo=? AND fingerprint IS NOT NULL", (scope,)):
            out.setdefault(r["fingerprint"], set()).add(r["submitted_at"])
    return out


# ── Écriture ─────────────────────────────────────────────────────────────────
def make_fp(group: str, items: dict) -> str:
    """Empreinte d'une réponse : profil + réponses nettoyées (clé = libellé normalisé de la question)."""
    payload = json.dumps([group, sorted((norm_key(k), v or "") for k, v in items.items())], ensure_ascii=False)
    return hashlib.sha1(payload.encode("utf-8")).hexdigest()


def insert_response(c, *, group, answers, source, submitted_at, batch_id, is_demo, fp) -> str:
    rid = "R-" + uuid.uuid4().hex[:10].upper()
    c.execute(
        "INSERT INTO responses(response_id,source,submitted_at,profile_group,import_batch_id,is_demo,archived,fingerprint,created_at) "
        "VALUES(?,?,?,?,?,?,0,?,?)", (rid, source, submitted_at, group, batch_id, is_demo, fp, now_str()))
    for qid, (raw, clean) in answers.items():
        if (raw is None or raw == "") and (clean is None or clean == ""):
            continue
        c.execute("INSERT OR REPLACE INTO answers(response_id,question_id,raw_value,clean_value) VALUES(?,?,?,?)", (rid, qid, raw, clean))
    return rid


# ════════════════════════════════════════════════════════════════════════════
# 4. NETTOYAGE ET VALIDATION DES VALEURS
# ════════════════════════════════════════════════════════════════════════════
def match_option(s: str, options: list):
    k = norm_key(s)
    for o in options:
        if norm_key(o) == k:
            return o
    return None


def parse_multi(raw: str, options: list):
    """Découpe une réponse à choix multiples. Les options connues sont reconnues en priorité
    (certaines contiennent des virgules) ; le reste est découpé sur « , » ou « ; »."""
    rest = clean_spaces(raw)
    found = []
    for opt in sorted(options, key=len, reverse=True):
        if not opt:
            continue
        idx = rest.lower().find(opt.lower())
        if idx >= 0:
            found.append(opt)
            rest = rest[:idx] + "|" + rest[idx + len(opt):]
    leftovers = [clean_spaces(x) for x in re.split(r"\||;|,\s", rest) if clean_spaces(x)]
    values = found + [x for x in leftovers if x not in found]
    unexpected = [x for x in leftovers if x not in options]
    return values, unexpected


def clean_cell(raw, qtype: str, options: list):
    """Retourne (valeur_nettoyée | None, avertissement | None). La valeur brute n'est jamais modifiée."""
    s = clean_spaces(raw)
    if s == "":
        return None, None
    if qtype in ("unique", "ordinale"):
        canon = match_option(s, options) if options else None
        if canon:
            return canon, None
        return s, (f"Valeur inattendue « {s[:60]} »" if options else None)
    if qtype == "multiple":
        vals, unexpected = parse_multi(s, options)
        w = f"Option(s) inattendue(s) : {', '.join(unexpected)[:80]}" if (options and unexpected) else None
        return json.dumps(vals, ensure_ascii=False), w
    if qtype == "nombre":
        try:
            x = float(s.replace(",", ".").replace(" ", ""))
            return (str(int(x)) if x == int(x) else str(x)), None
        except ValueError:
            return None, f"Valeur non numérique « {s[:40]} »"
    return s, None


def guess_type(series: pd.Series) -> str:
    vals = series.astype(str).map(clean_spaces)
    vals = vals[vals != ""]
    if vals.empty:
        return "texte"
    num = pd.to_numeric(vals.str.replace(",", ".", regex=False), errors="coerce")
    if num.notna().mean() > 0.95:
        return "nombre"
    nun = vals.nunique()
    tokens = {t for v in vals for t in re.split(r",\s+|;\s*", v) if t.strip()}
    if (vals.str.contains(r",\s|;", regex=True).mean() > 0.15) and len(tokens) <= 20 and len(tokens) < nun:
        return "multiple"
    return "unique" if nun <= 12 else "texte"


def derive_options(series: pd.Series, qtype: str) -> list:
    vals = [clean_spaces(v) for v in series.astype(str)]
    vals = [v for v in vals if v]
    if qtype in ("unique", "ordinale"):
        order = list(dict.fromkeys(vals)) if qtype == "ordinale" else list(pd.Series(vals).value_counts().index)
        if qtype == "ordinale" and all(re.fullmatch(r"\d+([.,]\d+)?", v) for v in order):
            order.sort(key=lambda v: float(v.replace(",", ".")))
        return order[:40]
    if qtype == "multiple":
        toks = [t.strip() for v in vals for t in re.split(r",\s+|;\s*", v) if t.strip()]
        return list(pd.Series(toks).value_counts().index)[:40] if toks else []
    return []


def ordinal_median(values: list, options: list):
    pos = sorted(options.index(v) for v in values if v in options)
    return options[pos[(len(pos) - 1) // 2]] if pos else None


# ════════════════════════════════════════════════════════════════════════════
# 5. LECTURE ET VALIDATION D'UN IMPORT (fonctions pures, testables)
# ════════════════════════════════════════════════════════════════════════════
def read_table(raw: bytes, filename: str, sheet=None) -> pd.DataFrame:
    name = filename.lower()
    if name.endswith((".xlsx", ".xlsm")):
        df = pd.read_excel(io.BytesIO(raw), sheet_name=sheet or 0, dtype=str)
    else:
        text = None
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        df = pd.read_csv(io.StringIO(text), sep=None, engine="python", dtype=str, keep_default_na=False)
    df = df.fillna("").astype(str)
    df.columns = [clean_spaces(c) for c in df.columns]
    empty_cols = [c for c in df.columns if c.lower().startswith("unnamed") and (df[c].str.strip() == "").all()]
    return df.drop(columns=empty_cols)


def default_assignment(col: str, series: pd.Series, qlist: list, colmap: dict, taken: set, profile_qid: int):
    """Propose l'affectation d'une colonne : mémoire des imports précédents, libellé identique, contenu."""
    nk = norm_key(col)
    by_id = {q["question_id"]: q for q in qlist}
    if any(w in nk for w in ("horodateur", "timestamp", "horodatage")):
        return A_TIME
    cand = None
    if col in colmap and colmap[col] in by_id:
        cand = by_id[colmap[col]]
    if cand is None:
        for q in qlist:
            if norm_key(q["label"]) == nk or (q["code"] != PROFILE_CODE and norm_key(q["code"]) == nk):
                cand = q
                break
    if cand is None and profile_qid not in taken:
        vals = series.astype(str).map(clean_spaces)
        vals = vals[vals != ""]
        if "proximit" in nk or nk in ("profil", "profile") or (len(vals) and vals.map(normalize_profile).notna().mean() > 0.9):
            cand = by_id.get(profile_qid)
    if cand is not None and cand["question_id"] not in taken:
        taken.add(cand["question_id"])
        return q_display(cand)
    return A_NEW


def q_display(q: dict) -> str:
    return f"[{q['code']}] {q['label'][:80]}"


def build_import_plan(df: pd.DataFrame, assignments: dict, new_types: dict, new_options: dict,
                      qlist: list, profile_qid: int, existing_fps: dict) -> dict:
    """Valide les lignes ; ne touche pas à la base. Numéro de ligne = numéro dans le fichier (en-tête = ligne 1)."""
    errors, cols, used = [], [], {}
    disp = {q_display(q): q for q in qlist}
    time_cols = [c for c, a in assignments.items() if a == A_TIME]
    if len(time_cols) > 1:
        errors.append("Plusieurs colonnes sont désignées comme horodatage : une seule est autorisée.")
    for c, a in assignments.items():
        if a in (A_IGNORE, A_TIME):
            continue
        if a == A_NEW:
            qt = new_types.get(c, "texte")
            cols.append(dict(col=c, qid=None, label=c, qtype=qt, options=new_options.get(c, []), new=True, required=0, profiles=[]))
        else:
            q = disp.get(a)
            if q is None:
                errors.append(f"Affectation inconnue pour la colonne « {c} ».")
                continue
            if q["question_id"] in used:
                errors.append(f"La question « {q['label'][:60]} » est associée à deux colonnes : « {used[q['question_id']]} » et « {c} ».")
                continue
            used[q["question_id"]] = c
            cols.append(dict(col=c, qid=q["question_id"], label=q["label"], qtype=q["qtype"], options=q["options_l"],
                             new=False, required=q["required"], profiles=q["profiles_l"]))
    prof_col = next((p["col"] for p in cols if p["qid"] == profile_qid), None)
    if prof_col is None:
        errors.append("Aucune colonne n'est associée à la question de profil (niveau de proximité avec le BTP) : import impossible.")
    plan = dict(cols=cols, errors=errors, accepted=[], rejected=[], warnings=[], n_rows=len(df),
                profile_col=prof_col, time_col=time_cols[0] if time_cols else None)
    if errors:
        return plan
    seen: dict = {}
    other = [p for p in cols if p["col"] != prof_col]
    for i, (_, row) in enumerate(df.iterrows()):
        line = i + 2
        if all(clean_spaces(v) == "" for v in row.values):
            plan["rejected"].append({"Ligne": line, "Motif": "Ligne entièrement vide"})
            continue
        grp = normalize_profile(row[prof_col])
        if not grp:
            shown = clean_spaces(row[prof_col])
            plan["rejected"].append({"Ligne": line, "Motif": f"Profil absent ou non reconnu : « {shown[:60]} »" if shown else "Profil absent"})
            continue
        ts = parse_ts(row[plan["time_col"]]) if plan["time_col"] else None
        if plan["time_col"] and clean_spaces(row[plan["time_col"]]) and ts is None:
            plan["warnings"].append({"Ligne": line, "Avertissement": f"Horodatage illisible « {clean_spaces(row[plan['time_col']])[:30]} » : date laissée vide"})
        answers, fp_items, n_answers = {}, {}, 0
        for p in cols:
            raw = row[p["col"]]
            if p["col"] == prof_col:
                answers[p["col"]] = (clean_spaces(raw), PROFILE_LABELS[grp])
                fp_items[p["label"]] = PROFILE_LABELS[grp]
                continue
            clean, w = clean_cell(raw, p["qtype"], p["options"])
            if w:
                plan["warnings"].append({"Ligne": line, "Avertissement": f"{p['label'][:40]} : {w}"})
            if clean is None and clean_spaces(raw) == "":
                if p["required"] and (not p["profiles"] or grp in p["profiles"]):
                    plan["warnings"].append({"Ligne": line, "Avertissement": f"Champ obligatoire vide : {p['label'][:50]}"})
            else:
                n_answers += 1
            answers[p["col"]] = (clean_spaces(raw) if raw is not None else None, clean)
            fp_items[p["label"]] = clean
        if other and n_answers == 0:
            plan["warnings"].append({"Ligne": line, "Avertissement": "Aucune réponse hormis le profil"})
        fp = make_fp(grp, fp_items)
        dup = None
        if fp in existing_fps:
            dup = "Exact (déjà en base)" if (ts and ts in existing_fps[fp]) else "Probable (mêmes réponses déjà en base)"
        elif fp in seen:
            dup = "Exact (même ligne répétée)" if (ts and ts in seen[fp]["ts"]) else f"Probable (mêmes réponses que la ligne {seen[fp]['first']})"
        entry = seen.setdefault(fp, {"first": line, "ts": set()})
        if ts:
            entry["ts"].add(ts)
        plan["accepted"].append(dict(line=line, group=grp, ts=ts, answers=answers, fp=fp, dup=dup))
    return plan


def commit_import(plan: dict, exclude_lines: set, scope: int, filename: str, source: str) -> dict:
    ok = [r for r in plan["accepted"] if r["line"] not in exclude_lines]
    with db() as c:
        profile_qid = ensure_profile_question(c, scope)
        pos = c.execute("SELECT COALESCE(MAX(position),0) FROM questions WHERE is_demo=?", (scope,)).fetchone()[0]
        nums = [int(m.group(1)) for (code,) in c.execute("SELECT code FROM questions WHERE is_demo=?", (scope,))
                if (m := re.fullmatch(r"Q(\d+)", code))]
        n = max(nums + [0])
        for p in plan["cols"]:
            if p["new"]:
                n += 1
                pos += 1
                cur = c.execute(
                    "INSERT INTO questions(code,label,qtype,options,profiles,required,position,active,is_demo) VALUES(?,?,?,?,?,0,?,1,?)",
                    (f"Q{n}", p["label"], p["qtype"], json.dumps(p["options"], ensure_ascii=False), "[]", pos, scope))
                p["qid"] = cur.lastrowid
            c.execute("INSERT OR REPLACE INTO column_map(original_column,question_id) VALUES(?,?)", (p["col"], p["qid"]))
        batch = "B-" + datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:4]
        by_col = {p["col"]: p["qid"] for p in plan["cols"]}
        for r in ok:
            ans = {by_col[col]: v for col, v in r["answers"].items()}
            insert_response(c, group=r["group"], answers=ans, source=source, submitted_at=r["ts"],
                            batch_id=batch, is_demo=scope, fp=r["fp"])
        c.execute(
            "INSERT INTO import_batches(import_batch_id,filename,source,imported_at,n_rows,n_ok,n_rejected,n_excluded,is_demo) VALUES(?,?,?,?,?,?,?,?,?)",
            (batch, filename, source, now_str(), plan["n_rows"], len(ok), len(plan["rejected"]), len(plan["accepted"]) - len(ok), scope))
    audit("import", f"{filename} · lot {batch} · {len(ok)} acceptées · {len(plan['rejected'])} rejetées · {len(plan['accepted']) - len(ok)} exclues · scope={'démo' if scope else 'réel'}")
    return {"batch": batch, "ok": len(ok), "rejected": len(plan["rejected"]), "excluded": len(plan["accepted"]) - len(ok)}


# ════════════════════════════════════════════════════════════════════════════
# 6. STATISTIQUES
# ════════════════════════════════════════════════════════════════════════════
def apply_filters(df: pd.DataFrame, f: dict) -> pd.DataFrame:
    if df.empty:
        return df
    out = df[df["profile_group"].isin(f["profiles"])]
    if f.get("sources"):
        out = out[out["source"].isin(f["sources"])]
    if f.get("date_from") is not None and f.get("date_to") is not None:
        dt = out["submitted_dt"]
        start = pd.Timestamp(f["date_from"])
        end = pd.Timestamp(f["date_to"]) + pd.Timedelta(days=1)
        out = out[((dt >= start) & (dt < end)) | (dt.isna() & bool(f.get("no_date", True)))]
    return out


def describe_filters(f: dict, scope: int) -> str:
    parts = ["Périmètre : " + ("données de DÉMONSTRATION" if scope else "données réelles"),
             "Profils : " + (", ".join(f["profiles"]) if f["profiles"] else "aucun")]
    if f.get("sources"):
        parts.append("Origines : " + ", ".join(f["sources"]))
    if f.get("date_from") is not None:
        parts.append(f"Période : du {f['date_from']:%d/%m/%Y} au {f['date_to']:%d/%m/%Y}" + (" (+ réponses sans date)" if f.get("no_date") else ""))
    return " · ".join(parts)


def get_q_data(q: dict, resp: pd.DataFrame, ans: pd.DataFrame):
    concerned = resp if not q["profiles_l"] else resp[resp["profile_group"].isin(q["profiles_l"])]
    a = ans[(ans["question_id"] == q["question_id"]) & ans["response_id"].isin(concerned["response_id"])]
    a = a[a["vals"].map(len) > 0]
    return concerned, a


def count_values(a: pd.DataFrame) -> pd.Series:
    if a.empty:
        return pd.Series(dtype=int)
    return a[["vals"]].explode("vals")["vals"].value_counts()


def modalities(q: dict, counts: pd.Series) -> list:
    opts = list(q["options_l"])
    extra = [v for v in counts.index if v not in opts]
    if q["qtype"] == "nombre":
        extra = sorted(extra, key=lambda x: float(x) if re.fullmatch(r"-?\d+(\.\d+)?", x) else 0)
    return opts + extra


def freq_table(q: dict, a: pd.DataFrame):
    """n_i, f_i = n_i / N, p_i = 100 × n_i / N ; N = nombre de répondants ayant une réponse valide à la question."""
    N = int(a["response_id"].nunique())
    counts = count_values(a)
    mods = modalities(q, counts)
    t = pd.DataFrame({"Modalité": mods, "n": [int(counts.get(m, 0)) for m in mods]})
    t["f"] = t["n"] / N if N else np.nan
    t["p"] = 100 * t["f"]
    return t, N


def group_tables(q: dict, resp: pd.DataFrame, ans: pd.DataFrame, mods: list):
    counts = pd.DataFrame(0, index=mods, columns=PROFILES, dtype=int)
    Ng, conc = {}, {}
    ans_q = ans[ans["question_id"] == q["question_id"]]
    for g in PROFILES:
        rg = resp[resp["profile_group"] == g]
        if q["profiles_l"] and g not in q["profiles_l"]:
            Ng[g], conc[g] = 0, 0
            continue
        a_g = ans_q[ans_q["response_id"].isin(rg["response_id"])]
        a_g = a_g[a_g["vals"].map(len) > 0]
        Ng[g], conc[g] = int(a_g["response_id"].nunique()), len(rg)
        for v, n in count_values(a_g).items():
            if v in counts.index:
                counts.loc[v, g] = int(n)
    denom = pd.Series(Ng).reindex(PROFILES).replace(0, np.nan)
    pct = counts.div(denom, axis=1) * 100
    return counts, pct, Ng, conc


def compute_question(q: dict, resp: pd.DataFrame, ans: pd.DataFrame) -> dict:
    concerned, a = get_q_data(q, resp, ans)
    n_conc, N = len(concerned), int(a["response_id"].nunique())
    out = {"q": q, "n_conc": n_conc, "N": N, "nonresp": (n_conc - N) / n_conc if n_conc else np.nan, "kind": q["qtype"]}
    if q["qtype"] in CATEGORICAL:
        t, _ = freq_table(q, a)
        counts, pct, Ng, conc_g = group_tables(q, resp, ans, list(t["Modalité"]))
        out.update(table=t, counts=counts, pct=pct, Ng=Ng, conc_g=conc_g)
        if q["qtype"] == "ordinale":
            ex = a[["response_id", "vals"]].explode("vals")
            meds = {"Tous": ordinal_median(list(ex["vals"]), q["options_l"])}
            ans_q = ans[ans["question_id"] == q["question_id"]]
            for g in PROFILES:
                ids = resp[resp["profile_group"] == g]["response_id"]
                ag = ans_q[ans_q["response_id"].isin(ids)]
                meds[g] = ordinal_median(list(ag[["vals"]].explode("vals")["vals"].dropna()), q["options_l"])
            out["medians"] = meds
    elif q["qtype"] == "nombre":
        nums = pd.to_numeric(pd.Series([v for vs in a["vals"] for v in vs]), errors="coerce").dropna()
        out["numbers"] = nums
    else:
        out["texts"] = a.merge(resp[["response_id", "profile_group"]], on="response_id")
    return out


def matrix_table(res: dict) -> pd.DataFrame:
    counts, pct, Ng, N = res["counts"], res["pct"], res["Ng"], res["N"]
    rows = []
    for m in counts.index:
        row = {"Modalité": m}
        for g in PROFILES:
            row[f"{g} (n={Ng[g]})"] = "—" if Ng[g] == 0 else f"{counts.loc[m, g]} ({fr(pct.loc[m, g])} %)"
        tot = int(counts.loc[m].sum())
        row[f"Total (N={N})"] = f"{tot} ({fr(100 * tot / N)} %)" if N else "—"
        rows.append(row)
    return pd.DataFrame(rows)


def dominant_modalities(res: dict) -> dict:
    out = {}
    for g in PROFILES:
        if res["Ng"][g] == 0:
            continue
        col = res["pct"][g]
        mx = col.max()
        out[g] = " / ".join(str(m) for m in col.index[col == mx]) if mx > 0 else "—"
    return out


def convergence_table(results: list) -> pd.DataFrame:
    rows = []
    for res in results:
        groups = [g for g in PROFILES if res["Ng"][g] > 0]
        if len(groups) < 2:
            continue
        sub = res["pct"][groups]
        gap = float((sub.max(axis=1) - sub.min(axis=1)).max())
        dom = dominant_modalities(res)
        rows.append({
            "Question": res["q"]["label"],
            "Type": QTYPES[res["q"]["qtype"]],
            "n par profil": " · ".join(f"{g[:3]}. {res['Ng'][g]}" for g in groups),
            "Modalité dominante": " | ".join(f"{g[:3]}. : {dom[g]}" for g in groups),
            "Même modalité dominante": "Oui" if len(set(dom.values())) == 1 else "Non",
            "Écart max. (points de %)": round(gap, 1),
        })
    df = pd.DataFrame(rows)
    return df.sort_values("Écart max. (points de %)", ascending=False) if not df.empty else df


def stats_long(results: list) -> pd.DataFrame:
    rows = []
    for r in results:
        if r["kind"] not in CATEGORICAL:
            continue
        q = r["q"]
        for _, row in r["table"].iterrows():
            rows.append({"Question": q["label"], "Type": QTYPES[q["qtype"]], "Profil": "Tous", "Modalité": row["Modalité"],
                         "n": row["n"], "N (réponses valides)": r["N"], "Fréquence f": round(row["f"], 4) if r["N"] else None,
                         "Pourcentage p (%)": round(row["p"], 2) if r["N"] else None})
        for g in PROFILES:
            if r["Ng"][g] == 0:
                continue
            for m in r["counts"].index:
                rows.append({"Question": q["label"], "Type": QTYPES[q["qtype"]], "Profil": g, "Modalité": m,
                             "n": int(r["counts"].loc[m, g]), "N (réponses valides)": r["Ng"][g],
                             "Fréquence f": round(r["pct"].loc[m, g] / 100, 4), "Pourcentage p (%)": round(r["pct"].loc[m, g], 2)})
    return pd.DataFrame(rows)


def wide_table(D: dict, resp: pd.DataFrame) -> pd.DataFrame:
    """Vue tabulaire : une ligne par répondant, une colonne par question (valeurs nettoyées)."""
    qs = D["qall_list"]
    labels = pd.Series([q["label"] for q in qs])
    names = {q["question_id"]: (q["label"] if (labels == q["label"]).sum() == 1 else f"{q['label']} [{q['code']}]") for q in qs}
    a = D["ans"][D["ans"]["response_id"].isin(resp["response_id"])].copy()
    a["disp"] = a["vals"].map(lambda v: " | ".join(map(str, v)))
    meta = resp[["response_id", "source", "submitted_at", "profile_group", "import_batch_id", "archived"]].copy()
    meta["archived"] = meta["archived"].map({0: "Non", 1: "Oui"})
    meta.columns = ["ID réponse", "Origine", "Date de soumission", "Profil", "Lot d'import", "Archivée"]
    if a.empty:
        return meta
    piv = a.pivot_table(index="response_id", columns="question_id", values="disp", aggfunc="first")
    order = [q["question_id"] for q in qs if q["question_id"] in piv.columns]
    piv = piv[order].rename(columns=names)
    return meta.merge(piv, left_on="ID réponse", right_index=True, how="left")


# ════════════════════════════════════════════════════════════════════════════
# 7. GRAPHIQUES (Plotly) — thème « laboratoire »
# ════════════════════════════════════════════════════════════════════════════
def seq_colors(n: int, c0="#1F5FBF", c1="#BFE9FF"):
    if n <= 1:
        return [c0] * max(n, 1)
    a = [int(c0[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    return ["#%02x%02x%02x" % tuple(int(a[j] + (b[j] - a[j]) * i / (n - 1)) for j in range(3)) for i in range(n)]


def mod_colors(q: dict, n: int):
    return seq_colors(n) if q["qtype"] == "ordinale" else [OKABE[i % len(OKABE)] for i in range(n)]


def style_fig(fig, title=None, xt=None, yt=None, height=420):
    fig.update_layout(
        template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(11,31,63,0.55)",
        font=dict(family="Inter, system-ui, sans-serif", color="#EAF0F8", size=13),
        title=dict(text=title, font=dict(size=15)) if title else None,
        margin=dict(l=10, r=10, t=56 if title else 20, b=10), height=height, separators=", ",
        legend=dict(orientation="h", y=-0.22, font=dict(size=11)),
        hoverlabel=dict(bgcolor="#0B1F3F", font=dict(family="JetBrains Mono, monospace", size=12)),
    )
    fig.update_xaxes(title=xt, gridcolor="rgba(76,201,240,.14)", zeroline=False, linecolor="rgba(76,201,240,.35)")
    fig.update_yaxes(title=yt, gridcolor="rgba(76,201,240,.14)", zeroline=False, linecolor="rgba(76,201,240,.35)")
    return fig


def chart_bar(t: pd.DataFrame, N: int, title: str, color=CYAN):
    if N == 0 or t.empty:
        return None
    mods = [str(m) for m in t["Modalité"]]
    p = t["p"].astype(float).values
    horiz = len(mods) > 5 or max(len(m) for m in mods) > 16
    text = [f"{n} ({fr(v)} %)" for n, v in zip(t["n"], p)]
    cd = np.column_stack([mods, t["n"].astype(str), [str(N)] * len(mods)])
    rng = [0, max(1.0, float(np.nanmax(p))) * 1.32]
    if horiz:
        fig = go.Figure(go.Bar(x=p, y=[wrap_label(m) for m in mods], orientation="h", text=text, textposition="outside",
                               cliponaxis=False, marker_color=color, customdata=cd,
                               hovertemplate="<b>%{customdata[0]}</b><br>Effectif n = %{customdata[1]}<br>Pourcentage = %{x:.1f} %<br>N = %{customdata[2]}<extra></extra>"))
        style_fig(fig, title, "Pourcentage des réponses valides (%)", "Modalité", max(320, 48 * len(mods) + 130))
        fig.update_yaxes(autorange="reversed")
        fig.update_xaxes(range=rng)
    else:
        fig = go.Figure(go.Bar(x=[wrap_label(m, 16) for m in mods], y=p, text=text, textposition="outside", cliponaxis=False,
                               marker_color=color, customdata=cd,
                               hovertemplate="<b>%{customdata[0]}</b><br>Effectif n = %{customdata[1]}<br>Pourcentage = %{y:.1f} %<br>N = %{customdata[2]}<extra></extra>"))
        style_fig(fig, title, "Modalité", "Pourcentage des réponses valides (%)", 430)
        fig.update_yaxes(range=rng)
    return fig


def chart_pie(t: pd.DataFrame, q: dict, title: str):
    t = t[t["n"] > 0]
    if t.empty:
        return None
    fig = go.Figure(go.Pie(labels=[str(m) for m in t["Modalité"]], values=t["n"], hole=0.28, sort=False,
                           marker=dict(colors=mod_colors(q, len(t)), line=dict(color="#07142B", width=2)),
                           texttemplate="%{percent:.1%}<br>(n=%{value})", textposition="inside",
                           hovertemplate="<b>%{label}</b><br>Effectif n = %{value}<br>Part = %{percent:.1%}<extra></extra>"))
    style_fig(fig, title, height=430)
    fig.update_layout(legend=dict(orientation="h", y=-0.1))
    return fig


def chart_stack(res: dict, title: str):
    groups = [g for g in PROFILES if res["Ng"][g] > 0]
    if not groups:
        return None
    q, pct, counts = res["q"], res["pct"], res["counts"]
    xs = [f"{g}<br>(n = {res['Ng'][g]})" for g in groups]
    cols = mod_colors(q, len(pct.index))
    fig = go.Figure()
    for i, m in enumerate(pct.index):
        fig.add_bar(name=textwrap.shorten(str(m), 42, placeholder="…"), x=xs, y=[pct.loc[m, g] for g in groups],
                    marker_color=cols[i], marker_line=dict(color="#07142B", width=1),
                    customdata=[[str(m), int(counts.loc[m, g]), res["Ng"][g]] for g in groups],
                    hovertemplate="<b>%{customdata[0]}</b><br>%{x}<br>Effectif n = %{customdata[1]} sur %{customdata[2]}<br>Part du groupe = %{y:.1f} %<extra></extra>")
    style_fig(fig, title, "Profil du répondant", "Part des réponses valides du groupe (%)", 470)
    fig.update_layout(barmode="stack")
    fig.update_yaxes(range=[0, 100])
    return fig


def chart_group(res: dict, title: str):
    groups = [g for g in PROFILES if res["Ng"][g] > 0]
    if not groups:
        return None
    pct, counts = res["pct"], res["counts"]
    mods = [str(m) for m in pct.index]
    horiz = len(mods) > 5 or max(len(m) for m in mods) > 16
    fig = go.Figure()
    for g in groups:
        cd = [[m, int(counts.loc[m, g]), res["Ng"][g]] for m in pct.index]
        vals = [pct.loc[m, g] for m in pct.index]
        labs = [wrap_label(m, 28 if horiz else 16) for m in mods]
        common = dict(name=f"{g} (n={res['Ng'][g]})", marker_color=PROFILE_COLORS[g], customdata=cd)
        if horiz:
            fig.add_bar(x=vals, y=labs, orientation="h", hovertemplate="<b>%{customdata[0]}</b><br>" + g + "<br>n = %{customdata[1]} sur %{customdata[2]}<br>Part du groupe = %{x:.1f} %<extra></extra>", **common)
        else:
            fig.add_bar(x=labs, y=vals, hovertemplate="<b>%{customdata[0]}</b><br>" + g + "<br>n = %{customdata[1]} sur %{customdata[2]}<br>Part du groupe = %{y:.1f} %<extra></extra>", **common)
    if horiz:
        style_fig(fig, title, "Part des réponses valides du groupe (%)", "Modalité", max(380, 38 * len(mods) * len(groups) // 2 + 150))
        fig.update_yaxes(autorange="reversed")
    else:
        style_fig(fig, title, "Modalité", "Part des réponses valides du groupe (%)", 450)
    fig.update_layout(barmode="group")
    return fig


def chart_heat(res: dict, title: str):
    groups = [g for g in PROFILES if res["Ng"][g] > 0]
    if not groups:
        return None
    z = res["pct"][groups]
    fig = go.Figure(go.Heatmap(
        z=z.values, x=[f"{g} (n={res['Ng'][g]})" for g in groups], y=[wrap_label(m, 34) for m in z.index],
        colorscale=[[0, "#0B1F3F"], [0.5, "#2F7BFF"], [1, "#BFE9FF"]], zmin=0, zmax=max(100.0, float(np.nanmax(z.values))) if z.size else 100,
        texttemplate="%{z:.0f} %", colorbar=dict(title="%"),
        hovertemplate="%{y}<br>%{x}<br>Part du groupe = %{z:.1f} %<extra></extra>"))
    style_fig(fig, title, "Profil", "Modalité", max(320, 42 * len(z.index) + 140))
    fig.update_yaxes(autorange="reversed")
    return fig


def chart_timeline(resp: pd.DataFrame):
    d = resp.dropna(subset=["submitted_dt"])
    if d.empty:
        return None
    s = d.groupby(d["submitted_dt"].dt.date).size()
    full = pd.date_range(s.index.min(), s.index.max(), freq="D")
    s = s.reindex(full.date, fill_value=0)
    fig = go.Figure()
    fig.add_bar(x=list(s.index), y=s.values, name="Réponses par jour", marker_color=CYAN,
                hovertemplate="%{x|%d/%m/%Y}<br>%{y} réponse(s)<extra></extra>")
    fig.add_scatter(x=list(s.index), y=s.cumsum().values, name="Cumul", mode="lines+markers", yaxis="y2",
                    line=dict(color="#F0E442", width=2), hovertemplate="%{x|%d/%m/%Y}<br>Cumul = %{y}<extra></extra>")
    style_fig(fig, "Volume de réponses dans le temps", "Date de soumission", "Réponses par jour", 380)
    fig.update_layout(yaxis2=dict(title="Cumul", overlaying="y", side="right", gridcolor="rgba(0,0,0,0)"))
    return fig


# ════════════════════════════════════════════════════════════════════════════
# 8. EXPORTS : CSV, XLSX, PDF
# ════════════════════════════════════════════════════════════════════════════
def csv_bytes(df: pd.DataFrame, header_lines: list) -> bytes:
    head = "".join(f"# {l}\n" for l in header_lines)
    return ("﻿" + head + sanitize_df(df).to_csv(index=False, sep=";")).encode("utf-8")


def xlsx_bytes(sheets: dict, meta: list) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as xw:
        pd.DataFrame(meta, columns=["Paramètre", "Valeur"]).to_excel(xw, sheet_name="Filtres", index=False)
        for name, df in sheets.items():
            sanitize_df(df).to_excel(xw, sheet_name=name[:31], index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max((len(str(c.value)) if c.value is not None else 0) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 60)
    return bio.getvalue()


def mpl_png(kind: str, res: dict) -> bytes:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    def clean(s):
        return textwrap.shorten(str(s), 58, placeholder="…").replace("$", r"\$")

    if kind == "bar":
        t = res["table"]
        fig, ax = plt.subplots(figsize=(7.2, max(2.0, 0.42 * len(t) + 1.0)))
        labels = [clean(m) for m in t["Modalité"]][::-1]
        vals = list(t["p"].fillna(0))[::-1]
        ax.barh(labels, vals, color="#0072B2")
        mx = max(vals + [1])
        for i, (v, n) in enumerate(zip(vals, list(t["n"])[::-1])):
            ax.text(v + mx * 0.01, i, f"{n} ({fr(v)} %)", va="center", fontsize=8)
        ax.set_xlim(0, mx * 1.25)
        ax.set_xlabel("Pourcentage des réponses valides (%)", fontsize=8)
    else:
        groups = [g for g in PROFILES if res["Ng"][g] > 0]
        mods = list(res["pct"].index)
        fig, ax = plt.subplots(figsize=(7.2, max(2.4, 0.2 * len(mods) * len(groups) + 1.3)))
        h = 0.8 / max(len(groups), 1)
        ypos = np.arange(len(mods))
        for i, g in enumerate(groups):
            vals = [float(res["pct"].loc[m, g]) for m in mods]
            ax.barh(ypos + i * h, vals, h, label=f"{g} (n={res['Ng'][g]})", color=PROFILE_COLORS[g], edgecolor="#333", linewidth=0.3)
        ax.set_yticks(ypos + h * (len(groups) - 1) / 2)
        ax.set_yticklabels([clean(m) for m in mods], fontsize=8)
        ax.invert_yaxis()
        ax.set_xlabel("Part des réponses valides du groupe (%)", fontsize=8)
        ax.legend(fontsize=7, loc="lower right")
    ax.tick_params(labelsize=8)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    fig.tight_layout()
    bio = io.BytesIO()
    fig.savefig(bio, format="png", dpi=170)
    w, h_ = fig.get_size_inches()
    plt.close(fig)
    return bio.getvalue(), w / h_


def build_pdf(title: str, scope: int, filters_txt: str, kpis: list, results: list, with_groups: bool) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import Image, KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    ss = getSampleStyleSheet()
    H1 = ParagraphStyle("H1", parent=ss["Title"], fontSize=20, textColor=colors.HexColor("#0B1F3F"), alignment=0)
    H2 = ParagraphStyle("H2", parent=ss["Heading2"], fontSize=12, textColor=colors.HexColor("#0B3C8C"), spaceBefore=12)
    P = ParagraphStyle("P", parent=ss["BodyText"], fontSize=9, leading=12)
    SM = ParagraphStyle("SM", parent=P, fontSize=8, textColor=colors.HexColor("#555555"))
    CELL = ParagraphStyle("CELL", parent=P, fontSize=8, leading=10)
    CELLH = ParagraphStyle("CELLH", parent=CELL, textColor=colors.white, fontName="Helvetica-Bold")
    gen = datetime.now().strftime("%d/%m/%Y à %H:%M")

    def para(s, st_=P):
        return Paragraph(xml_escape(str(s)), st_)

    def table(df: pd.DataFrame, widths=None):
        data = [[Paragraph(xml_escape(str(c)), CELLH) for c in df.columns]]
        for row in df.itertuples(index=False):
            data.append([Paragraph(xml_escape(str(v)), CELL) for v in row])
        t = Table(data, colWidths=widths, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B3C8C")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#EEF4FB")]),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#B8C4D6")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
        return t

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(colors.HexColor("#666666"))
        canvas.drawString(1.8 * cm, 1.1 * cm, f"{APP_NAME} · rapport généré le {gen} · résultats agrégés" + (" · DONNÉES DE DÉMONSTRATION" if scope else ""))
        canvas.drawRightString(A4[0] - 1.8 * cm, 1.1 * cm, f"Page {doc.page}")
        canvas.restoreState()

    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=A4, leftMargin=1.8 * cm, rightMargin=1.8 * cm, topMargin=1.6 * cm, bottomMargin=1.8 * cm,
                            title=title, author=APP_NAME)
    W = A4[0] - 3.6 * cm
    el = [Paragraph(xml_escape(title), H1), para(f"Généré le {gen}", SM), para(filters_txt, SM), Spacer(1, 6)]
    if scope:
        el.append(Paragraph('<font color="#C0392B"><b>DONNÉES DE DÉMONSTRATION — jeu fictif de test, sans valeur d\'enquête.</b></font>', P))
    el.append(Paragraph("Indicateurs clés", H2))
    el.append(table(pd.DataFrame(kpis, columns=["Indicateur", "Valeur"]), [W * 0.5, W * 0.5]))
    el.append(Paragraph("Notes de lecture", H2))
    for line in [
        "Fréquence f = n / N ; pourcentage p = 100 × n / N, où n est l'effectif de la modalité et N le nombre de réponses valides à la question.",
        "N peut varier d'une question à l'autre (réponses manquantes, questions réservées à certains profils).",
        "Les comparaisons entre profils utilisent les pourcentages au sein de chaque groupe (effectif du groupe indiqué).",
        "Questions à choix multiples : le total des choix peut dépasser le nombre de répondants ; les pourcentages peuvent dépasser 100 % au total.",
        "Étude descriptive : aucune conclusion causale ne doit être tirée de ces résultats.",
    ]:
        el.append(para("• " + line))
    for r in results:
        q = r["q"]
        block = [Paragraph(xml_escape(f"{q['code']} — {q['label']}"), H2),
                 para(f"{QTYPES[q['qtype']]} · N = {r['N']} réponse(s) valide(s) sur {r['n_conc']} répondant(s) concerné(s) · taux de non-réponse : "
                      + (fr(100 * r['nonresp']) + " %" if r["n_conc"] else "—"), SM)]
        if r["kind"] in CATEGORICAL and r["N"] > 0:
            t = r["table"].rename(columns={"n": "Effectif n", "f": "Fréquence f", "p": "Pourcentage p (%)"}).copy()
            t["Fréquence f"] = t["Fréquence f"].map(lambda x: fr(x, 3))
            t["Pourcentage p (%)"] = t["Pourcentage p (%)"].map(lambda x: fr(x))
            block.append(table(t, [W * 0.52, W * 0.14, W * 0.17, W * 0.17]))
            png, ratio = mpl_png("bar", r)
            block += [Spacer(1, 4), Image(io.BytesIO(png), width=min(W, 15.5 * cm), height=min(W, 15.5 * cm) / ratio)]
            if q["qtype"] == "ordinale" and r.get("medians", {}).get("Tous"):
                block.append(para(f"Médiane ordinale (valeur inférieure) : {r['medians']['Tous']}", SM))
            if with_groups and sum(1 for g in PROFILES if r["Ng"][g] > 0) >= 2:
                m = matrix_table(r)
                block += [Spacer(1, 4), para("Répartition selon le profil — effectif (pourcentage au sein du groupe)", SM), table(m)]
                png2, ratio2 = mpl_png("group", r)
                block += [Spacer(1, 4), Image(io.BytesIO(png2), width=min(W, 15.5 * cm), height=min(W, 15.5 * cm) / ratio2)]
        elif r["kind"] == "nombre" and len(r.get("numbers", [])):
            n_ = r["numbers"]
            block.append(table(pd.DataFrame([{"Moyenne": fr(n_.mean(), 2), "Médiane": fr(n_.median(), 2), "Min": fr(n_.min(), 2), "Max": fr(n_.max(), 2)}])))
        elif r["kind"] == "texte":
            block.append(para("Question à réponse libre : les verbatims ne sont pas reproduits dans le rapport agrégé."))
        else:
            block.append(para("Aucune réponse valide pour cette question avec les filtres actuels."))
        el.append(KeepTogether(block[:3]))
        el.extend(block[3:])
    doc.build(el, onFirstPage=footer, onLaterPages=footer)
    return bio.getvalue()


# ════════════════════════════════════════════════════════════════════════════
# 9. STYLE (direction artistique « laboratoire de mathématiques »)
# ════════════════════════════════════════════════════════════════════════════
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');
:root{--bg:#07142B;--panel:#0F2A52;--line:#1D3F73;--tech:#2F7BFF;--cyan:#4CC9F0;--text:#EAF0F8;--muted:#9FB3CC;--alert:#FF6B6B;
--mono:'JetBrains Mono',ui-monospace,'SF Mono',Menlo,Consolas,monospace;}
html,body,.stApp{font-family:'Inter',system-ui,-apple-system,'Segoe UI',Roboto,sans-serif;}
.stApp{background-color:var(--bg);
 background-image:linear-gradient(rgba(76,201,240,.045) 1px,transparent 1px),linear-gradient(90deg,rgba(76,201,240,.045) 1px,transparent 1px),
 linear-gradient(rgba(76,201,240,.10) 1px,transparent 1px),linear-gradient(90deg,rgba(76,201,240,.10) 1px,transparent 1px);
 background-size:24px 24px,24px 24px,120px 120px,120px 120px;background-attachment:fixed;color:var(--text);}
[data-testid="stHeader"]{background:transparent;}
footer{visibility:hidden;}
.block-container{padding-top:2rem;padding-bottom:3rem;max-width:1280px;}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#06112A,#0A1B3A);border-right:1px solid var(--line);}
h1,h2,h3,h4,h5,h6,p,li,label,[data-testid="stWidgetLabel"] p,[data-testid="stMarkdownContainer"],[data-testid="stCaptionContainer"]{color:var(--text);}
[data-testid="stCaptionContainer"]{color:var(--muted);}
.mono{font-family:var(--mono);}
.crumb{font-family:var(--mono);font-size:.72rem;letter-spacing:.14em;color:var(--cyan);text-transform:uppercase;}
.h-title{font-size:1.9rem;font-weight:700;line-height:1.2;margin:.15rem 0 .1rem;color:#fff;}
.h-sub{color:var(--muted);font-size:.95rem;margin-bottom:.9rem;}
.brand{font-family:var(--mono);font-weight:600;letter-spacing:.08em;color:#fff;font-size:1rem;line-height:1.25;}
.brand small{display:block;color:var(--cyan);font-weight:400;font-size:.65rem;letter-spacing:.16em;margin-top:2px;}
.statusbar{font-family:var(--mono);font-size:.82rem;color:var(--muted);border:1px solid var(--line);border-radius:8px;padding:7px 12px;
 background:rgba(11,31,63,.65);margin:2px 0 14px;overflow-x:auto;white-space:nowrap;}
.statusbar b{color:#fff;}
.kpi-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin:6px 0 18px;}
@media(max-width:900px){.kpi-grid{grid-template-columns:repeat(2,minmax(0,1fr));}}
.kpi{position:relative;background:linear-gradient(160deg,rgba(15,42,82,.93),rgba(9,26,54,.93));border:1px solid var(--line);
 border-radius:12px;padding:13px 16px 11px;overflow:hidden;}
.kpi::after{content:"";position:absolute;right:-20px;top:-20px;width:70px;height:70px;border:1px solid rgba(76,201,240,.28);border-radius:50%;}
.kpi::before{content:"";position:absolute;right:-6px;top:-6px;width:30px;height:30px;border:1px solid rgba(76,201,240,.18);border-radius:50%;}
.kpi-n{font-family:var(--mono);font-size:.7rem;color:var(--cyan);letter-spacing:.1em;}
.kpi-v{font-family:var(--mono);font-size:1.95rem;font-weight:600;line-height:1.2;color:#fff;word-break:break-word;}
.kpi-l{color:var(--text);font-size:.85rem;font-weight:500;}
.kpi-s{color:var(--muted);font-size:.74rem;margin-top:2px;}
.formula{border:1px dashed rgba(76,201,240,.55);border-radius:10px;padding:10px 14px;background:rgba(11,31,63,.6);margin:6px 0 12px;}
.formula .eq{font-family:var(--mono);font-size:1.02rem;color:#fff;}
.formula .leg{color:var(--muted);font-size:.78rem;margin-top:4px;}
.formula .leg span{border-bottom:1px dotted var(--cyan);cursor:help;margin-right:10px;}
.banner-demo{border:1px solid var(--alert);color:var(--alert);background:rgba(255,107,107,.09);padding:8px 12px;border-radius:8px;font-weight:600;font-size:.85rem;margin-bottom:10px;}
.banner-real{border:1px solid rgba(76,201,240,.45);color:var(--cyan);background:rgba(76,201,240,.07);padding:6px 12px;border-radius:8px;font-size:.8rem;margin-bottom:10px;font-family:var(--mono);}
.callout{border-left:3px solid var(--cyan);background:rgba(15,42,82,.7);border-radius:6px;padding:10px 14px;margin:8px 0;color:var(--text);font-size:.92rem;}
.callout.alert{border-left-color:var(--alert);background:rgba(255,107,107,.08);}
.chip{display:inline-block;font-family:var(--mono);font-size:.75rem;border:1px solid var(--line);border-radius:999px;padding:2px 10px;margin:2px 6px 2px 0;color:var(--text);background:rgba(11,31,63,.7);}
.cover-wrap img{width:100%;height:auto;display:block;border-radius:14px;border:1px solid var(--line);box-shadow:0 22px 60px rgba(0,0,0,.5);}
.cover-cap{font-family:var(--mono);color:var(--muted);font-size:.78rem;text-align:center;letter-spacing:.06em;margin-top:10px;}
.stButton>button,.stDownloadButton>button,[data-testid="stFormSubmitButton"]>button{background:linear-gradient(90deg,#1E63D6,#2F7BFF);color:#fff;border:0;border-radius:9px;font-weight:600;padding:.5rem 1.1rem;}
.stButton>button:hover,.stDownloadButton>button:hover,[data-testid="stFormSubmitButton"]>button:hover{filter:brightness(1.12);color:#fff;}
.stButton>button:disabled{opacity:.45;}
div[data-baseweb="input"]>div,div[data-baseweb="select"]>div,div[data-baseweb="textarea"],textarea,input{background-color:#0B1F3F!important;color:var(--text)!important;border-color:var(--line)!important;}
div[data-baseweb="popover"] ul,ul[role="listbox"]{background:#0B1F3F!important;color:var(--text)!important;}
button[data-baseweb="tab"]{color:var(--muted);font-weight:500;}
button[data-baseweb="tab"][aria-selected="true"]{color:var(--cyan);}
[data-testid="stExpander"]{border:1px solid var(--line);border-radius:10px;background:rgba(11,31,63,.45);}
[data-testid="stSidebar"] [role="radiogroup"] label{padding:3px 0;}
hr{border-color:var(--line);}
</style>
"""

FORMULA_HTML = """
<div class="formula"><div class="eq"><i>f</i><sub>i</sub> = <i>n</i><sub>i</sub> / <i>N</i> &nbsp;;&nbsp; <i>p</i><sub>i</sub> = 100 × <i>n</i><sub>i</sub> / <i>N</i></div>
<div class="leg"><span title="Fréquence relative de la modalité i">f<sub>i</sub></span><span title="Pourcentage de la modalité i">p<sub>i</sub></span>
<span title="Effectif : nombre de répondants ayant choisi la modalité i">n<sub>i</sub></span>
<span title="Nombre de réponses valides à CETTE question (varie selon les réponses manquantes)">N</span>
— N est affiché à côté de chaque résultat.</div></div>
"""


# ════════════════════════════════════════════════════════════════════════════
# 10. COMPOSANTS D'INTERFACE
# ════════════════════════════════════════════════════════════════════════════
def show_df(df, **kw):
    try:
        st.dataframe(df, width="stretch", hide_index=True, **kw)
    except Exception:
        st.dataframe(df, use_container_width=True, hide_index=True, **kw)


def show_editor(df, **kw):
    try:
        return st.data_editor(df, width="stretch", hide_index=True, **kw)
    except Exception:
        return st.data_editor(df, use_container_width=True, hide_index=True, **kw)


def show_plot(fig, key=None):
    if fig is None:
        callout("Pas assez de données pour tracer ce graphique.")
        return
    try:
        st.plotly_chart(fig, width="stretch", theme=None, key=key, config=PLOT_CONFIG)
    except Exception:
        st.plotly_chart(fig, use_container_width=True, theme=None, key=key, config=PLOT_CONFIG)


def callout(text, alert=False):
    st.markdown(f'<div class="callout{" alert" if alert else ""}">{text}</div>', unsafe_allow_html=True)


def header(title, subtitle=""):
    st.markdown(f'<div class="crumb">CIVI-MATH ANALYTICS / LABORATOIRE STATISTIQUE</div><div class="h-title">{esc(title)}</div>'
                f'<div class="h-sub">{esc(subtitle)}</div>', unsafe_allow_html=True)


def scope_banner(scope):
    if scope:
        st.markdown('<div class="banner-demo">⚠ DONNÉES DE DÉMONSTRATION — jeu fictif de test, jamais mélangé aux réponses réelles.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="banner-real">● DONNÉES RÉELLES</div>', unsafe_allow_html=True)


def kpi_grid(items):
    html_ = '<div class="kpi-grid">' + "".join(
        f'<div class="kpi"><div class="kpi-n">{esc(n)}</div><div class="kpi-v">{esc(v)}</div><div class="kpi-l">{esc(l)}</div><div class="kpi-s">{esc(s)}</div></div>'
        for n, l, v, s in items) + "</div>"
    st.markdown(html_, unsafe_allow_html=True)


def formula_box():
    st.markdown(FORMULA_HTML, unsafe_allow_html=True)


def get_admin_password():
    pw = os.environ.get("CIVIMATH_ADMIN_PASSWORD")
    if pw:
        return pw
    try:
        return st.secrets.get("CIVIMATH_ADMIN_PASSWORD")
    except Exception:
        return None


def is_admin() -> bool:
    return get_admin_password() is None or bool(st.session_state.get("admin_ok"))


def need_admin(action: str) -> bool:
    if is_admin():
        return True
    callout(f"🔒 {esc(action)} : réservé aux administrateurs. Connectez-vous dans <b>Paramètres → Accès administrateur</b>.", alert=True)
    return False


def goto(page: str):
    st.session_state["nav"] = page


def question_widget(q: dict, key: str, default=None):
    """Widget adapté au type de question. Retourne None / str / list."""
    label = q["label"] + (" *" if q["required"] else "")
    opts, qt = q["options_l"], q["qtype"]
    if qt in ("unique", "ordinale") and opts:
        idx = opts.index(default) if default in opts else None
        return st.radio(label, opts, index=idx, key=key, horizontal=(qt == "ordinale" and len(opts) <= 6))
    if qt == "multiple" and opts:
        dflt = [v for v in (default or []) if v in opts]
        return st.multiselect(label, opts, default=dflt, key=key)
    if qt == "multiple":
        return st.text_input(label + " (séparer les choix par « ; »)", value="; ".join(default or []), key=key)
    if qt == "nombre":
        return st.text_input(label + " (nombre)", value=default or "", key=key)
    if qt == "texte":
        return st.text_area(label, value=default or "", key=key)
    return st.text_input(label, value=default or "", key=key)


def encode_answer(q: dict, v):
    """(brut, nettoyé) ou None si vide ; ValueError si invalide."""
    if v is None:
        return None
    if q["qtype"] == "multiple":
        vals = list(v) if isinstance(v, (list, tuple)) else [clean_spaces(x) for x in re.split(r"[;|]", str(v)) if clean_spaces(x)]
        return (", ".join(vals), json.dumps(vals, ensure_ascii=False)) if vals else None
    s = clean_spaces(v)
    if not s:
        return None
    if q["qtype"] == "nombre":
        clean, w = clean_cell(s, "nombre", [])
        if clean is None:
            raise ValueError(f"« {q['label'][:50]} » : valeur numérique attendue")
        return s, clean
    if q["qtype"] in ("unique", "ordinale") and q["options_l"]:
        canon = match_option(s, q["options_l"])
        return s, canon or s
    return s, s


def sidebar_filters(D: dict) -> dict:
    scope, resp_all = D["scope"], D["resp_all"]
    f = {"profiles": list(PROFILES), "sources": [], "date_from": None, "date_to": None, "no_date": True}
    with st.sidebar.expander("Filtres globaux", expanded=False):
        f["profiles"] = st.multiselect("Profils", PROFILES, default=PROFILES, key=f"flt_prof_{scope}")
        sources = sorted(resp_all["source"].unique()) if not resp_all.empty else []
        if sources:
            f["sources"] = st.multiselect("Origine des réponses", sources, default=sources, key=f"flt_src_{scope}_{'|'.join(sources)}")
        dts = resp_all["submitted_dt"].dropna()
        if not dts.empty:
            mn, mx = dts.min().date(), dts.max().date()
            rng = st.date_input("Période", value=(mn, mx), min_value=mn, max_value=mx, key=f"flt_date_{scope}_{mn}_{mx}")
            if isinstance(rng, (tuple, list)) and len(rng) == 2:
                f["date_from"], f["date_to"] = rng
            elif isinstance(rng, (tuple, list)) and len(rng) == 1:
                f["date_from"] = f["date_to"] = rng[0]
            f["no_date"] = st.checkbox("Inclure les réponses sans date", True, key=f"flt_nodate_{scope}")
    return f


def question_picker(D, label, key, types=CATEGORICAL, include_profile=True, default_non_profile=True):
    qs = [q for q in D["qlist"] if q["qtype"] in types and (include_profile or q["code"] != PROFILE_CODE)]
    if not qs:
        return None
    if default_non_profile:
        qs = sorted(qs, key=lambda q: q["code"] == PROFILE_CODE)
    names = {f"[{q['code']}] {q['label']}": q for q in qs}
    return names[st.selectbox(label, list(names), key=key)]


# ════════════════════════════════════════════════════════════════════════════
# 11. PAGES
# ════════════════════════════════════════════════════════════════════════════
def empty_state(scope):
    callout("Aucune réponse dans ce périmètre pour le moment. Utilisez <b>Import</b> (CSV/XLSX exporté de Google Forms), <b>Collecte</b> "
            "(saisie directe)" + (", ou générez un jeu de test dans <b>Paramètres → Données de démonstration</b>." if not scope else "."))


# ── Vue d'ensemble ──────────────────────────────────────────────────────────
def page_overview(D, F, scope):
    header("Vue d'ensemble", "Indicateurs principaux, profils représentés et accès rapide aux modules")
    scope_banner(scope)
    resp_all_f, resp_f = apply_filters(D["resp_all"], F), apply_filters(D["resp"], F)
    ans_f = D["ans"][D["ans"]["response_id"].isin(resp_f["response_id"])]
    total, valid = len(resp_all_f), len(resp_f)
    last = D["resp_all"]["created_at"].max() if not D["resp_all"].empty else None
    last_txt = pd.to_datetime(last).strftime("%d/%m/%Y %H:%M") if last else "—"
    n_prof = int((resp_f["profile_group"].value_counts() > 0).sum()) if valid else 0
    st.markdown(f'<div class="statusbar">N = <b>{valid}</b> réponses valides | <b>{n_prof}</b> profil(s) | Dernière mise à jour : <b>{last_txt}</b></div>', unsafe_allow_html=True)
    if valid == 0:
        empty_state(scope)
    counts = resp_f["profile_group"].value_counts().reindex(PROFILES, fill_value=0)
    major = counts.idxmax() if valid else "—"
    analysed = sum(1 for q in D["qlist"] if not ans_f[(ans_f["question_id"] == q["question_id"]) & (ans_f["vals"].map(len) > 0)].empty)
    kpi_grid([
        ("01", "Effectif total", str(total), f"dont {total - valid} archivée(s)/exclue(s)"),
        ("02", "Réponses valides", str(valid), "non archivées, filtres appliqués"),
        ("03", "Profil majoritaire", major, f"n = {int(counts.max()) if valid else 0} · {fr(100 * counts.max() / valid) if valid else '—'} % des valides"),
        ("04", "Questions analysées", str(analysed), f"sur {len(D['qlist'])} question(s) actives"),
    ])
    st.markdown("**Raccourcis**")
    cols = st.columns(4)
    for col, p in zip(cols, ["Collecte", "Import", "Analyses", "Comparaison"]):
        col.button(f"{PAGE_ICONS[p]}  {p}", key=f"sc_{p}", on_click=goto, args=(p,))
    if valid == 0:
        return
    st.markdown("&nbsp;")
    q = question_picker(D, "Question affichée dans les graphiques", "ov_q")
    if q is None:
        return
    res = compute_question(q, resp_f, ans_f)
    t1, t2, t3, t4, t5 = st.tabs(["Distribution globale", "Comparaison (barres empilées 100 %)", "Matrice des profils", "Évolution", "Tableau des réponses"])
    with t1:
        st.caption(f"N = {res['N']} réponse(s) valide(s) · {res['n_conc']} répondant(s) concerné(s)")
        show_plot(chart_bar(res["table"], res["N"], q["label"]) if res["kind"] in CATEGORICAL else None, key="ov_bar")
    with t2:
        if res["kind"] not in CATEGORICAL:
            callout("Comparaison disponible pour les questions à choix.")
        elif q["qtype"] == "multiple":
            callout("Question à choix multiples : les parts peuvent dépasser 100 % au total ; barres groupées utilisées au lieu des barres empilées.")
            show_plot(chart_group(res, q["label"]), key="ov_grp")
        else:
            show_plot(chart_stack(res, q["label"]), key="ov_stack")
    with t3:
        if res["kind"] in CATEGORICAL:
            show_df(matrix_table(res))
    with t4:
        show_plot(chart_timeline(resp_f), key="ov_time") if resp_f["submitted_dt"].notna().any() else callout("Aucune date de soumission disponible pour tracer la courbe temporelle.")
    with t5:
        wide = wide_table(D, resp_f)
        show_df(wide.head(200))
        st.caption("200 premières lignes — utilisez la page Données pour la recherche, la pagination et l'export.")


# ── Collecte ────────────────────────────────────────────────────────────────
def page_collecte(D, scope):
    header("Collecte directe", "Saisie d'une réponse : sections conditionnelles selon le profil, contrôles de cohérence")
    scope_banner(scope)
    nonce = st.session_state.get("collect_nonce", 0)
    if st.session_state.get("collect_ok"):
        st.success(f"Réponse enregistrée sous l'identifiant {st.session_state.pop('collect_ok')}. Vous pouvez la retrouver dans la page Données.")
    grp_label = st.radio("Niveau de proximité avec le domaine du BTP / Génie Civil *", list(PROFILE_LABELS.values()), index=None, key=f"c_prof_{scope}_{nonce}")
    if not grp_label:
        callout("Choisissez d'abord votre niveau de proximité : les questions suivantes dépendent du profil.")
        return
    grp = normalize_profile(grp_label)
    qs = [q for q in D["qlist"] if q["code"] != PROFILE_CODE and (not q["profiles_l"] or grp in q["profiles_l"])]
    if not qs:
        callout("Aucune question n'est encore définie pour ce profil. Importez un export Google Forms (les questions sont créées automatiquement) ou ajoutez-les dans <b>Paramètres → Questionnaire</b>.")
        return
    with st.form(f"form_{scope}_{nonce}"):
        vals = {q["question_id"]: question_widget(q, f"c_{scope}_{nonce}_{q['question_id']}") for q in qs}
        submitted = st.form_submit_button("Enregistrer la réponse")
    if not submitted:
        return
    errors, answers, fp_items = [], {}, {}
    for q in qs:
        try:
            enc = encode_answer(q, vals[q["question_id"]])
        except ValueError as e:
            errors.append(str(e))
            continue
        if enc is None and q["required"]:
            errors.append(f"Réponse obligatoire manquante : « {q['label'][:70]} »")
        if enc is not None:
            answers[q["question_id"]] = enc
        fp_items[q["label"]] = enc[1] if enc else None
    if not answers and not errors:
        errors.append("Aucune réponse saisie.")
    if errors:
        for e in errors:
            st.error(e)
        return
    pq = get_profile_qid(scope)
    answers[pq] = (grp_label, PROFILE_LABELS[grp])
    fp_items[PROFILE_QUESTION_LABEL] = PROFILE_LABELS[grp]
    with db() as c:
        rid = insert_response(c, group=grp, answers=answers, source="Saisie application", submitted_at=now_str(),
                              batch_id=None, is_demo=scope, fp=make_fp(grp, fp_items))
    audit("saisie", f"réponse {rid} · profil {grp} · scope={'démo' if scope else 'réel'}")
    st.session_state["collect_ok"] = rid
    st.session_state["collect_nonce"] = nonce + 1
    st.rerun()


# ── Import ──────────────────────────────────────────────────────────────────
def page_import(D, scope):
    header("Import de données", "CSV / XLSX exporté depuis Google Forms ou Google Sheets · aperçu, validation, doublons, rapport d'erreurs")
    scope_banner(scope)
    if scope:
        callout("L'import sera enregistré dans le périmètre <b>démonstration</b> (test).")
    nonce = st.session_state.get("imp_nonce", 0)
    if st.session_state.get("imp_done"):
        r = st.session_state.pop("imp_done")
        st.success(f"Import terminé — lot {r['batch']} : {r['ok']} ligne(s) enregistrée(s), {r['rejected']} rejetée(s), {r['excluded']} doublon(s) exclu(s).")
    up = st.file_uploader("Fichier d'export des réponses (CSV, XLSX)", type=["csv", "xlsx", "xlsm"], key=f"up_{scope}_{nonce}")
    if up is None:
        callout("Dans Google Forms : onglet <b>Réponses → Lier à Google Sheets</b>, puis <b>Fichier → Télécharger → CSV ou Excel</b>. "
                "Les colonnes sont associées aux questions ; les questions inconnues peuvent être créées automatiquement.")
        return
    raw = up.getvalue()
    sheet = None
    if up.name.lower().endswith((".xlsx", ".xlsm")):
        try:
            names = pd.ExcelFile(io.BytesIO(raw)).sheet_names
            sheet = st.selectbox("Feuille à importer", names) if len(names) > 1 else names[0]
        except Exception as e:
            st.error(f"Fichier Excel illisible : {e}")
            return
    try:
        df = read_table(raw, up.name, sheet)
    except Exception as e:
        st.error(f"Impossible de lire le fichier ({type(e).__name__}). Vérifiez qu'il s'agit bien d'un export CSV/XLSX valide.")
        return
    if df.empty or len(df.columns) < 2:
        st.error("Le fichier ne contient aucune ligne de données exploitable.")
        return
    sig = f"{up.name}|{up.size}|{sheet}|{scope}|{nonce}"
    if st.session_state.get("imp_sig") != sig:
        st.session_state["imp_sig"] = sig
        st.session_state.pop("imp_plan", None)
    st.markdown(f"**Aperçu** — {len(df)} ligne(s) × {len(df.columns)} colonne(s)")
    show_df(df.head(15))

    qlist, pq = D["qlist"], get_profile_qid(scope)
    colmap, taken, rows = load_colmap(), set(), []
    for c in df.columns:
        a = default_assignment(c, df[c], qlist, colmap, taken, pq)
        qt = guess_type(df[c])
        ex = next((clean_spaces(v) for v in df[c] if clean_spaces(v)), "")
        rows.append({"Colonne source": c, "Exemple": ex[:50], "Affectation": a, "Type (si nouvelle)": QTYPES[qt],
                     "Options (si nouvelle, séparées par « | »)": " | ".join(derive_options(df[c], qt)) if a == A_NEW else ""})
    st.markdown("**Association des colonnes** — vérifiez l'affectation (les noms d'origine sont mémorisés pour les imports suivants)")
    choices = [A_IGNORE, A_TIME, A_NEW] + [q_display(q) for q in qlist]
    edited = show_editor(
        pd.DataFrame(rows), key=f"map_{sig}", disabled=["Colonne source", "Exemple"],
        column_config={
            "Affectation": st.column_config.SelectboxColumn("Affectation", options=choices, required=True, width="large"),
            "Type (si nouvelle)": st.column_config.SelectboxColumn("Type (si nouvelle)", options=list(QTYPES.values()), required=True),
            "Options (si nouvelle, séparées par « | »)": st.column_config.TextColumn(width="large"),
        })
    if st.button("Analyser et valider le fichier", key="imp_check"):
        assign = dict(zip(edited["Colonne source"], edited["Affectation"]))
        types = {c: QTYPE_BY_LABEL[t] for c, t in zip(edited["Colonne source"], edited["Type (si nouvelle)"])}
        opts = {c: [o.strip() for o in str(s).split("|") if o.strip()] for c, s in zip(edited["Colonne source"], edited["Options (si nouvelle, séparées par « | »)"])}
        st.session_state["imp_plan"] = build_import_plan(df, assign, types, opts, qlist, pq, load_existing_fps(scope))
    plan = st.session_state.get("imp_plan")
    if not plan:
        return
    for e in plan["errors"]:
        st.error(e)
    if plan["errors"]:
        return
    dups = [r for r in plan["accepted"] if r["dup"]]
    kpi_grid([("01", "Lignes lues", str(plan["n_rows"]), up.name), ("02", "Acceptées", str(len(plan["accepted"])), f"dont {len(dups)} doublon(s) détecté(s)"),
              ("03", "Rejetées", str(len(plan["rejected"])), "voir motifs ci-dessous"), ("04", "Avertissements", str(len(plan["warnings"])), "lignes acceptées, à vérifier")])
    t1, t2, t3 = st.tabs([f"Lignes rejetées ({len(plan['rejected'])})", f"Avertissements ({len(plan['warnings'])})", f"Doublons ({len(dups)})"])
    with t1:
        if plan["rejected"]:
            rj = pd.DataFrame(plan["rejected"])
            show_df(rj)
            st.download_button("Télécharger le rapport d'erreurs (CSV)", csv_bytes(rj, [f"Rapport d'erreurs d'import — {up.name}"]), "rapport_erreurs_import.csv", "text/csv")
        else:
            st.write("Aucune ligne rejetée.")
    with t2:
        show_df(pd.DataFrame(plan["warnings"])) if plan["warnings"] else st.write("Aucun avertissement.")
    exclude = set()
    with t3:
        if dups:
            st.caption("Les doublons ne sont jamais exclus sans votre confirmation : décochez « Exclure » pour les conserver.")
            dd = pd.DataFrame([{"Exclure": True, "Ligne": r["line"], "Type de doublon": r["dup"], "Profil": r["group"], "Date": r["ts"] or "—"} for r in dups])
            ed = show_editor(dd, key=f"dup_{sig}", disabled=["Ligne", "Type de doublon", "Profil", "Date"])
            exclude = set(ed.loc[ed["Exclure"], "Ligne"])
        else:
            st.write("Aucun doublon détecté.")
    n_final = len(plan["accepted"]) - len(exclude)
    new_q = [p["label"] for p in plan["cols"] if p["new"]]
    if new_q:
        callout(f"{len(new_q)} nouvelle(s) question(s) seront créées : " + " · ".join(esc(x[:50]) for x in new_q[:6]) + ("…" if len(new_q) > 6 else ""))
    src = st.selectbox("Origine des données", ["Google Forms", "Import fichier"], key="imp_src")
    if st.button(f"Confirmer l'import de {n_final} réponse(s)", key="imp_commit", disabled=n_final == 0):
        res = commit_import(plan, exclude, scope, up.name, src)
        st.session_state["imp_done"] = res
        st.session_state["imp_nonce"] = nonce + 1
        st.session_state.pop("imp_plan", None)
        st.rerun()


# ── Données ─────────────────────────────────────────────────────────────────
def page_donnees(D, F, scope):
    header("Gestion des données", "Recherche, filtres, consultation, correction contrôlée, archivage et export")
    scope_banner(scope)
    resp_f = apply_filters(D["resp_all"], F)
    if D["resp_all"].empty:
        empty_state(scope)
        return
    c1, c2, c3 = st.columns([3, 1, 1])
    query = c1.text_input("Recherche (dans toutes les colonnes)", key="dq")
    show_arch = c2.checkbox("Inclure archivées", False, key="d_arch")
    size = c3.selectbox("Lignes par page", [25, 50, 100], key="d_size")
    if not show_arch:
        resp_f = resp_f[resp_f["archived"] == 0]
    wide = wide_table(D, resp_f)
    if query.strip():
        mask = wide.astype(str).apply(lambda col: col.str.contains(query.strip(), case=False, regex=False)).any(axis=1)
        wide = wide[mask]
    pages = max(1, int(np.ceil(len(wide) / size)))
    page = st.number_input(f"Page (1–{pages})", 1, pages, 1, key="d_page") if pages > 1 else 1
    st.caption(f"{len(wide)} réponse(s) · page {page}/{pages} · tri : cliquer sur l'en-tête d'une colonne")
    show_df(wide.iloc[(page - 1) * size: page * size])
    meta = [("Application", f"{APP_NAME} v{APP_VERSION}"), ("Date de génération", datetime.now().strftime("%d/%m/%Y %H:%M")),
            ("Filtres", describe_filters(F, scope)), ("Recherche", query or "—"), ("Archivées incluses", "oui" if show_arch else "non"), ("Nombre de lignes", len(wide))]
    e1, e2 = st.columns(2)
    e1.download_button("Exporter en CSV", csv_bytes(wide, [f"{APP_NAME} — export du {datetime.now():%d/%m/%Y %H:%M}", describe_filters(F, scope)]),
                       f"donnees_{datetime.now():%Y%m%d_%H%M}.csv", "text/csv", on_click=audit, args=("export CSV", f"{len(wide)} lignes"))
    e2.download_button("Exporter en XLSX", xlsx_bytes({"Données": wide}, meta), f"donnees_{datetime.now():%Y%m%d_%H%M}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", on_click=audit, args=("export XLSX", f"{len(wide)} lignes"))
    st.markdown("---")
    with st.expander("Consulter / corriger / archiver une réponse"):
        ids = list(wide["ID réponse"])
        if not ids:
            st.write("Aucune réponse.")
            return
        rid = st.selectbox("Réponse", ids, key="d_rid")
        row = D["resp_all"][D["resp_all"]["response_id"] == rid].iloc[0]
        a = D["ans"][D["ans"]["response_id"] == rid].set_index("question_id")
        st.markdown(f"<span class='chip'>{esc(rid)}</span><span class='chip'>{esc(row['profile_group'])}</span><span class='chip'>{esc(row['source'])}</span>"
                    f"<span class='chip'>{esc(row['submitted_at'] or 'sans date')}</span>" + ("<span class='chip'>ARCHIVÉE</span>" if row["archived"] else ""), unsafe_allow_html=True)
        detail = [{"Question": q["label"], "Réponse": " | ".join(a.loc[q["question_id"], "vals"]) if q["question_id"] in a.index else "—"} for q in D["qall_list"]]
        show_df(pd.DataFrame(detail))
        b1, b2, b3 = st.columns(3)
        if b1.button("Restaurer" if row["archived"] else "Archiver", key="d_arch_btn") and need_admin("Archivage"):
            with db() as c:
                c.execute("UPDATE responses SET archived=? WHERE response_id=?", (0 if row["archived"] else 1, rid))
            audit("archivage", f"{rid} → {'restaurée' if row['archived'] else 'archivée'}")
            st.rerun()
        if b3.button("Supprimer définitivement", key="d_del") and need_admin("Suppression"):
            st.session_state["d_confirm_del"] = rid
        if st.session_state.get("d_confirm_del") == rid:
            callout(f"Supprimer <b>{esc(rid)}</b> ? Action irréversible.", alert=True)
            if st.button("Oui, supprimer", key="d_del_yes"):
                with db() as c:
                    c.execute("DELETE FROM answers WHERE response_id=?", (rid,))
                    c.execute("DELETE FROM responses WHERE response_id=?", (rid,))
                audit("suppression", rid)
                st.session_state.pop("d_confirm_del", None)
                st.rerun()
        if b2.button("Corriger les réponses", key="d_fix") and need_admin("Correction"):
            st.session_state["d_fixing"] = rid
        if st.session_state.get("d_fixing") == rid and is_admin():
            with st.form(f"fix_{rid}"):
                new = {}
                for q in D["qlist"]:
                    if q["code"] == PROFILE_CODE:
                        continue
                    cur = a.loc[q["question_id"], "vals"] if q["question_id"] in a.index else []
                    default = cur if q["qtype"] == "multiple" else (cur[0] if cur else None)
                    new[q["question_id"]] = question_widget(q, f"fx_{rid}_{q['question_id']}", default)
                ok = st.form_submit_button("Enregistrer les corrections")
            if ok:
                try:
                    with db() as c:
                        for q in D["qlist"]:
                            if q["code"] == PROFILE_CODE:
                                continue
                            enc = encode_answer(q, new[q["question_id"]])
                            if enc is None:
                                c.execute("DELETE FROM answers WHERE response_id=? AND question_id=?", (rid, q["question_id"]))
                            else:
                                c.execute("INSERT OR REPLACE INTO answers(response_id,question_id,raw_value,clean_value) VALUES(?,?,?,?)", (rid, q["question_id"], enc[0], enc[1]))
                    audit("correction", rid)
                    st.session_state.pop("d_fixing", None)
                    st.rerun()
                except ValueError as e:
                    st.error(str(e))


# ── Analyses ────────────────────────────────────────────────────────────────
def page_analyses(D, F, scope):
    header("Statistiques descriptives", "Effectifs, fréquences, taux de non-réponse, tableaux croisés et distribution des réponses")
    scope_banner(scope)
    resp_f = apply_filters(D["resp"], F)
    if resp_f.empty:
        empty_state(scope)
        return
    ans_f = D["ans"][D["ans"]["response_id"].isin(resp_f["response_id"])]
    q = question_picker(D, "Question à analyser", "an_q", types=tuple(QTYPES))
    if q is None:
        callout("Aucune question active.")
        return
    res = compute_question(q, resp_f, ans_f)
    st.markdown(f"<span class='chip'>{esc(q['code'])}</span><span class='chip'>{esc(QTYPES[q['qtype']])}</span>"
                + (f"<span class='chip'>Profils concernés : {esc(', '.join(q['profiles_l']))}</span>" if q["profiles_l"] else "<span class='chip'>Tous profils</span>")
                + ("<span class='chip'>Obligatoire</span>" if q["required"] else ""), unsafe_allow_html=True)
    kpi_grid([("01", "N (réponses valides)", str(res["N"]), "dénominateur des pourcentages"), ("02", "Répondants concernés", str(res["n_conc"]), "après filtres"),
              ("03", "Taux de non-réponse", (fr(100 * res["nonresp"]) + " %") if res["n_conc"] else "—", "réponses vides / concernés"),
              ("04", "Modalités", str(len(res["table"])) if "table" in res else "—", QTYPES[q["qtype"]])])
    if res["N"] == 0:
        callout("Aucune réponse valide à cette question avec les filtres actuels.", alert=True)
        return
    if res["kind"] in CATEGORICAL:
        formula_box()
        if q["qtype"] == "multiple":
            callout("<b>Choix multiples</b> : un répondant peut cocher plusieurs modalités. N = nombre de <b>répondants</b> ; la somme des pourcentages peut dépasser 100 %.")
        t = res["table"]
        disp = t.rename(columns={"n": "Effectif n", "f": "Fréquence f", "p": "Pourcentage p (%)"}).copy()
        disp["Fréquence f"], disp["Pourcentage p (%)"] = disp["Fréquence f"].round(4), disp["Pourcentage p (%)"].round(1)
        disp["N"] = res["N"]
        show_df(disp)
        if q["qtype"] == "ordinale":
            m = res["medians"]
            st.markdown("**Médiane ordinale** (valeur inférieure en cas d'effectif pair) : " + " · ".join(f"{k} : <b>{esc(v) if v else '—'}</b>" for k, v in m.items()), unsafe_allow_html=True)
        pie_ok = q["qtype"] != "multiple" and len(t[t["n"] > 0]) <= 8
        mode = st.radio("Graphique", ["Diagramme en barres", "Diagramme circulaire"] if pie_ok else ["Diagramme en barres"], horizontal=True, key="an_mode")
        show_plot(chart_pie(t, q, q["label"]) if mode == "Diagramme circulaire" else chart_bar(t, res["N"], q["label"]), key="an_chart")
        st.markdown("#### Tableau croisé selon le profil")
        st.caption("Effectif (pourcentage au sein du groupe). Le dénominateur de chaque groupe est son effectif de réponses valides.")
        show_df(matrix_table(res))
        st.download_button("Exporter ce tableau (CSV)", csv_bytes(stats_long([res]), [f"{APP_NAME} — {q['label']}", describe_filters(F, scope)]), "tableau_statistique.csv", "text/csv",
                           on_click=audit, args=("export CSV", f"statistiques {q['code']}"))
    elif res["kind"] == "nombre":
        n_ = res["numbers"]
        if n_.empty:
            callout("Aucune valeur numérique exploitable.", alert=True)
            return
        show_df(pd.DataFrame([{"N": len(n_), "Moyenne": round(n_.mean(), 2), "Médiane": round(n_.median(), 2), "Écart-type": round(n_.std(ddof=1), 2) if len(n_) > 1 else None,
                               "Min": n_.min(), "Max": n_.max()}]))
        fig = go.Figure(go.Histogram(x=n_, marker_color=CYAN, hovertemplate="Valeurs : %{x}<br>Effectif : %{y}<extra></extra>"))
        show_plot(style_fig(fig, q["label"], "Valeur", "Effectif", 380), key="an_hist")
    else:
        callout("Question à réponse libre : analyse qualitative. Les verbatims sont affichés ci-dessous (à manipuler avec précaution).")
        tx = res["texts"].copy()
        tx["Réponse"] = tx["vals"].map(lambda v: " ".join(v))
        show_df(tx[["profile_group", "Réponse"]].rename(columns={"profile_group": "Profil"}))


# ── Comparaison ─────────────────────────────────────────────────────────────
def page_comparaison(D, F, scope):
    header("Comparaison des profils", "Professionnels / Clients / Observateurs : mêmes questions, dénominateurs explicites")
    scope_banner(scope)
    resp_f = apply_filters(D["resp"], F)
    if resp_f.empty:
        empty_state(scope)
        return
    ans_f = D["ans"][D["ans"]["response_id"].isin(resp_f["response_id"])]
    sizes = resp_f["profile_group"].value_counts().reindex(PROFILES, fill_value=0)
    st.markdown("".join(f"<span class='chip'>{g} : n = {int(sizes[g])}</span>" for g in PROFILES), unsafe_allow_html=True)
    small = [g for g in PROFILES if 0 < sizes[g] < 10]
    if small:
        callout("Effectif faible (n < 10) pour : <b>" + ", ".join(small) + "</b>. Interprétez les pourcentages avec prudence.", alert=True)
    q = question_picker(D, "Question à comparer", "cmp_q", include_profile=False)
    if q is None:
        callout("Aucune question à choix à comparer : importez les réponses du questionnaire complet.")
        return
    res = compute_question(q, resp_f, ans_f)
    if sum(1 for g in PROFILES if res["Ng"][g] > 0) == 0:
        callout("Aucune réponse valide à cette question avec les filtres actuels.", alert=True)
        return
    formula_box()
    st.caption("Les pourcentages sont calculés au sein de chaque groupe (et non sur les effectifs bruts), car les tailles de groupes peuvent différer.")
    if q["qtype"] == "multiple":
        callout("<b>Choix multiples</b> : le total des choix peut dépasser le nombre de répondants ; les barres empilées à 100 % ne s'appliquent pas.")
    views = ["Barres groupées", "Matrice (carte de chaleur)"] if q["qtype"] == "multiple" else ["Barres empilées à 100 %", "Barres groupées", "Matrice (carte de chaleur)"]
    view = st.radio("Vue", views, horizontal=True, key="cmp_view")
    fig = chart_stack(res, q["label"]) if view.startswith("Barres empilées") else chart_group(res, q["label"]) if view == "Barres groupées" else chart_heat(res, q["label"])
    show_plot(fig, key="cmp_fig")
    st.markdown("#### Matrice de comparaison")
    show_df(matrix_table(res))
    if q["qtype"] == "ordinale":
        st.markdown("**Médiane ordinale par profil** : " + " · ".join(f"{k} : <b>{esc(v) if v else '—'}</b>" for k, v in res["medians"].items()), unsafe_allow_html=True)
    st.markdown("#### Convergences et divergences (toutes les questions à choix)")
    allres = [compute_question(x, resp_f, ans_f) for x in D["qlist"] if x["qtype"] in CATEGORICAL and x["code"] != PROFILE_CODE]
    conv = convergence_table(allres)
    if conv.empty:
        st.write("Pas assez de groupes renseignés pour comparer.")
    else:
        show_df(conv)
        st.caption("Écart max. = plus grande différence, en points de pourcentage, entre deux profils sur une même modalité. Indicateur descriptif : "
                   "aucun test d'inférence n'est réalisé et aucune relation de cause à effet ne peut en être déduite.")


# ── Rapports ────────────────────────────────────────────────────────────────
def page_rapports(D, F, scope):
    header("Rapports et exports", "Rapport PDF (indicateurs + graphiques) et classeur Excel — filtres et date de génération inclus")
    scope_banner(scope)
    resp_f = apply_filters(D["resp"], F)
    if resp_f.empty:
        empty_state(scope)
        return
    ans_f = D["ans"][D["ans"]["response_id"].isin(resp_f["response_id"])]
    title = st.text_input("Titre du rapport", "Enquête BTP — perceptions et pratiques", key="rp_title")
    qs = {f"[{q['code']}] {q['label']}": q for q in D["qlist"]}
    default = [k for k, q in qs.items() if q["qtype"] in CATEGORICAL]
    chosen = st.multiselect("Questions à inclure", list(qs), default=default, key=f"rp_q_{scope}_{len(qs)}")
    with_groups = st.checkbox("Inclure la comparaison par profil", True, key="rp_groups")
    st.caption("Les rapports privilégient les résultats agrégés ; les réponses libres ne sont pas reproduites.")
    if st.button("Générer le rapport PDF et le classeur Excel", key="rp_go") and chosen:
        with st.spinner("Calcul des statistiques et mise en page…"):
            results = [compute_question(qs[k], resp_f, ans_f) for k in chosen]
            counts = resp_f["profile_group"].value_counts().reindex(PROFILES, fill_value=0)
            ftxt = describe_filters(F, scope)
            kpis = [("Réponses valides (N)", len(resp_f)), ("Effectif total (archivées incluses)", len(apply_filters(D["resp_all"], F)))] + \
                   [(f"Profil {g}", f"{int(counts[g])} ({fr(100 * counts[g] / len(resp_f))} %)") for g in PROFILES] + [("Questions incluses", len(results))]
            pdf = build_pdf(title, scope, ftxt, kpis, results, with_groups)
            meta = [("Application", f"{APP_NAME} v{APP_VERSION}"), ("Titre", title), ("Date de génération", datetime.now().strftime("%d/%m/%Y %H:%M")), ("Filtres", ftxt)]
            xlsx = xlsx_bytes({"Données": wide_table(D, resp_f), "Statistiques": stats_long(results)}, meta)
        st.session_state["rp_out"] = (pdf, xlsx, datetime.now().strftime("%Y%m%d_%H%M"))
        audit("rapport", f"{len(results)} questions · {len(resp_f)} réponses")
    out = st.session_state.get("rp_out")
    if out:
        c1, c2 = st.columns(2)
        c1.download_button("⬇ Rapport PDF", out[0], f"rapport_{out[2]}.pdf", "application/pdf", on_click=audit, args=("export PDF", out[2]))
        c2.download_button("⬇ Classeur Excel (données + statistiques)", out[1], f"export_{out[2]}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", on_click=audit, args=("export XLSX", out[2]))


# ── Paramètres ──────────────────────────────────────────────────────────────
DEMO_QUESTIONS = [
    ("D1", "[DEMO] Fréquence d'utilisation de matériaux locaux", "ordinale", ["Jamais", "Rarement", "Parfois", "Souvent", "Toujours"], []),
    ("D2", "[DEMO] Principales difficultés rencontrées sur un chantier", "multiple", ["Coûts", "Délais", "Main-d'œuvre", "Réglementation", "Qualité des matériaux"], []),
    ("D3", "[DEMO] La sécurité sur les chantiers est suffisante", "ordinale", ["Pas du tout d'accord", "Plutôt pas d'accord", "Neutre", "Plutôt d'accord", "Tout à fait d'accord"], []),
    ("D4", "[DEMO] Le budget initial a-t-il été respecté ?", "unique", ["Oui", "Non", "Partiellement"], ["Professionnel", "Client"]),
    ("D5", "[DEMO] Nombre de chantiers suivis ou réalisés", "nombre", [], []),
    ("D6", "[DEMO] Commentaire libre", "texte", [], []),
]


def generate_demo(n: int = 120):
    rng = np.random.default_rng(42)
    with db() as c:
        pq = ensure_profile_question(c, 1)
        qids = {}
        for i, (code, label, qt, opts, profs) in enumerate(DEMO_QUESTIONS, start=1):
            row = c.execute("SELECT question_id FROM questions WHERE code=? AND is_demo=1", (code,)).fetchone()
            qids[code] = row[0] if row else c.execute(
                "INSERT INTO questions(code,label,qtype,options,profiles,required,position,active,is_demo) VALUES(?,?,?,?,?,0,?,1,1)",
                (code, label, qt, json.dumps(opts, ensure_ascii=False), json.dumps(profs, ensure_ascii=False), i)).lastrowid
        batch = "DEMO-" + datetime.now().strftime("%H%M%S")
        for _ in range(n):
            g = str(rng.choice(PROFILES, p=[.4, .35, .25]))
            ts = (datetime.now() - timedelta(days=int(rng.integers(0, 60)), hours=int(rng.integers(0, 24)), minutes=int(rng.integers(0, 60)))).strftime("%Y-%m-%d %H:%M:%S")
            a, items = {pq: (PROFILE_LABELS[g], PROFILE_LABELS[g])}, {PROFILE_QUESTION_LABEL: PROFILE_LABELS[g]}
            shift = {"Professionnel": 0, "Client": 1, "Observateur": 2}[g]
            for code, label, qt, opts, profs in DEMO_QUESTIONS:
                if profs and g not in profs:
                    continue
                if rng.random() < 0.06:
                    continue
                if qt == "ordinale":
                    k = int(np.clip(round(rng.normal(2.2 + (0.5 if code == "D1" and g == "Professionnel" else 0) + (0.3 if code == "D3" and g == "Professionnel" else -0.2 * shift), 1.1)), 0, 4))
                    v = (opts[k], opts[k])
                elif qt == "unique":
                    o = opts[int(rng.choice(3, p=[.5, .2, .3] if g == "Professionnel" else [.3, .3, .4]))]
                    v = (o, o)
                elif qt == "multiple":
                    ch = [o for o in opts if rng.random() < (.55 if o in ("Coûts", "Délais") else .3)] or [opts[0]]
                    v = (", ".join(ch), json.dumps(ch, ensure_ascii=False))
                elif qt == "nombre":
                    x = str(int(rng.integers(1, 40) if g == "Professionnel" else rng.integers(0, 6)))
                    v = (x, x)
                else:
                    v = ("Commentaire fictif de démonstration.", "Commentaire fictif de démonstration.")
                a[qids[code]] = v
                items[label] = v[1]
            insert_response(c, group=g, answers=a, source="Démonstration", submitted_at=ts, batch_id=batch, is_demo=1, fp=make_fp(g, items))
    audit("démo", f"{n} réponses fictives générées")


def delete_demo():
    with db() as c:
        c.execute("DELETE FROM answers WHERE response_id IN (SELECT response_id FROM responses WHERE is_demo=1)")
        c.execute("DELETE FROM responses WHERE is_demo=1")
        c.execute("DELETE FROM column_map WHERE question_id IN (SELECT question_id FROM questions WHERE is_demo=1)")
        c.execute("DELETE FROM questions WHERE is_demo=1")
        c.execute("DELETE FROM import_batches WHERE is_demo=1")
    audit("démo", "données de démonstration supprimées")


def db_snapshot() -> bytes:
    fd, tmp = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    try:
        src, dst = sqlite3.connect(DB_PATH), sqlite3.connect(tmp)
        src.backup(dst)
        dst.close()
        src.close()
        with open(tmp, "rb") as fh:
            return fh.read()
    finally:
        os.remove(tmp)


def restore_snapshot(data: bytes):
    fd, tmp = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    try:
        with open(tmp, "wb") as fh:
            fh.write(data)
        con = sqlite3.connect(tmp)
        names = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        con.close()
        if not {"questions", "responses", "answers"} <= names:
            raise ValueError("Ce fichier n'est pas une sauvegarde CIVI-MATH valide.")
        if os.path.exists(DB_PATH):
            shutil.copyfile(DB_PATH, f"{DB_PATH}.avant-restauration-{datetime.now():%Y%m%d%H%M%S}")
        shutil.copyfile(tmp, DB_PATH)
    finally:
        os.remove(tmp)


def page_parametres(D, scope):
    header("Paramètres", "Questionnaire, accès, données de démonstration, sauvegarde, journal")
    scope_banner(scope)
    t1, t2, t3, t4, t5 = st.tabs(["Questionnaire", "Accès administrateur", "Démonstration", "Sauvegarde", "Journal & infos"])
    with t1:
        st.caption("Dictionnaire des variables : une question = une variable. Les noms de colonnes d'origine sont conservés dans la table de correspondance et "
                   "aucune question n'est fusionnée automatiquement. Options séparées par « | » ; profils séparés par des virgules (vide = tous).")
        qall = D["qall"]
        if qall.empty:
            st.write("Aucune question.")
        else:
            cnt = D["ans"].groupby("question_id").size().to_dict()
            ed = pd.DataFrame({
                "ID": qall["question_id"], "Code": qall["code"], "Libellé": qall["label"], "Type": qall["qtype"].map(QTYPES),
                "Options": qall["options_l"].map(lambda o: " | ".join(o)), "Profils concernés": qall["profiles_l"].map(", ".join),
                "Obligatoire": qall["required"].astype(bool), "Position": qall["position"], "Actif": qall["active"].astype(bool),
                "Réponses": qall["question_id"].map(lambda i: cnt.get(i, 0))})
            edited = show_editor(ed, key=f"qed_{scope}_{len(qall)}", disabled=["ID", "Code", "Réponses"],
                                 column_config={"Type": st.column_config.SelectboxColumn("Type", options=list(QTYPES.values()), required=True)})
            if st.button("Enregistrer les modifications du questionnaire", key="q_save") and need_admin("Modification du questionnaire"):
                errs = []
                with db() as c:
                    for _, r in edited.iterrows():
                        old = qall[qall["question_id"] == r["ID"]].iloc[0]
                        qt = QTYPE_BY_LABEL.get(r["Type"], old["qtype"])
                        profs = [p.strip() for p in str(r["Profils concernés"]).split(",") if p.strip()]
                        if any(p not in PROFILES for p in profs):
                            errs.append(f"{r['Code']} : profils valides = {', '.join(PROFILES)}")
                            continue
                        if r["Code"] == PROFILE_CODE:
                            c.execute("UPDATE questions SET position=?, active=1 WHERE question_id=?", (int(r["Position"]), int(r["ID"])))
                            continue
                        if r["Réponses"] and ((old["qtype"] == "multiple") != (qt == "multiple")):
                            errs.append(f"{r['Code']} : changement de type incompatible avec les réponses existantes")
                            continue
                        opts = [o.strip() for o in str(r["Options"]).split("|") if o.strip()]
                        c.execute("UPDATE questions SET label=?, qtype=?, options=?, profiles=?, required=?, position=?, active=? WHERE question_id=?",
                                  (clean_spaces(r["Libellé"]), qt, json.dumps(opts, ensure_ascii=False), json.dumps(profs, ensure_ascii=False),
                                   int(bool(r["Obligatoire"])), int(r["Position"]), int(bool(r["Actif"])), int(r["ID"])))
                if errs:
                    for e in errs:
                        st.error(e)
                else:
                    audit("questionnaire", "modification du dictionnaire")
                    st.rerun()
        with st.expander("➕ Ajouter une question"):
            with st.form("q_add"):
                code = st.text_input("Code (facultatif, ex. Q12)")
                label = st.text_input("Libellé de la question")
                qt = st.selectbox("Type", list(QTYPES.values()))
                opts = st.text_area("Options (une par ligne, dans l'ordre pour une échelle ordinale)")
                profs = st.multiselect("Profils concernés (vide = tous)", PROFILES)
                req = st.checkbox("Réponse obligatoire")
                ok = st.form_submit_button("Ajouter")
            if ok and need_admin("Ajout de question"):
                if not clean_spaces(label):
                    st.error("Le libellé est obligatoire.")
                else:
                    with db() as c:
                        pos = c.execute("SELECT COALESCE(MAX(position),0)+1 FROM questions WHERE is_demo=?", (scope,)).fetchone()[0]
                        code = clean_spaces(code) or f"Q{pos}"
                        try:
                            c.execute("INSERT INTO questions(code,label,qtype,options,profiles,required,position,active,is_demo) VALUES(?,?,?,?,?,?,?,1,?)",
                                      (code, clean_spaces(label), QTYPE_BY_LABEL[qt], json.dumps([o.strip() for o in opts.splitlines() if o.strip()], ensure_ascii=False),
                                       json.dumps(profs, ensure_ascii=False), int(req), pos, scope))
                            audit("questionnaire", f"ajout {code}")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error(f"Le code « {code} » existe déjà.")
        with st.expander("Table de correspondance des colonnes d'origine"):
            with db() as c:
                cm = pd.read_sql_query("SELECT m.original_column AS 'Colonne d''origine', q.code AS Code, q.label AS Question FROM column_map m JOIN questions q USING(question_id) WHERE q.is_demo=?", c, params=(scope,))
            show_df(cm) if not cm.empty else st.write("Aucune correspondance enregistrée.")
    with t2:
        pw = get_admin_password()
        if pw is None:
            callout("Mode <b>local</b> : aucun mot de passe administrateur n'est défini, tous les utilisateurs ont les droits d'administration. "
                    "Pour un usage multi-utilisateur, définissez la variable d'environnement <code>CIVIMATH_ADMIN_PASSWORD</code> "
                    "(ou un secret Streamlit du même nom) — ne placez jamais de mot de passe dans le code.", alert=True)
        elif st.session_state.get("admin_ok"):
            st.success("Connecté en administrateur.")
            if st.button("Se déconnecter", key="adm_out"):
                st.session_state["admin_ok"] = False
                st.rerun()
        else:
            with st.form("adm"):
                x = st.text_input("Mot de passe administrateur", type="password")
                if st.form_submit_button("Se connecter"):
                    if hmac.compare_digest(x.encode(), pw.encode()):
                        st.session_state["admin_ok"] = True
                        audit("connexion admin", "succès")
                        st.rerun()
                    else:
                        audit("connexion admin", "échec")
                        st.error("Mot de passe incorrect.")
        st.markdown("**Droits** — consultation, saisie et import : tous ; modification du questionnaire, correction, archivage, suppression, restauration : administrateur.")
    with t3:
        callout("Le jeu de démonstration est <b>fictif</b>, stocké dans un périmètre séparé (bascule « Démonstration » dans le menu latéral) et ne se mélange jamais aux réponses réelles.")
        n = st.slider("Nombre de réponses fictives", 30, 400, 120, 10, key="demo_n")
        d1, d2 = st.columns(2)
        if d1.button("Générer le jeu de démonstration", key="demo_gen"):
            generate_demo(n)
            st.success("Jeu de démonstration créé. Basculez sur « Démonstration » dans le menu latéral pour l'explorer.")
        if d2.button("Supprimer les données de démonstration", key="demo_del") and need_admin("Suppression des données de démonstration"):
            delete_demo()
            st.success("Données de démonstration supprimées.")
    with t4:
        st.caption("Sauvegardez régulièrement la base. La restauration remplace la base actuelle (une copie de sécurité est conservée à côté du fichier).")
        st.download_button("⬇ Télécharger une sauvegarde de la base (.db)", db_snapshot(), f"civimath_sauvegarde_{datetime.now():%Y%m%d_%H%M}.db",
                           "application/octet-stream", on_click=audit, args=("sauvegarde", "téléchargement"))
        up = st.file_uploader("Restaurer depuis une sauvegarde (.db)", type=["db", "sqlite"], key="restore_up")
        if up is not None:
            sure = st.checkbox("Je confirme le remplacement de la base actuelle", key="restore_ok")
            if st.button("Restaurer", disabled=not sure, key="restore_go") and need_admin("Restauration"):
                try:
                    restore_snapshot(up.getvalue())
                    init_db()
                    audit("restauration", up.name)
                    st.success("Base restaurée.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Restauration impossible : {e}")
    with t5:
        with db() as c:
            log = pd.read_sql_query("SELECT ts AS Date, action AS Action, details AS Détails FROM audit_log ORDER BY id DESC LIMIT 200", c)
            lots = pd.read_sql_query("SELECT import_batch_id AS Lot, filename AS Fichier, source AS Origine, imported_at AS Date, n_rows AS Lignes, n_ok AS Acceptées, n_rejected AS Rejetées, n_excluded AS Exclues, is_demo AS Test FROM import_batches ORDER BY imported_at DESC LIMIT 50", c)
        st.markdown("**Historique des imports**")
        show_df(lots) if not lots.empty else st.write("Aucun import.")
        st.markdown("**Journal des opérations** (imports, exports, modifications)")
        show_df(log) if not log.empty else st.write("Journal vide.")
        callout("<b>Mode B — API Google Forms</b> : non activé dans cette version. Un lien public de réponse ne donne pas les droits de lecture des réponses ; "
                "la synchronisation automatique nécessite un projet Google Cloud, une authentification et un accès autorisé au formulaire. "
                "L'import CSV/XLSX (Mode A) couvre le besoin de la première version.")
        st.caption(f"{APP_NAME} v{APP_VERSION} · base SQLite : {os.path.abspath(DB_PATH)} · Les identifiants ne doivent jamais figurer dans le code public.")


# ════════════════════════════════════════════════════════════════════════════
# 12. COUVERTURE + POINT D'ENTRÉE
# ════════════════════════════════════════════════════════════════════════════
def _enter():
    st.session_state["entered"] = True


def _leave():
    st.session_state["entered"] = False


def render_cover():
    st.markdown("<style>[data-testid='stSidebar'],[data-testid='stSidebarCollapsedControl'],[data-testid='collapsedControl']{display:none!important}"
                ".block-container{max-width:1100px!important}</style>", unsafe_allow_html=True)
    if len(COVER_B64) > 100:
        st.markdown(f'<div class="cover-wrap"><img alt="Une enquête, plusieurs regards… Professionnels du BTP, clients, observateurs : 3 profils, 1 même réalité ?" '
                    f'src="data:image/jpeg;base64,{COVER_B64}"/></div>', unsafe_allow_html=True)
    else:
        header("Une enquête, plusieurs regards…", "Professionnels du BTP · Clients · Observateurs")
    st.markdown("&nbsp;")
    _, mid, _ = st.columns([1, 2, 1])
    with mid:
        st.button("Entrer dans CIVI-MATH ANALYTICS  →", key="enter", on_click=_enter)
    st.markdown(f'<div class="cover-cap">{APP_NAME} · LABORATOIRE STATISTIQUE · BTP &amp; GÉNIE CIVIL · v{APP_VERSION}</div>', unsafe_allow_html=True)


def main():
    st.set_page_config(page_title=APP_NAME, page_icon="📐", layout="wide")
    init_db()
    st.markdown(CSS, unsafe_allow_html=True)
    if not st.session_state.get("entered"):
        render_cover()
        return
    with st.sidebar:
        st.markdown(f'<div class="brand">CIVI-MATH<br>ANALYTICS<small>LABORATOIRE STATISTIQUE</small></div>', unsafe_allow_html=True)
        st.markdown("&nbsp;")
        page = st.radio("Navigation", PAGES, key="nav", label_visibility="collapsed", format_func=lambda p: f"{PAGE_ICONS[p]}  {p}")
        st.markdown("---")
        mode = st.radio("Données affichées", ["Réelles", "Démonstration (test)"], key="scope_radio")
    scope = 1 if mode.startswith("Démo") else 0
    D = load_context(scope)
    F = sidebar_filters(D)
    st.sidebar.markdown("---")
    st.sidebar.button("↩ Page de couverture", key="back_cover", on_click=_leave)
    st.sidebar.caption(f"v{APP_VERSION} · " + ("mode administrateur" if is_admin() else "mode consultation"))
    if page == "Vue d'ensemble":
        page_overview(D, F, scope)
    elif page == "Collecte":
        page_collecte(D, scope)
    elif page == "Import":
        page_import(D, scope)
    elif page == "Données":
        page_donnees(D, F, scope)
    elif page == "Analyses":
        page_analyses(D, F, scope)
    elif page == "Comparaison":
        page_comparaison(D, F, scope)
    elif page == "Rapports":
        page_rapports(D, F, scope)
    else:
        page_parametres(D, scope)


# ════════════════════════════════════════════════════════════════════════════
# 13. AUTO-TESTS  (python app.py --selftest)
# ════════════════════════════════════════════════════════════════════════════
def run_selftests() -> int:
    global DB_PATH
    tmp = tempfile.mkdtemp()
    DB_PATH = os.path.join(tmp, "test.db")
    fails = []

    def check(name, cond):
        print(("  OK   " if cond else "  ÉCHEC"), name)
        if not cond:
            fails.append(name)

    print("Profils")
    check("professionnel", normalize_profile("Professionnel du BTP") == "Professionnel")
    check("client", normalize_profile(PROFILE_LABELS["Client"]) == "Client")
    check("observateur", normalize_profile(PROFILE_LABELS["Observateur"]) == "Observateur")
    check("inconnu", normalize_profile("Autre") is None and normalize_profile("") is None)

    print("Choix multiples avec virgules dans les options")
    opts = ["Coûts", "Délais", "Autre, précisez"]
    v, unexpected = parse_multi("Coûts, Autre, précisez, Délais, Bruit", opts)
    check("options reconnues", set(v) == {"Coûts", "Délais", "Autre, précisez", "Bruit"} and unexpected == ["Bruit"])

    print("Statistiques")
    init_db()
    qd = dict(question_id=1, code="Q1", label="Q", qtype="unique", options_l=["A", "B", "C"], profiles_l=[], required=0)
    a = pd.DataFrame({"response_id": ["r1", "r2", "r3", "r4"], "question_id": 1, "vals": [["A"], ["A"], ["B"], []]})
    a = a[a["vals"].map(len) > 0]
    t, N = freq_table(qd, a)
    check("N = 3", N == 3)
    check("effectifs A/B/C = 2/1/0", list(t["n"]) == [2, 1, 0])
    check("p(A) = 66,67", abs(t.loc[0, "p"] - 66.6667) < 1e-3 and abs(t.loc[0, "f"] - 2 / 3) < 1e-9)
    check("ordinal median", ordinal_median(["B", "A", "C"], ["A", "B", "C"]) == "B")

    print("Import")
    df = pd.DataFrame({
        "Horodateur": ["28/10/2026 10:00:00", "28/10/2026 10:05:00", "28/10/2026 10:05:00", "29/10/2026 09:00:00", "", "30/10/2026 08:00:00"],
        "Niveau de proximité": ["Professionnel du BTP", PROFILE_LABELS["Client"], PROFILE_LABELS["Client"], "Inconnu", "", PROFILE_LABELS["Observateur"]],
        "Difficultés": ["Coûts, Délais", "Coûts", "Coûts", "Délais", "", ""],
        "Satisfaction": ["Oui", "Non", "Non", "Oui", "", "Peut-être"],
    })
    types = {"Difficultés": "multiple", "Satisfaction": "unique"}
    optsn = {"Difficultés": ["Coûts", "Délais"], "Satisfaction": ["Oui", "Non"]}
    qlist = load_questions(0).to_dict("records")
    pq = get_profile_qid(0)
    assign = {"Horodateur": A_TIME, "Niveau de proximité": q_display(next(q for q in qlist if q["question_id"] == pq)), "Difficultés": A_NEW, "Satisfaction": A_NEW}
    plan = build_import_plan(df, assign, types, optsn, qlist, pq, {})
    check("3 acceptées + doublon, 2 rejetées", len(plan["accepted"]) == 4 and len(plan["rejected"]) == 2)
    check("doublon exact détecté", sum(1 for r in plan["accepted"] if r["dup"]) == 1)
    check("numéros de ligne (fichier)", {r["Ligne"] for r in plan["rejected"]} == {5, 6})
    check("valeur inattendue signalée", any("inattendue" in w["Avertissement"] for w in plan["warnings"]))
    dup_line = next(r["line"] for r in plan["accepted"] if r["dup"])
    res = commit_import(plan, {dup_line}, 0, "test.csv", "Import fichier")
    check("import : 3 réponses enregistrées", res["ok"] == 3 and res["excluded"] == 1)
    D = load_context(0)
    check("3 réponses en base", len(D["resp"]) == 3)
    check("profils corrects", sorted(D["resp"]["profile_group"]) == ["Client", "Observateur", "Professionnel"])
    qm = next(q for q in D["qlist"] if q["label"] == "Difficultés")
    r = compute_question(qm, D["resp"], D["ans"])
    check("multiple : N = 2 répondants, Coûts = 2", r["N"] == 2 and int(r["table"].set_index("Modalité").loc["Coûts", "n"]) == 2)
    check("multiple : pourcentage Coûts = 100 %", abs(r["table"].set_index("Modalité").loc["Coûts", "p"] - 100) < 1e-9)
    check("non-réponse = 1/3", abs(r["nonresp"] - 1 / 3) < 1e-9)
    check("pct au sein du groupe (Client Coûts = 100)", abs(r["pct"].loc["Coûts", "Client"] - 100) < 1e-9)
    plan2 = build_import_plan(df, assign, types, optsn, load_questions(0).to_dict("records"), pq, load_existing_fps(0))
    check("réimport : doublons détectés en base", sum(1 for x in plan2["accepted"] if x["dup"]) >= 3)
    check("profil manquant → erreur", bool(build_import_plan(df, {**assign, "Niveau de proximité": A_IGNORE}, types, optsn, qlist, pq, {})["errors"]))

    print("Démonstration séparée des données réelles")
    generate_demo(40)
    check("démo = 40, réel = 3", len(load_context(1)["resp"]) == 40 and len(load_context(0)["resp"]) == 3)
    pdf = build_pdf("Test", 1, "filtres", [("N", 40)], [compute_question(q, load_context(1)["resp"], load_context(1)["ans"]) for q in load_context(1)["qlist"][:4]], True)
    check("PDF généré", pdf[:4] == b"%PDF")
    delete_demo()
    check("suppression démo", len(load_context(1)["resp"]) == 0 and len(load_context(0)["resp"]) == 3)
    shutil.rmtree(tmp, ignore_errors=True)
    print("\nRésultat :", "TOUS LES TESTS PASSENT" if not fails else f"{len(fails)} échec(s)")
    return 1 if fails else 0


# ── Image de couverture (JPEG fournie par l'utilisateur, encodée en base64) ──
COVER_B64 = (
    "/9j/4AAQSkZJRgABAQAASABIAAD/4QDsRXhpZgAATU0AKgAAAAgAAYdpAAQAAAABAAAAGgAAAAAAC5ADAAIAAAAUAAAApJAEAAIAAAAUAAAAuJAQAAIAAAAH"
    "AAAAzJARAAIAAAAHAAAA1JASAAIAAAAHAAAA3JKQAAIAAAAEMDAwAJKRAAIAAAAEMDAwAJKSAAIAAAAEMDAwAKABAAMAAAABAAEAAKACAAQAAAABAAAE8KAD"
    "AAQAAAABAAADSwAAAAAyMDI2OjEwOjI4IDEzOjQyOjA4ADIwMjY6MTA6MjggMTM6NDI6MDgAKzAxOjAwAAArMDE6MDAAACswMTowMAAA/+0AfFBob3Rvc2hv"
    "cCAzLjAAOEJJTQQEAAAAAABEHAFaAAMbJUccAgAAAgACHAI/AAYxMzQyMDgcAj4ACDIwMjYxMDI4HAI3AAgyMDI2MTAyOBwCPAALMTM0MjA4KzAxMDA4QklN"
    "BCUAAAAAABC4ekS0X6FIGX25K7KKLlsK/8IAEQgDSwTwAwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAMCBAEFAAYHCAkKC//EAMMQAAEDAwIEAwQG"
    "BAcGBAgGcwECAAMRBBIhBTETIhAGQVEyFGFxIweBIJFCFaFSM7EkYjAWwXLRQ5I0ggjhU0AlYxc18JNzolBEsoPxJlQ2ZJR0wmDShKMYcOInRTdls1V1pJXD"
    "hfLTRnaA40dWZrQJChkaKCkqODk6SElKV1hZWmdoaWp3eHl6hoeIiYqQlpeYmZqgpaanqKmqsLW2t7i5usDExcbHyMnK0NTV1tfY2drg5OXm5+jp6vP09fb3"
    "+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAQIAAwQFBgcICQoL/8QAwxEAAgIBAwMDAgMFAgUCBASHAQACEQMQEiEEIDFBEwUwIjJRFEAGMyNhQhVxUjSBUCSR"
    "oUOxFgdiNVPw0SVgwUThcvEXgmM2cCZFVJInotIICQoYGRooKSo3ODk6RkdISUpVVldYWVpkZWZnaGlqc3R1dnd4eXqAg4SFhoeIiYqQk5SVlpeYmZqgo6Sl"
    "pqeoqaqwsrO0tba3uLm6wMLDxMXGx8jJytDT1NXW19jZ2uDi4+Tl5ufo6ery8/T19vf4+fr/2wBDAAICAgICAgMCAgMFAwMDBQYFBQUFBggGBgYGBggKCAgI"
    "CAgICgoKCgoKCgoMDAwMDAwODg4ODg8PDw8PDw8PDw//2wBDAQIDAwQEBAcEBAcQCwkLEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQ"
    "EBAQEBAQEBAQEBD/2gAMAwEAAhEDEQAAAfinbfVeDttW21bbU4buG4ttjbbC22NttW21bbVttW21bbVttWUmaxRkWDtqykqpO01p0VKkrlRtpttFTomojatt"
    "q22rbJqZ2rbattq22rbTW2itspaJ2a2TqVtqnaaTtpYnabbaXZSa201G2qdOqJ2pW2l2yqiVJl22rbZrJVlbKymWNOrZWpM6IZSdDbattq22pSVak5UAxK4D"
    "IyohGmYREzFKlJrZUGnaRRlaolWMlSkm22gmFxCFaaTJNQ5XoU+2umw2YsH0MMtYxX4101Sxg/mvyl/mGp/mGp/mGp/NfqfqrtT/ADDU/wAw0X2Zpg9zLReZ"
    "nqdmYOSIhtgXKRQIkJ1aJ1YgyED2wG2020RUSlVbbVMbVonVttW21bbVttW21bbLKToNO2NpnVtopelUottW20220u2mp21RO1RMTW2VW2VKnETW20u21KSq"
    "K07VlRNTKZZpSrKEyuKTBIpMq1J0zUTppClaohWiicqCMqYphcUhUzBKolqJVgU5S6GpeIyVKgiCRQ5UqA5WqI5LFDks0GTxXPba2tGLqvgWQ6JsHVZs3DAg"
    "8BwJsHUWQ4TlAcY2DljYOMYjXUaBYRsHUWQ6jYOaO5r3dCSlIi4WERQooyIipWhRkxpFGnVEKTW21bbVtoqdtW21bbVtsttM0nbVts1E7VtlVttW2wpnKaHt"
    "IWMrUmZ1bbVphVRCtWTM1tprTtLtlUnZVbaahUTW20umdWmdNoXqTM6kq2qNOm0Kmk7TLGViU5WinKUZErwhyTSjleIQqdDaZpKpVSMfUKD6IpLqFJdAalak"
    "5WITK0m2VMvMbZemxrrGuhtlRTtqfsH7AjbYHZSa22rbYW21bbG22FtsLbattq22rPmL9gyicpSqJqNMVO2rKSSh7RU7ak5Sa22rbattq201GyqTsqkzClom"
    "JpOmaTphrTOrbattpdtqXtNJ2UsnTDSkzFTKVVtEw0wqOidLttW2VWiZpMzNROVNG0yxOmWFZU+ytFOmKnToRp0dpitlSZGVJCFSsAcrxESpMImdSZXpUYuo"
    "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"
    "vtd1it/BsaJtEp8YX6bi78Kf7RvGP+0rwXdJRdqAUJfBtouazt7SwRvv+0jwibNc13tFld2yvBm5Be0bcdtsvGV0hFp4R/2j+MP9pHhDcU291c28dzFP4JuQ"
    "vaPDltt6PEidvs9t8HXCYb7Qi48IWcstjaWm3Rbr/tM8Lf7WBxuf8ZtZUQ3MVpt60br4VXPc7BsE+3TXFwi2hWsrWX4c43v+Jva7oWd30qTceHVFe37JHbKv"
    "0WVta7XdC0u9CLrZba5Xa2dtYB2KwjeQ59juOZt23CwC/Z23/Hne/wCOaoNvKm4ivbBF47baUxSVBe/fvrNNtcW1/tSbl22yTCSjuJBPeR2NrCncoU2022Ws"
    "cyL6yh93t0c+YWdtjcRCCf71C6OjxYDxDCXR0YS8XR4FhDwDCQ7yK8S47VUSaMB07aOjxdC6dlaJ2/8AxEfu6VeLxev81R0L8SLEDMNzuknh23jn3Xf70XO0"
    "7GIINv2y7N74mFHR4lphWtyjkixRJM/c7a2X7wu7tLVU0SE3pAVPm9GaVg46PSmjL81/u2n2vLfN1udp3xHjLbFJ3HxgpaDVZ2XxLZ7bYb94jtd1s45Vwrsv"
    "GkeFx4zsEo23xMYb7cfFdleWKFriksfGRSkeLtod541iQm6uri8m2PxJabZY774jtdzshUPbvGE9ukeL9oULrxpapF7f3O4zJWuJVj4zRhN40sUp23xNyr+/"
    "8W2NxZ7NuEe3X48ZbcDKrmyPavEN3tiY/F+1qEvjDbEp3TfLvdSO21bhDYG4321mt2Q7LeJ7MI8Q2REviK1SLu/ub9Zdpul5aJPiTSzvLm/3V3H+M2u/YpO/"
    "WYCd8Krs77bY2swtrn9OWrmWJp6O1vprJSd8tCLzeTIiLeoER7hdIvV211cWaot8gI/TVo7rdJbhNCAjfIsLu5VeTWF/7o7vdEzRxLVCsbzBiuVU8gDo6Ojx"
    "YQ8ewY7hCiEpYID0LxDAdGAHQdtylUixt5FXNhi6OjFK7RPJc2rDAeLxDnFIdo12ulEhLKCyCHR07pQHjqEpd/cqtlgVeLxYSH4riTLGq0ns5NiuLW23XxQm"
    "4ltPC12f0ZawCDxc0FNVSRxO3uARcQ2tUXCgV3Rx2dVdql3Kwia93QYYZt4u0s8YOL8mWOK/3bT+88vGf+1Ojp2o6OjxdHi6dqPF07UdO1HR074vF0dHR496"
    "PF4sD7lO9HgwnttlxDb3ZwUmXw9arVZbbBYO8ukWdu8Xg8Hi8XiwOxDweDwYS8Hy3gwijxeDCHgwh4vltKWB9wUfHtT7gBYUXxdB2DoHT7l0lKttgGFpR0ej"
    "8/D+toA6dqF0c4+h2n/aWph1PananbV5F1e5Lpcijr38VAe7bteJgVsdvdK3/elrttr8LTpn2u3u4rvxj4hyj2bYry4vNrRYCoypLum2pNzuN0iP3+6vbTbd"
    "qTc7fttpYQlV7bQi53CGBVSyTWAmtS9aVLJLqar/AHbrQ/013Cm57nNutz/qTy/1QQ7Xc7y0CfEa6L8RXKhLPNcrH8/R01o6OjowHR4un3qsF1D0LoGIg6AM"
    "Ogeg7AAuiQ7qcwF6Pmn3ujuunZKqTZebDCdfDg/iwHbR1Ha4FbbaK/opdagvX7lHp2p23w4Ipqw6PxdUWG32YXNsyVf0r8SXMFxsuybQdx2q1uLseJ/EO5WN"
    "xtnhK8uRtk5vpol7ftdsd13u8tr+fcr+5Goe1Sb3PawbHuBf6DjWINs2+2eFoym0cIta42jxtaY2rpaulrUi35dLZkQvGJ0if0D+gf0L+if0T+if0T+jf0bp"
    "G/o39G+h9D6H0PpfS9Ho9KaPR6PR6PpfS9Ho9Ho9Hp9zT+Zo8Xi6f6g8+9GB20/mwaMKYLFC6AMJJZHYdt2OKK9iP9dANb2eBG03V9Ymxtl8y27Dj4b/AHf3"
    "A5/8W2ZQ/RhSM7NapbX7mvejxfiNONhTWjCWA/Fu7CUjeLDcFR3VpZ7xuPiY39r+k7xNlXNfS0yqQYvE26zWnuXiG9XuKJor4nTy8JEr2fpDUrueMHF+TLHF"
    "f7r/AFR5f75aOnan8wGPvU+7R0dHRpaTR5pZlL1LowkPg95/xagro10/SqeO7I5sENvJMNoFNrp2Tx8NkFNHTso4phnTPFcLSi12y/tYrRFymZVlRNjTtR07"
    "UeLo6PxNROz4a4s0S90XexQXcy5FbyPDq4CFBVC7Lb76/MXh68G8XnglNpY+D9k2m72qGGC2SS5fDWyXNz4shtbfd34QQf0EY3ySX7sp6PSsPHR6U0ej0a6c"
    "vR6PR6dtO2nbR6fe0en3/L/fTR07Adh96n3NHR0eJeLo6OnejoXR4l07b0KWu9bsvanYXCryyX/tWS7BKE+F7W0urbZNk12odkjXwxWuJLxc8sFsFbimVNv+"
    "lf0VdbTbe7bNFAuylsrHm/o+1TZ223jm+5XQfu+6pf8Aryl8/dkv366S7PfkylG87Sowz28z8Uo/1gurtNsEXFiJr3xDaRI3TdL/AHgX+z7jt8O3+F7aZG3b"
    "ba3HiTeUwHavBcEytpSbb+nG+f7RvBH/ABj5ZeofibY9w3Dddl8JXu5Q+FrSQeHxBQcoPlDvDx+7J+7/ANUHh/qin+pfN0Y7jvR0dHRgOjo6OjxYTV4jsNWA"
    "Hwd0Ve/4UO//AOJ+I0JW9lR/rVMKbvwGzgL8OkZeEdkR/rTi1EITb3EM42C+t7aTdri+lHuF3K4NusrY3EqIYbO/hXtksu73NrtSLs2U/wCkErJvhbW8+6Ju"
    "fer4P9IFL/S1iGndtpUY57eV7WaC3v8A3nd9xvPCqEbnKhdnabb4h3a1g8NznbbPbbawI3lCB43v7W9tIt32izFqYrzxBukc36H2eW5trO4is7Dxxu86bnw7"
    "4Ir/AEfxfLYhJd4jlweGTl4d8G0Phooq+UHgB3h4/dV+7/mhU/z3l/Mj+bq6ury/1GO9OwdHRgfcAdA6fdAYenYYMIgL5YDToyQXdH/XY0e/0908TrVFbbIo"
    "q2mb/aypGmwD/iPy3aofCOwbhaHaDuUa3cjd5YNp2u0mfhaVESvEN1Y3MG2XG7wJNnfzNO1WUT2gH9FT39hZW23btZ21sd4sFXB3XbF20G4WKr+OWGV8tTn8"
    "SBG7LQFNe1bZIdv2jbFpuvDu1m/u9lFhtm8226QWXh243XC2urz+io3mN/0h2evje/s7u0RYp3Ebba2sni7xBF4dQ9gsZdyv929+2vftr3LdLuz8GfpVOx++"
    "3yH+mLVDi3/Y1O+uLSa18OV/ot4KB/o0yyod4eP3Vfu3AkKuPE6fAPhQWO6/Vb4guPG3hY+FN0qHV1dXFGZpvEW/2H1bWe67ivdtzqO+Q7ZB1fgHxhtHhhN5"
    "PHc3lQ6sl/lq6uo71HbwxtG03P1dDV1A+5UO62vax9VWQ7VHb6sttsN137dN8+rfadwh2LwJ45s/Dl1tGx7vse7fV3v+573u31d+H9z3jcPD+8eJPEs3gTwr"
    "JYXX1Y+KZ/FXh2Twzu/8wHR0707UdGA6MB4h4PF0en3adg0kBnUjtR8GFF6sO8P+ug1fiAfxLxckiw2Mf61bneW1hup3zauRa7tcWGyGC4XsHhFSTutxd2ls"
    "7y+ztNuh3Ja/Dtja3M+7wpig3O/sbY3HjCyQ7jxRu04knuZUFSAzKh5oLqkugeKWieaJi5uUzo8S76ho8X72lxeI7uK2l8Urmu7vxdaXNluXiKwv7Xw5vlna"
    "Kh3yztvD/hHeZb+2ORHjXBSbba7+0tiqcbxufhW5g2rwrtt8q13ew3W23/Zed7r4Kv8A3bYLbxJZERbztEzQUTjxKdnsF7ZZ2Nz4f8IbR7z4f/Ru/QnfN+8T"
    "bJuV14ivNri7Qcfur9h2/wDjP11HTbre5vb/AOuiWPl73tPg3bdm3/c9q8aXu4y+CPq4itIPBn1kWfhTcLHwxu3jrxHsOw3fgTwraeJb+T6wvA1vN498K7Vb"
    "WXg7YvD23eFrHxj4I8Q3F9sdmfFO9X3hD6uzu9z4E8U+H/qk2+w3GTY9jk8QeINz3fwN4Ekk2bwr9Yez/VdZWl34m3PdPBHg/dvEez+HPFPhP6prGxv9yuPE"
    "PgTwhfeNdh8Pbp4Z8AeFIfEt9L9YHgWym8deFtpi2zwiP+YW/Vx4VtPEd8frC8HQ3Pjuy8NwXz+rvwlZ70brx34CuJfFlhabb9Wf1Y7fYXmw/l8V2FjB9XL+"
    "qL/jJPFPgvxTf+I/BOxSeCIN0vP0juf1Z/8AGafWb/xme0f7V/rkI998N29zeeIPrfliO5D7w7U7UYH3A/OjHYdqOjoXR0dO1CxES/dFUwINA6d6d7y4NtbX"
    "o/11Ae/6bb4vngNjss9/Ntni22uIJdv3uyg2uLxNfw2v6Xu/0V4YjslXxtY9utp903826fECrZdjuF3t67vcNzvXWNBNyGbiQtSyp6l4aYl8DUNK0hmR80vn"
    "KDFyX7yGJ0l81DyjL6XiC8AXFdXlszPdKRZ+MPENmbueS8ul3F0tPhndbTarre93ju9zg3tUVn4f8ZWew7JZeP7sGw3Tbd7tl7Fskr8eWdtt24bJJudzc+DF"
    "b4rYje+J4X44vb6S/wB/8R3u8d4OP3Vew7f/ABjxvvPhraBH9ZPgnaxv+/33iPcvrWP/ABGfqxCD40+srmHxp9VJkHjDxelA+sb66R/Hvq//AEfF9XfvH1OP"
    "fvF/hCXwd4O8d2O17XF4U+rXxQd32vcfB+/q8Y+BPFltvP1b7VcbR9S3t/VOEHxZ4rMh8UfU0pf6X8C4D6yvHP8Axl3gL/mnP1ND/XXe0n9M7aD/ALJv6tfc"
    "0eBuf9Tj3bxX4NV4N8H/APNLvBfi5fhO9Wr6pvEs/jbwWrwovR7B0/VK7zI/U39VP/GOj2fGmn1ZP6oT/wASTxR478V2HiHwT4zvvEl/4p2dGx+IPq0H/Ey+"
    "s0f8THaf9q3jPxH4c2S5H1neGtvRuu63m9X33KOjxdHTtTvR07UY7hLxdGEVfLLxo6PEPEPTtq0nWiVAoLwLpTtR0e7gHb923G1g3WG7vbw+ILKZG1+IrfbL"
    "TZ9glvFbZ4vROlRdCWMQ6il3ve63iOgHmgMzKLKiWk0ejo6OKMUlSA8dZKBVS8j2oe9XXtRqiol5Pml81T5ynz3zw+ch5oLql0S8WmRaSjft6he4bpuW6Ktb"
    "+8s17D4vvNhtEfWYqu/+I7LxBfTHb5E9oeP3Vew4P8Z+uSeGZFGoafWjPDL4b2fcrjZd03XYvD/1lI2zbvDv1YWidwm3LxD9cskct59XPiyy2gy/U/Dc3Hjm"
    "Hwhs9j4WsvB/ifw3sv1d23hbcvEPifaN48bbj9Xvh/xOBLtH1aeGvqbkihV4e35XhrxNvXg3ZPHUyZth+q/ZvqruP+Jb42UFeK/A00SPq5+p6WOHc98od722"
    "eD/ZRfVz4rs9jnuPqhtru48aWvg7Y9s8JXMCfqx+ru98MCf/AGTtbj6zPEe27gH9W/iiwskT/VHZR3Hi292m5+rn6rfEdhtdzffVn4Ws7jx/f2O4+BH9U0qI"
    "vEW8fV1tO7bps+zeD/AJ3zdlb5u/1cLRF4v+sZaJvF21ADdvrdlhnv0pYH3h2A7070dPuUY+6GkJfIBFKdsXiGgx4lCXi9Q6l6dgGE1G42UCLfeUCDdJbS5u"
    "3v8At6oNv3/a7Cz2m03GO2sN3v8A9KT8pIa0gPNIBlJdSe5FPuavV2/7uU9pva1eQdXqXR6PpdeyeK/Z/mUJyK04HIh81bMyyxIWJCHzQ8g8kuqe8PH7qvY7"
    "UHfEOj4HzdB2ClJFHSr1PbUF6F0fA0eIeIpiO9BVgqAo8Q6PqxA7UdTTEOjoGEgFlIL5aWEAMBlILAo6MIAYDo6d6Ojo9e1PuUdO1PvBIdNcQGAl1AeTqC9G"
    "HV1evajo6OhYDKCpG52SVJ3Xb7WG4vbzbbJ7ju0l+Li+urlmYU5qmVqrx+7RypoCKfdtv3a+JcoqsgNUNE1dT93zpqr2e54s8e1HAKrufa7hijNHUPiAHypH"
    "ypHChYPLW+Wt8pb5a3y1spVjgp4KeCngp4qeKnip4F4qeKngp4KeCngp4KeKnip4qeJeJeJeJeBdC8TTEuhdC6F0LoXiXQujCS6F0LoXQvEvEvEujoXQvF4l"
    "4uhdHR0dHR0dHR0eLo8XR0eLo8WB3o6OjwdHTvU/cTRiOEgoo6Ojp3owgl4EOaWO0gufF9kgXO/X9yuSdUpKyWSSy/Nye3GK/drrL7Cx2Vx7W37pTPGYkL1a"
    "JFxKkWiQdh3CSR5q9jueyuwq/KD27v2vuk9k8IikPmrfNW4VrL5sj5i3zFvmLfNW1rVhmt5qeSyTs2/CPJTht7y5FlaX25T3MF3Zz5F5Keankp5KeSnkp5Ke"
    "Snkp5KeSnkp1LyLyLJVTIvIurq6l1LqXUupYq6l1L1er1er1ev8AqIMOg7Ah5B8e1GB9wdqPEujo6MDuAXTtQNMGQgtBKpVgpJhRieXm57W3mRP4b8Pzk+CN"
    "ir4g2+Da9026wl3K8vfCm7WV3LGqJcVrcTJoHKOuAMdy8auX92eA9pQePa2/dK0Zc2q8CwgPRLzDqlopljUiFFKJ7K9jueCJLRJlUFFpT0jhB7d37fmdD91P"
    "BAqe0P3lfu+207bsn1cbAPrmvfefGGybP4q8OfU/rt3gzfdw2Dd/E273m9b1aWV7uC7ywvrBcENxcyR7buc1zd2F9YK8HbOvdd9+sPw2Nt3tIKivYN+ii++f"
    "Z/mR/qgB4CnbX7o+7R0YDCQ8Q8Xg8Hi6NKWlIcNraSiW15Toa0YqHDIpxRVAhSlrTmE26SIsQ/HAp4i8G/8AGTbnbC58RXn+M+BUXh2bZfD2+bltkw64GOAd"
    "l9XO23u3+KvDafDl3/ssrhSblJQKaIHUsuiqYJpb+wviXN7fS+l6OtewPUmFCodSxxavZdXq9ae7xJhV7CaVSgqNKCL25klUgByFjdSRPyHE0dGkaI49oeP3"
    "V/u34Xt0XfiP65blat4drvO72Vr9ToHuH1Up/wCJf4n21e8fWP4r8XR+Ax4R8VI8fR+BrBe2fWL408bzeFd52Td0fWP4W+rrxJulhu31keMt42vcvqp22CLZ"
    "7T63PEMG6eL902feN5+8eH3afcH8/o6On84nGlHR0dHR4un3KvJhThhmmZhkSBSI42S1fo62Un9HSB+6TBxwKrDZEiXblFyWdwkotFqabSQGK1AfKkYhkard"
    "DRCgPx8APEng/wD4ya4/4yG6P8Z+rc/61+EF4+G5vag4hgPat82qHa/rFuxd31n+l5YbqpVbckz3qvD8EpXQlZI8N7ha7Vusaytr9ou18P7XP4WvPCHhhNon"
    "wnsa77e9g2iw2bbtou90e57ddbTejipSnN4V3aHaoLZU5S/IJUWapIIe33VtaXF3PHcSqH0ceVcpVPyi9uYkLrrFdbkIU2V0qMpFJNvvYbdHhu5X4eYHT4W8"
    "JWe8K7Qcfur/AHb2a+G2bx9cO2ySKe0+Bje+GPqbNbD6qv8AjMOYiH65fGXjaDw/v0H1sW0C/Bu6fpv6yvrO/wCM0+p0UsPBBH9MPrZP/Eu8M+Lt38LTWf1h"
    "eFfEs/1h+FrTwzuv3vL+eo6feo6OjoHR0eLxeLxeLxdB2AYYSwPugB07UdOySh42pBSA0SyQn3qYuIhTNtAp4TIMV7dRGC9WoKlhITcRofvQYuKtE8ZYUC1C"
    "NJ94CWbpT94VUzqU/G9Tv3hDTxNL/wAZHeClz9Xl3Dbbd4Su1foBftQDXyzKXtidg3K98bcyK726a3ns7zRRLTwXxLCFFx1jQpYJzFbfctp/o1J4l2YWVufD"
    "y908UToR4X8OX9nZL3q/TuW6NXG33bdILCwViXbKX7z4h3TxTsd3zpru6PC3tbm8F7st9t9kf3fEB+UXt2e17ZN4OfgC9ns9/Rvxg8EWd8vb7vfvGEviHbRv"
    "lqPAzHs+DbW1uT2h4/dX+7ZfhH6wrS224Wn1NoX458ex75b/AFZ+JNl2G0+r7drHZfEvirdYrvxcrxL4F8b2Vru/1c+Cbfwjv1taeMfHO6WW8eJ/q28TbJsN"
    "ls+4K2ndfEW6fVp4ntvA/jLaLbarWw+qXZ7rxr4oPindvvHh/qkfzVOwDQlFNHo0pSWI0vlBiFRfu0zwUHq8KvAsg0xDoFPlrS7e/mQPo5WizuXHa7mAiK6D"
    "SAkGa0S/erNm/iAN2H76sM3SlNRq9XNe2lsbPfNruUeM5UTb34YK0+IJrm7l8R7kMbzwHKqGHZ7lSPDq/bgIq18NqQf0l9YlrDbbtB4y8P221z1lAiNQkB4R"
    "h5xJZuUs3Ki+cp84vnPmh81LzQ8439Exyw/o6jdb33LBL5YokLQreNyvN8uBEpJMKiNvvr7a13d7uN8MBjaoswu4XtiXJQKQqi7VVn/R1+BzGjdtx3KSinHD"
    "AdvtoEf0Af5fCi9ztoO0PH7q/wB392jo6feo6Oj1/mDw/mR/qEB0dHR0YDp2HYdhRijFGKMKIepaSQeassqJfF0LjQkswAgQSUUihTAstERLgsw7eCMEJCXL"
    "cyBm6umpS1MV7BhNWIVF8hT5JdzuFlbI2Tc7BO2eMp0T7x4X6/EF1tcI8Q7pGlF/4EUhFtsu52tt4bmClSJComiJRHKTQzGhmSSZmZpKGaQvI/cRESnk0PKD"
    "kGKnV1+9UuqmQsJ50j58jFzIH70p+9F+8PngjnR0Msb5iS8kPKN2t7cWUy/pV4JeCaR7rex7Vy3jp4Z8THw9J2h4/dV+7/1QeH8yP5mjo8SwkvBTxPbVjtUO"
    "odQ6h5OodQ8g8wwsPMPmBiVL5oYnDEzEoLyS4xakQWVst/otYI2xZaNsUD7gae5ID5FsGmK0DT7qlmaF8xL5r5r5j5pfNLyfMU8lHti5+UmHbtx2m3sfGV1b"
    "bhvO0zq23crzxhcz388vPkzjAM7M6mVrU81Op+4ofRn7sH7mV0dx7f8AMoDX+7+7k6kOpPaJGSp0hKvu1LzU+Yt81feHj91X7vttm07lvNwfqq8XiO/2++2u"
    "62PZb/xDe7hYXO13v88iOWU6gs8P9QCj6XUPJhSmlReSX0OsYZUh1ev3df5pES1BhgkMLIcU5SbfcFUQrmteaXmwpL+jfSxRgjsO4dHQujo6Oj3XwTuu5bqn"
    "6s98U932qfZb+QrDUVF40AqWAayUzj1+8f3Z4fct/wDF5GXde3/MeSB0q/d/zNv7V17X3xxPHtBx+6r2GiNcy90vbT6sPDKfrC8YJuJFWn1neD/qkH/Er3jw"
    "ZvviXxXv/gjxB4cighmuprf6qfFc0e97Duvh662jZd0326m+qjxZFDcW9xaTS+AfESEX/gHetuv992DcPDl7smx3/iG9vtmvrDdrX6qvFk8e+eHN48OTfVT4"
    "e3Lb1eL/AAfvO0zbH4b3jxFNe/Vb4rtLdQUhe/8AhbdvDQ2babzfr+72O+sN5/2Xnib9J73sF9sW6Wf1XeK7qPfPDe8eHJGPuU+/UvI/e1YQovlqeJdO2LoR"
    "9yodQ69rW9ktnIpEyqGvmMaIEbt5bVBivbakpVKaMMPV0LowGGA6OjCSwl4h4hoRGWiKOmOJSVPxtr4omGtGtPTCkY4uX95BxPH7h/cnh9y2/cSe07gVW6af"
    "eHBHsq/d/wAzD7V17X3g08V6K7Qcfuq9h+EEpX4p+uJaj4hf1MLV7z9XCUo8e/WH4w3dO/8A1ceIbvxPafVnawWXjjxvD9YcG/eI/Gt/4j2/a4L/AMOfV1t8"
    "31m7fffW9t8QX4s8QXfh7wXJ4k3qe9+tG3h3bYvqvgi2nZdl8VXOz7/f7n9YHiW88VxbjP8AVh9T+4Xtze+Id33W5vtyh3fwx4J8NXP1ibbvH1sbXBZ+IPrl"
    "/d/Vh/xmfic0+tH60fFG57OrwJuSr7xt9YNr47/TO7+Nd43bZgx9zV0dGKPpdA+kPJAYKC8WAH0h0SXi6F1LClPqL6mFyB82Rlay8lMLU6k9g0sMBpqHkT2A"
    "LSksJLESi0G5Q03FzUTVaOSWlNqWURpZUklKVlphUWUKS0QKI5MjEEjFvK/d1PkhhEaWmVKH7ytolUT4yNfEkoqNp2rb5/By/Yh9nzl/eQcVce1nCLi7R4UT"
    "JEr9y7Pkw3d/JZy330dSEh29OUsdRoHa7htFta6ZbhdWFy9Pu0L1on2D+7oXr3Io6F4qeCg8FUi9q79rVqQpA+4j2uCpOPaHj91XsOwvF7fffWRsqvEe0118"
    "I2avAvhD6qJFL8X/AFk7dNt/iz6m9tmFwNv3fxJ4mtfrO8X7RL9Y1pt+8eE1bvuyPqyH1meN1HxpdeNlQfWPr4HfguT+k3gLxhXw19Xn1TbHZblum9/Wj4pX"
    "eeIF7lL9U31M6bvv1U734s8Qb3a+Erf6xvHV1P4yuPFq7r64LeSfa/qosbmfxP4r/wCapfXF/ta2zbdz3W5i+sLxv4dufFHuHivwHwYo6h6OoeTqy6vN1J7h"
    "5FpIaQhTEReDVUPmFiVYPvEhGYfMDUal0dHiwlpQKxJBSImImISWIC0W9XFYlTTtbTtyQxYhizLFiH7gH7nAGiCNJ6HyrVj3cMSoD56nz1vmrL5inmT95Cy/"
    "F+viNeo27eLO08MEVATQDjL+8t+J9rtb8zn2CtwXZr/d1an5ppkspduaQzVLwxM/H74JD1IT7H5Kn7lT9wJNEe1ckhWSmsmlXV16cmg9S+K6d4eP3Vex28Jf"
    "WBufhZJ+tTwqFeK/Ge6eK5vB3iGPwxvFv9Z233c3iL6zhebd4a8R3vhjcJPrF8GXi/GPji98Vnwh43v/AAoUfWJ4KtF+J/E9/wCKr3xN4ztt/wDDz8GeKD4U"
    "3Txr4r/pZuHhXxRd+FtyuPrC8FSS739ZB33w14X8R3fhbdfGHjXYvEe2eEfrAu/DlvF9Y/hDbXvu+33iK/8ADv1oJsNrV9aiItw3fxDHufi3xv4qg8V3uy7x"
    "e7BuCvrG8JbkfFXj258Q2+ro8S8XR0er1+/Rgd9XUupPbV0YdHiGEsJeDCGlAdulGEfuwcU0AaV2amkxqaRGlqlQCZCphZD5heTyeX8xXsGO1HR0ZDgOQ8WG"
    "viFrFATRMSipOtJf3ltxPF+W3UG4W8FtyV05fmpnjFApZlgAVb6Ry8VO44inZMClNcRQewDIox7KfY/Jo0cuujqKd6l6pCPauuLWDTt5OPitq9ntBx+6v2O9"
    "HTtR0+9T+Yo6Ojpp3o6OjxdOwdXqwKvEPF4h4h4h4Avlh8sPlMIowA6JejxdHiXiXQsJYSlgIaQlo5IcabNTTZW5fukCWYrVLtRaKj5UIYFuGg24aLiFDN6K"
    "GVSmFFirFWO4DDo6PF0HajAYDDHbEvAsxF2dvyYvFaaeIqaLc3sW/sBy/vLbir2nqxUHbt59ytFZFAiWzGov3YtCFU5LEWLmyCspCyjJ8oMQpfvHSuLq5Iou"
    "JKQIsn7sX7sXgQMjjyFPkrD5Sw8C+Wp3E3PRy1vBdSlbTUEpC1cghIQyaqoWgkH6MlIQFS0rxR2g4/dPs/6o8vva/d0/mtGC8nUFgpeYY5ZY5b+jYKanEvoD"
    "93Spi3DTBAHHFCGhVH9LT3bN2kKygW6HSFD57zKilElTKYiLwAJvo6C5jUeah5xtMiCRV6vXtXsGAGMXp2GIYUGlTQkKcvjzYYZI/HO0IG+XCdy3TUNVS1RB"
    "QTGhAOIdYK5JqFwg+8RP3pD96D96LM6w/epGbiR8+R82R5rYSpTKCl4qZyS8lPNbzW81F5qZWovmrA50j50j50jzlfPW+ct85T5qmJlPnKLFwoP3gv3gvnks"
    "zEMTh+8B+8IfPS/eEPnofOjZmjL5sfeHj90+z/qjy/1TRhjsGGA6MBhJdGKsNK1PmqpbLKUlZJqSRq0xJpq9SwlhD5SS/d4WIEP3dLFnCzaCnuSw4Y1RJIqE"
    "oIYHYBggsHsQXqHFIoNHg7w7KpHgnw0/F1hbbbu6wplCi5UkJQMgQM7gAS2wBUv22LhKU++ygVKxKnp+7borHKgAFLn4/eA7gMJ6fvji4h1XAor7wHYW8veH"
    "j91Xs/6o8v8AUo7D7gowUtKiHmosU++GKO3XVNWGGGlhhjsC6nuCe471eTCgF5MLdasJyaLSQuFCkhAfjzXxCsPwpa2VzYz+xDwP7y6/e2vtS/vEivbbYIrm"
    "/v8Awfs9sZvYo9u3KbbQ/Kjtv3E7VTK44ji/KmnYcGQ0jROo8/ueTR7RFDF7VxxZglTD2OnZPEPAC07Q8WCUnmrL5imJVgkmn+qPL/Ug+4GHo9GC6sF1YLCm"
    "FPJgsqxP6SsUn9KbaHbbptyY0bhZLUKMMEMFgsFpr2HeryeYfMFcw+YHzQ+Y5bqK3Qre9qF3Fd2s5rR5MSqDjuJK28qiI347/wCMiW9v3Ve228qCpHu67YSD"
    "ruv3tr7cv7xHDy20gbheGJUktcD2LHsU0t/3M51Ueratpj3VYHVvm2I22c8Py9h7NGeHkggM6n7gpg0+0tx+1s+z2W5WPjXYtq2KS4nC/D/ZXFo4uh/RfaHj"
    "2oXipojXksGlC9f9T+X+paurq69h3DBYLqwWZUg++oAG5QuW+tJRPTndtrVDFuVvva/fkbrti0jddsccqJEZMLYWwt5vOj575oYUlT1dTWpdVOqi01LuL+/v"
    "CkNHSfDy1r2ppcHJDgUGhWnjddfES6MlwyJhXvKgoqOt1+8tfbl/eIflEsxr3vd7a73OYdLQKshp9h2/7ucip9rwoIFBBovxXdRXsB4fl7JOjsOQq80x/L4d"
    "XaxbkQSo204jxU+WtiNZctuuEitSapR7Ww3SINs+sa7try5lx/Q1C9WElS1JIKUmtC4pUfovtDx71LBNT/qny/mR/qAMFqWUqBqOwLDD3cfx6jIZdCwk1wha"
    "UJri8HayRoVtah+jwWmnarq6vJ1eVGZQkbvvg21G37/ZX6ea+a0zAMUJDD8Of7SGC0SamzRPewW4QPE0pVvCi1Ko5TVG6LSpS+N3+8t6la/bQqjzD5oe22A3"
    "S8l1QwglJIeQoSXbfu5/bPG0UiIZpd4qx91JLzLCtZapUFF8xThoq5PAexQqFpLgJJJMMlOqmjKNdxdrlQFqrUkp9qwVjZeO45or9S4zZ5KdVNCjnKVBYJqH"
    "CmP3HtDx+4OP+qfL+Zq6h1eTyDyDyDzebzLzL5j5peanmp1L1dSyTWpYWsMTLaZ1MXD94L3I53eIasXpWiaICcsYWEwvCFrQh4Q121ZVYcwB8988vnl88vmq"
    "fOU+cpiUvxQSpVAzwtFfxTJpLSwwX4emx2nnh80Pm1cSz78hRL8RIVDu2bUpqlAEt57ysmqrr95afvJfbTxrpEsxr95m2+7X7JaVkOlX5F2vsT8fO0htriVz"
    "xRoR2T7Vx7aOLPtKLr0W6QYbH/GJP8SaOM5BKkkDzaPaX7XihCAK9HYaKk9sNPCO5UiPtDx+55/6p8v5jyHejxdHR0D0fS9Ho6B0DoHR0a5UxFV0hpuYlH7g"
    "d/T3qqWrFmlapojGv8XqOTX6BqENF8h2FwsWfvEj94lfvEj94kL50jE0jN0pCf0pABBd85O/rUo0dNIpVi2M5Si2mE6Ro8tal7Isja+aWJVO5ilubWLw5vah"
    "+g/FKHvKLq2vukspQ5EAkR4M1BXEZSEGF8la3yFhxxgGSKMiRBUpdSDV0U8VvBbxU4FUEywSVNRJNHi8XiHShWMjQB0Dx1IWWK4pzSmLOJR5hj5T5dHiXe71"
    "Pe2XKL5RYiIK05NWS3gmnKQ+Uh8pLollKWQKd4eP3PM/6p8v5gkUyDCqn71eyFEuv3VrwTMvNbgpzquvcO/P8Zq1KauNQAj2s42hSCc0OWRISqRD22qrLoqu"
    "kSYlRzpweLGNb9ITY002MfxbfAAX5W6awXiB7lZJSiEro+akmqSdmxTtO97gUR7buMsO1WV7Hdx2ZUm0SrXxd/tdL8JQWM2wZKx2y2VuNxYWd1uV1eR3Vlco"
    "lWormkQr3iRoklWormQwtZZuEh+8JfvAD94S/eEPnoZljZKcskvJLyS80vmJfMS+YGVvIPmB818180vml88vnqL5hfML5qnzFPmKfNW+Ytla3zFvmLeanmp5"
    "KdS6n7sPH/VtQ66fzPkULCUKS6gNXGPh2K0gc0uI1UZgk88NC8w811KibY9oP3y7iQqglkUt6P3wA3MnNm62rNqyr1UFSqsjC5nldFqM9D70/fJorUK5S591"
    "vJU7fuiYk5Cl7MYbWqjKpay/LaryO2g3W8ju1HQBVTb7pRN/uKBHaX6ChadAA/MLX+hIg1J6Smhi3S3TYK3m9UvfplzX5L8MyqRs5T9Hs61R3nhqf3Tc97kE"
    "+5xjqn/ePw8vlbx4hv8A3uxO6WsSJAKfdShJSqIPANQA7Gn8ykD+aHFfH7/l5doeP3TT+YtNm3G/TL4X3eJMkS4V/wA59X3hCHxPezeN/q+26fxz4X2QbNV5"
    "PJ1q6urzfMHb6tdos948SeO/rBFbzxR4Zl8CeANtsLvwiDoF0eZeZp2iICle00ylIEpIeRw7Q/vHa+17ytieSvkv2mWRqAqgrl/GQwJngtrRooF8ULFFqfkN"
    "yKUXt/zbdI1UNeA96OInAZWnFCwpoko5ZklEV6LYL1QOJ4++4WMPA+ykgKF+qKALGO6mt2/D5A2lQ+j2v/HNmmEF3uut6j2px9I9nNN03eRC0RLuaS+yWWR0"
    "9kDoXwa+P80PZ/mBx85Pvhj2Rw7Q8f5nb9lvdxEvhG/QiRC4pPDOzJv5N18TR7euHxle53lrZ+IrBQKFbdsl/uTl8IbghMsa4V/zFpum57enavBnibeoPGFx"
    "aeGPBD8A21veeLvEc3g/wVulmjwf9ZNj4H21I8cb9d+CPB28+HfD9p478V3vjrwVtF14u8MbHuGwfVdvm03b8d+J9ku1X227en6pvq1FfBv1feGYPE+8Xvjj"
    "wjtF946tPCgV93R1dQ6h1qk9ovbaF4urBeQa6FXSzR9JfS045YIqlCAaQhycuhEDTTCT94ePlmWqWrCwGZktUqSkJDwNVghCFlL5wapCoq1f978/ORVCiQAK"
    "kNBJU3F3CANytXuMqJrh7Lf2ttZr/d7eul3aSiFd0Y1zD2p/3j224Tabhul4m4I3LBMnslnhJj7t2R7C+Dl/mgx7P3fJ+SePnLxHE07+SWGPZT7OUbyjcS46"
    "5RvJDWtClZIeSHkh5IdQ7C3F3e73un6Ht9t8V3iLnxfZIw2Um38NHi/BcquZuVombxLvW5p2SzsfFd8i48X2aDD/ADCULUNr8beJdltrxdv4/wDAD+rbXxp9"
    "aeZ8Y/VMtafF234D65vrNNfGn1Y1T4Ppp4BBP1d/U5/xkPi3/jKb9X/MGfqxNfB/hHxNceFdzuL76qfFM3jbwMfDCP5kcD2BxJUCah1DqHVnjUMkMkVqloxy"
    "+hqkw1yhayAFL0Self7xTDBSDkGVa5VZVV01SQ1K/ioWQ86kKJJLUr6NR1SdZl1kGoWqrS58SroIvAErdquMWEivo7OdME4JZ9nzn9ss9pQlKlnpL1LPDsj9"
    "2vtL/Nj2fu+T8g/OXiOJ490MB/lT7PaHj908O233ItL3xBtC91g2/wAM381z4wvUiLwlfIktd32S526SKGWZexbcNls/0l/rzv22q3i0sPC+5T3PjC5jRa/z"
    "H1ceLNv2pMv1Q7dcyeJd62Dwt4Wf1aD/AImfjDw5sXjDdtt2nw/9V9n4I3FV39YP1kSol8Y/V14sg8N7je/VLabjd+Lt42Xwt4X+qSeC28QeKylfia/mhP1P"
    "/VpPBF4R8AXHhmLcJ/qeRNP9Y2/7SnaB/NYk9jp2Io9GAS6s1r1M1ZrWtBGrrK01QsVzW18whQuKZUGWS1Bp4lOuKgEgkpTUlGJzSVIFWofQYqDJo8nmS81U"
    "rGypFcgXkGteTTKkCWW3Mka7IndKC5cFxGi3WoYo9p5dOWsqqrL81IxJo1+z91HsL7S/zHl2/L3HF+THsp4+a0lTwNeUXyy+WyKPCjweApU94eP3fLvtW+bh"
    "ZJufE26hE0sk64Z5bWWy8YxFP9LdpjTu+/3W69tq3vcLATeKN0EdxcT3Uv8AMF6gcA/q0IHjP60f+MyNSaMCnYLWkdqPEOg7ZKpR6d9PuIKUvz0p2Gp1SdVN"
    "QAaVrSIeWCa1oWavWvXRGeVbgFPvFf4w1gEKSjGhpwfNq1YB5JDM6jHb3KrSaSUSnoYUhJMqSzMSMywsvi6BhBL5ehgWk4Jph2SAQu3SVixmLkSK8mMtUCaL"
    "TIVFCyShaXmjHSqjU699Go6VejoO6PZX2l+75ni/Lt+XuOL07VaTr5ysant+Vq4+Y4Jr3i4/dONO2yeG41Il8R7NtzT4yslmew2ffYLfwlfyrj8JWCB/RfaV"
    "O68H6bR4WOU/iDZdtaPGloSq32bxDDuO3zbbc/zPl/qFQKT91K8XzFD7gVQ1JdSBkadLyozQnRqozR0q0CisTUJ1xQ14YnlkK0dSHkGmyWqyzZoX0urGLwSA"
    "MKJKS8ElqBBCKkYgbtts21o3KGTb93ltQNs0qtMOARbCMXBFvbyWvvcV1yVXRSqQVeaw+cp88MzpIkUFKaAgjlpJMSnylPlqDqaYqeBdHo6NPsk9l9tKdvMu"
    "jxLxPb8uLxDol9L6X0vpdUuodXqwOrFRGDUCkUNcSXy9cdIojKvtFx+6R0vY7VN5uXivcVwo7eFttMSN08VFCri9u7kpVIk2XiHcrM291Z77Y39nJZXT2++l"
    "2668SwR3m0/f2zar/eLvdvqvutl2Dy+7te2Xe8Xy/Ang83F/Y3W13n834A8M7XuFvulxZ3N+1BIJOnijwSPDmz7Z4C2298F7btd/vFzeWlzYXJL8O+H9h23w"
    "jKrmrfhLwJunitq+qO0uEbpt15tF9ozRmjOLRjkQiqcAfow1EAZVCi9A86BM6kAULu8Cs6PJxpK1SAoWCWkIUlMRQqqUtM4BXLkg3U6wua5nWbicpMpLyero"
    "WQcuprqCH5tXAaGX2u9VPmLD563zWJUvmh5VZZo9Q9a6s9TweLo6urqXkp1U+p07YVfLeDwo6B0S+l1S6jsleIzNeYp8xTzU8i6n70XH7qvYfhP/AGo73sd9"
    "uF7/AEV3RzbHd2lz4juvc9v+5sl4qyv/ABdagv8Aohu9D4S3Z38a7fw19/Y/HP8ARrw2u8ub36o/Lyn+qi8Kbn6osrLZfqvnu7Txj4HvPCZ8L7Rve9bn/ss7"
    "VN74s2Xfdk3Sx+rm83Lw3F9UME0O1/VfuUq/FX1a3OxWG3bde7xep+qiztk+LPA+5+FC/CHhtPindbD6o5pprHwtf7x4g/2V9smPxP4V3LwteeFfBO6+Kz/s"
    "qtqnO87Te7HuH1m/8Ybt/gr33wd9U/h6ML8f+ELKKTwn4N3PxbNJ9Wlrfx7ztF9sV+x9YBg8H/V1Dfr8X/W7JbL8U6ss8ahppkSmqVJrkGrOn8YoT1rwxUtB"
    "CCQ8iR5ZoAiwlUqKNJUmPLC3a5EupKpFqlZSoM1Dop0VXUFEiQAsEmRISuXUrOXWWutR3Xw8puLPH7lNPuVU81sKUT1Pg6pdUuqXVLyS8wHzHzC8y8i/y5Kd"
    "T/MeSvvDv5csvlFxRGvKLweBeGmNWsdFHa3U9nL/AEp3V/0p3Z2u6Xe5bv4t9qjp2o0aK8Vn/W8eLt4ofFu7l3u+blfx/fLh/wCaM+R9n6yb25tvBn1PXUqN"
    "+8c3c934r8QXEl/9Uf1cbfdReCVfVT40VN9Y1hdy+BYL2fb/AKn/AADcTW3i763ru4X4i8E3Elz9W31NWSCNw+rfxxu97Ps26xfVo/qk/wCMq+sLdLy+8VfV"
    "fZ3MPhFf1XeM55vGm23x+rjw9d+ONx23w59WPimx3f65UBO//Wgr/iIbCf8AmEH1Q/8AGW+O0j+mFvsG6f7Lm1+q3xpYXX1y23S4o1TSb1uG2fVfsdxcXN7c"
    "0UyFPWuoYJyJNUqVX6UtSFEYOkJTOlKe1aMKNemlY3EElVY65R1jXHWanMq6ilS8qMKSGCZV3FnNbXJBDhUlKx7ioLQmh51a3YauA7q4eU3Fnj9xXs9x28u0"
    "nD+aHs/zIa/veXby7Rcfuq9j7llP7pdqm2Lco/cvDL9y8Mv3Lwy/cvDT9z8NB+Jt0triOn83CP8AmDDPs/Wif+Ij9URP9KvGP/GVbgP+YNfV8o734HnR4ztr"
    "vxP4b8R7Htc3/NF/BB/4l31t/wDGW+AP+ad/U7uEIl3i08YbJfXfhnxfD4bf1Sf8ZX41/wCMt+rWdO7eFLmLxhZXm/8AhrxPtHh+Bd5tf1V+FbbxL4l8Q/XO"
    "P9e/rRH/ABENg/5pD9Ui0o8XfWFtu4QeKYvfPFn1Z26PF15d+M/Du8eHi7CK8nvvHdl4strzvV5F5l80hmZTzUXr2qQVqr2PbIvMpcS6qHMKlplCgZCqTqX0"
    "gdLABCbAEGyhQFFNUonUhXOSKl5FxLSlpmjEkUKrkXAxkerq1cPyzcWeP3F+z3Hb8ppRr9nskVP3vL8v8z5L+95fdi4/dVTD7ur1er1er1/nU+M91R4bZ4b5"
    "4y3XxBtvh/f77w1fbjezbpfS+MN0m8N7Xu9/st4j64vEAh3rf918Q3avGW6q8M7XuE207l4h3+98S3+0eMNz2XZ7a5ns57b64PEcMXiLxhvXicvw/wCIL3w1"
    "fbnfz7rf7fuN9tN3H9cHiFMO++JN38Rz+GfHW9+GEbp9a3iDcB4l8UX3iq53zxluniDb7Xxhudn4csL672283j6zN53vaNh8Sbv4auJPrj8Qqh3LdL/eLx2V"
    "7c7deeJ/GO7eKz36GcWQGcHRL0ejo1KocqPN5s6nsjAGtDVKjkhJilWJJJZEj3h8yFRjlNVSrUamgJLhuLiJKTR6HsStJRcTIcqlSr5Sg6KD1Z4flm4ulTQh"
    "0PeT2e6e35fKjk9ntH7Z4p4/c8vy/wA0r7x4/di4/dV7H+qPL/UdPu6d6fep3xeLoH0vpfS9HV5tdD90AntTsQA9HlosyBOagQoZREczz8g0NKSpRhUkrRDD"
    "cKlyK1kGf2gtQYkkL5qC8Y1PkaTe0/OpeRDq9HJ7NHQjsO35PLyl9nsj21+0Pu+X5f5nyX90cfvQ8fuq9n/fhV1eTr97i9Q8no8CXy3gl/Rh5ID5gfNfNLK6"
    "uSlPuV+5lRl6PrwK1g5AqiI5nn5BocMFwHIsl0JZiNRazET+0njFLhD1JBWgjQBWR7efctdMXq6sUfS/y+XlL7NHQNOIUqhPl938v80p0Lo6PTtp96Lj91Xs"
    "f6o8v9WEEd/Ptq9Gj2TVycPvScWmIqC0JSP5hBdVh5nJAPM8/INNH7xI9CysIfvagReLrKsrUOLPsaU0wX0vRpUA/oi+UGY1Blq4dumnb8nl5L9n7n5fup4U"
    "LxdHiHQPpfS+l1HYvX7nl3HHtFx+6fY/33aPTtp34MnIffR7Jq1/eDUBWjSaBSqn+YDSEkVQHzMVQWE9zbUsvcH5UNMgGtSVNaSDRngOx9j8v5ZuyuPaqwBK"
    "zy3yklmJQZBDq9GPY8nJ7Lo6F4KeBpyy8KPFL6HUMHSv8z5HtR0en8xFx+6fY/3+p9kjVXCgej0dUvIPNlZrmWSSPv0LoWiFa2IZAxEoNMQoSa5EMKSSKVMU"
    "aWaVRBmuSO0QcrZLuFpWsdj7PlrjMxxPHtl0tQ6ci+ap80PoL5aXqkZ6ZPR5pfMfMLzUyoup+6OGnav3/I/zkXH7p9n/AH31/m69qvOgzZWT2CSocp9DStIO"
    "lemtY3IpJTo9Ho9Ho6upevaJWBE6C8ULAUQnzuF3FyrlMW6nqkw+55LXs9Fm0MnKQpyjFQ7eXl+WXU+Z4/cl9n7tVUNaULXw/mhw/mfI/wA5Fx+6r2HbwLuZ"
    "E7HBT9B2r/Qdq/0Hav8AQdq/0Fav9BWj/QNo/wBAWj/QFo/6P2j/AKPWb/o7Zv8Ao5Zv+jdm/wCjVk/6M2T/AKMWT/ovZP8AotYv+iti/wCilg/6J2D/AKJb"
    "e/6I7e/6Iba/6H7a/wCh+2v+h22v+hu2v+hu2P8AoZtj/oZtj/oZtj3rw4dqPIS+Ql8hL5CXyEvkJfIS+Ql8hL93S+Ql+7pfu6X7ul+7pfu6X7ul+7pfu6X7"
    "ul+7ofu6X7ul+7pfu6X7ql+6pclriEqoBg+itUOG6sk22ro1OhL5a2IZC+RI+XR0SwEFhNCQssKkD5qHMvrTdKDStCyQp5ANfJUwpCCLjFmVJZ5lM5Goksdj"
    "w8vKQEFnj9xfsfdT7K/Za+H815ffIoX5H+Y0enYYuIIrRL0awArR0SzjTpeyFIuf9Qj74/mvFKkDaf8AUmnarq6l69uA1eCnyVU5RdEVoioArQscwNJkJxJf"
    "JjfLQ1USkTLDRKormUvNEsePOgcuKl4NNUFalreABUhABwDCwkKlSXm+at6rSmjxDIUB5Jo5OyuP3F+x91Hsr9nRr4fzXl/NcXifua9ul9Lx7xfzCFqiUnfr"
    "kD9Pyv8ApBK/6Qyv+kMr/pDM/wCkUz/pHM/6STP+ksz/AKSzP+k07/pPO/6UTv8ApTO/6VXL/pXdP+lly/6W3T/pdcv+l90/6YXb/pjdv+mV4/6Z3j/pneP+"
    "mt4/6bXj/pveP+m96/6cXr/pxev+nF6/6cXr3Der3cl89T56n7wXzy+ep+8KfvCn7wp+8KfvCn7wp89T56n7wp+8qfvSn7yp+8qfvSn7yt+8LfPW+ct88vnv"
    "3l+8qap1rdFlrqlNSXEUFy4B81VUSqK9A8ow84wozxh+8B84srWoYFpTU0QHRNckB8xIfNo+aSypSnRTxLoHwYVGGCQSpRYJxBYxLx0/KEpovtkQ9CcQygjs"
    "v2Pup9hfBr4f6lHbR0DI76PR0HeLj/M0dHTtR0dHR07UdHR0dO9O9HR0dO9O9P57R0evajo9Hp2oXRTxLwaYgom3CXFCZVrIQrmh6kpqBLQhjiutfPA1KCHi"
    "lj2isvMlpWQ+etmVZdSXQ9tHp2qWaUaeLXoWlWJlkjkVoUsUdFBhRoKNf3D2zLJyFD91HsK4NfD71C6fcBYoX5lPYDtQ/eqXV6updXUPTvFx/wBXeX+qAA8Q"
    "6JYwD6a5JdQ9HRJaqB5JZUQxk1VA6aaBrxpxBxdQ1LyZOuai6n7mn3NMe1SD92PGrV7VRSTDJo9nzYqwqg6S1Cn3D9wKLyBdEl0I7I9hXBr4OhLxenapde2r"
    "o6Ogej86n7nk9XRReL6X0uryL1dPuxcf9XeX+p6uutXV6vWtD2xDxSwUh16ipLBxZWS6l1J+95900r34dvLtx+8NC1ce2rAIHme1TTpoXq9GeNC9Xp90VdWK"
    "UPBkvIvUvV4uneuvHuKPR0S08S9H0vQPIvJToXiXi6fd0eneLj/q7y/1Jq9XQOj86h1o8nXXMvIup/njx+5R6fc8u/H7mr1dHpWqHm6rLFQ/PidGKUUkU89C"
    "KPzyLq9KUNPug9zSlO3mWO9CXiXi+l9Lq9WGEOjo+ByFO2jr2qHX78XH/V3l/qWrqf8AU/n38vv0PbR6dqF6ujp2HKU8Uhijl4sVrmX0kKSaefF8Gfao6Ojq"
    "Q8gXRJfB0PYdzwdXV1dXV1L1LxLo9H0uode+ag8nVPain5VdT/NRcf8AV3l/vtPH7uLo9O2rOXejEZfK0oh0dHiSzE+SWM0tJq5eLy1yBeGhBoeJ0YZ4upDq"
    "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"
    "X/gP/wADUCh/je3/AGN/55N9t9FTVa+b/wDFLR/2X/j/APoZ/wDkv/4z/kvNP/wxYv8A/9oADAMBAAIRAxEAABAAAAAAAAAAAAAAAABQAAAAQAAAAAAAAAAA"
    "AAAAAAYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABBCECADCCDDCDCACCAIAAAAAAAAAAAAAAAAAAADIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAB"
    "AAwBAAFDABDDBCCCBAACCQAAAAQAAAAAAAAAAAAAAAABwAAAAAAAAAAAAAAAAAAAAAAAAQACAAAAADCAAQEACgAAAAAAAAAAAAABAAAABAAAAAAAAAAAAAAA"
    "AAAYAAAAAAAAAAAAAAAAAAAAAAAAAAQARAAAAgwgCAAAgAAgAAAAAAAAAABEAAABAAAAAAAAAAAAAAAAACwAAAAAAAAAAAAAAAAAAAAAAAAAAgQhFAEQwQwB"
    "AAAEAAAAAAAAAAAAAAIAAAAAAAAAAAAAAAAAAAAACgAAAAAAAAAAAAAAABAAAAjAAwCDAAAACMEACgAgAAAEGAEQUgMgMo4AAAAYIBIAAAAAAAAAAAAAAAAD"
    "AgggAwAAAAAAAAAAAAAAQkiiAAAAAEMBcMAEMEAIAIBAAACCAYAAAgAIZAAAAIAAAAAAAAAAAAAAAACgAAAAAAAAAAAAAAAAAEAAAhAAAMEAAAAMMCQAIEEI"
    "AAAAAIAAIIAAAAIIAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAggQQgBeuABggQgAFAggAABgAAAAAAAIAAAAYAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAACAAMFMMAEAEMEAuQEBAWMBUELACAACQgAgAAAAAIAAAAAAQAAAAAAAAAAAAAAAAAAAAACBAAAAAAAAAAAAIwlQO6QEYMMMMMMkAMEeCIMFENIAQAA"
    "AIIIIAAAAAABQAAABQAAAoAAAAAAAAAAAACAAAAAAAAAAAAAcAIQgEAAEUMMUIIAMEMgMGOBRMMAIAijcmoI4IQIIIZooYItI4IDcYsI4IAAAAAAACAAAAAA"
    "AAABAAAAQzcICARBCBhCgSCCAMMYIRADBACIQAD8KIHAS5KtMDonQbkoeLsXqEe5gEAAAAAABAAAAAAAwADCCAQSAADAhQigAQQBAQoQMcMEQhgIAAEAABMi"
    "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"
    "CxYsWKCzfa/5ixYsFhYX/9k="
)


if __name__ == "__main__" and "--selftest" in sys.argv:
    sys.exit(run_selftests())
else:
    main()
