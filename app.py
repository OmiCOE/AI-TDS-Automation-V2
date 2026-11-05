"""
TDS Automation - web app (Streamlit)
=====================================
Give it a BOQ (no model numbers needed), a cover-page template (no placeholders needed)
and your TDS repository. For every BOQ line it
  1. picks the best TDS from the repository (reads the datasheets; optional AI judge),
  2. reads the model no. / make off that datasheet,
  3. reads your cover page, works out which cell is BOQ ref / Description / Make / Model and fills it,
  4. puts the cover page in front of the TDS and highlights the BOQ description + model in it,
and returns ONE zip. You review/adjust every match before the package is built.

Run locally:   streamlit run app.py          (see ONLINE_SETUP.md)
"""

import io
import os
import platform
import shutil
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="TDS Automation", page_icon="📄", layout="wide")

try:
    import cover_autofill
    import llm_judge
    import sharepoint_source as sp
    import tds_automation as tool
    import tds_matcher as matcher
except (ImportError, SystemExit) as _e:       # a file wasn't uploaded to the repo, or a package failed to install
    st.error("The app could not start.\n\n**Details:** " + str(_e) +
             "\n\nCheck that ALL of these files are in the repository root: app.py, tds_automation.py, "
             "tds_matcher.py, cover_autofill.py, llm_judge.py, sharepoint_source.py, requirements.txt - "
             "and open *Manage app → Logs* for the install log.")
    st.stop()

# Streamlit Community Cloud mounts the repo under /mount/src; there is no 'folder on this computer' there.
IS_CLOUD = Path("/mount/src").exists() or os.environ.get("TDS_HOSTED") == "1"
NONE_LABEL = "— none (cover page only) —"


# --------------------------------------------------------------------------
# helpers (no Streamlit calls -> unit-testable)
# --------------------------------------------------------------------------

def secret(name, default=""):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


def save_upload(data, suffix):
    d = Path(tempfile.mkdtemp(prefix="tds_in_"))
    p = d / ("input" + suffix)
    p.write_bytes(data)
    return p


def zip_to_dir(zip_files, dest=None):
    """Extract one or several uploaded .zip files (bytes or uploader file objects) into ONE folder,
    without making extra copies in memory."""
    tmp = Path(dest) if dest else Path(tempfile.mkdtemp(prefix="tds_db_"))
    if not isinstance(zip_files, (list, tuple)):
        zip_files = [zip_files]
    for zf in zip_files:
        src = io.BytesIO(zf) if isinstance(zf, (bytes, bytearray)) else zf
        try:
            with zipfile.ZipFile(src) as z:
                z.extractall(tmp)
        except zipfile.BadZipFile:
            raise ValueError(f"'{getattr(zf, 'name', 'that file')}' is not a valid .zip. In SharePoint use Download "
                             "on the folder, or zip the folder on your PC, then upload the .zip.")
    return tmp


def read_repository(root, progress=None):
    lib = matcher.TDSLibrary.from_dir(root, progress)
    if lib.count == 0:
        raise ValueError("No PDF files were found in the repository. Check the zip/folder contains the TDS PDFs.")
    return lib


def make_zip_parts(out_dir, dest_dir, max_mb=150):
    """Zip the output in parts of about max_mb so a big BOQ never needs one giant in-memory download."""
    out_dir, dest_dir = Path(out_dir), Path(dest_dir)
    item_dirs = sorted([d for d in out_dir.iterdir() if d.is_dir()], key=lambda d: (len(d.name), d.name))
    logs = [f for f in out_dir.iterdir() if f.is_file()]
    parts, cur, cur_size = [], [], 0
    for d in item_dirs:
        sz = sum(f.stat().st_size for f in d.rglob("*") if f.is_file())
        if cur and cur_size + sz > max_mb * 1e6:
            parts.append(cur)
            cur, cur_size = [], 0
        cur.append(d)
        cur_size += sz
    if cur or not parts:
        parts.append(cur)
    paths = []
    for i, ds in enumerate(parts, 1):
        name = "TDS_Output.zip" if len(parts) == 1 else f"TDS_Output_part{i}_of_{len(parts)}.zip"
        zp = dest_dir / name
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED, compresslevel=1) as z:
            for d in ds:
                for f in d.rglob("*"):
                    if f.is_file():
                        z.write(f, f.relative_to(out_dir))
            for f in logs:
                z.write(f, f.name)
        paths.append(zp)
    return paths


def plan_to_frame(plan, library):
    labels = matcher.option_labels(library.docs)
    rows = []
    for e in plan:
        sugg = e["match"]
        rows.append({
            "Use": bool(e["use"]), "S.No": e["sno"], "Description": e["description"],
            "Match": sugg or NONE_LABEL, "Confidence": e["confidence"], "Decided by": e["by"],
            "Model no.": e["model"], "Make": e["make"],
            "Why / other candidates": e["reason"] + ("  |  also: " + " · ".join(c[0] for c in e["candidates"][1:4]) if e["candidates"] else ""),
        })
    return pd.DataFrame(rows), sorted(labels.values())


def apply_edits(plan, edited, library):
    """Turn the user's edited review table back into a plan that run() can build from."""
    label_to_doc = {lab: d for d, lab in matcher.option_labels(library.docs).items()}
    by_sno = {r["S.No"]: r for r in edited.to_dict("records")}
    final = []
    for e in plan:
        r = by_sno[e["sno"]]
        e2 = dict(e)
        chosen = r["Match"] if r["Match"] != NONE_LABEL else ""
        doc = label_to_doc.get(chosen)
        e2["path"] = doc.path if doc is not None else None
        e2["use"] = bool(r["Use"]) and doc is not None
        e2["match"] = chosen
        user_model, user_make = str(r["Model no."] or "").strip(), str(r["Make"] or "").strip()
        if chosen != e["match"] and doc is not None:            # user switched the datasheet
            if user_model == e["model"]:
                user_model = matcher.infer_model(doc.name)
            if user_make == e["make"]:
                user_make = matcher.infer_make(doc.folder) or user_make
            e2["by"] = "you (manual)"
            e2["confidence"] = "Manual"
        e2["model"], e2["make"] = user_model, user_make
        final.append(e2)
    return final


def build_package(boq_path, cover_path, plan, highlight, progress_cb=None):
    work = Path(tempfile.mkdtemp(prefix="tds_job_"))
    out_dir = work / "Output"
    summary = tool.run(str(boq_path), str(cover_path), None, str(out_dir),
                       progress_cb=progress_cb, highlight=highlight, plan=plan)
    zips = make_zip_parts(out_dir, work)
    return zips, summary, out_dir


def clear_repo_cache():
    for k in ("sp_prov", "library", "repo_key", "plan", "plan_sig", "result"):
        st.session_state.pop(k, None)


# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------

st.title("📄 TDS Automation")
st.caption("BOQ (description only) + your cover page + TDS repository → a cover page and the right, "
           "highlighted TDS for every line, zipped.")

left, right = st.columns(2)

# ---- 1. repository ----------------------------------------------------------------
with left:
    st.header("1 · TDS repository")
    _opts = ["Upload a .zip of the repository", "SharePoint folder (needs Microsoft sign-in)"]
    if not IS_CLOUD:
        _opts.insert(0, "Folder on this computer (OneDrive-synced)")
    source = st.radio("Source", _opts)
    if source.startswith("SharePoint"):
        client_id = st.text_input("App (client) ID", value=secret("CLIENT_ID"))
        tenant_id = st.text_input("Tenant ID or domain", value=secret("TENANT_ID", "organizations"))
        folder_url = st.text_input("SharePoint TDS folder link", value=secret("TDS_FOLDER_URL"))
        if "token" not in st.session_state:
            if st.button("Sign in with Microsoft", disabled=not client_id):
                try:
                    msal_app, flow = sp.start_device_login(client_id, tenant_id or "organizations")
                    st.info(f"1) Open **{flow['verification_uri']}**  \n2) Enter code **{flow['user_code']}**")
                    with st.spinner("Waiting for you to finish signing in..."):
                        st.session_state["token"] = sp.finish_device_login(msal_app, flow)
                    st.rerun()
                except Exception as e:
                    st.error(str(e))
        else:
            st.success("Signed in to Microsoft ✔")
    elif source.startswith("Folder"):
        local_dir = st.text_input("Path of the TDS folder on this computer",
                                  placeholder=r"C:\Users\you\OneDrive - Company\TDS Database")
        st.caption("Run the app on your own PC and sync the SharePoint folder with OneDrive (no admin rights needed).")
    else:
        db_zip = st.file_uploader("Repository (.zip of PDFs, subfolders OK)", type=["zip"], accept_multiple_files=True)
        st.caption("In SharePoint: open the TDS folder → Download (arrives as a .zip). A very large repository "
                   "can be uploaded as several .zip files at once.")

# ---- 2. inputs ------------------------------------------------------------------------
with right:
    st.header("2 · Your files")
    boq_file = st.file_uploader("BOQ (.xlsx) - Description (and Make if you have it); no Model no. needed", type=["xlsx"])
    cover_file = st.file_uploader("Cover page template (.docx) - no placeholders needed", type=["docx"])
    if cover_file is not None:
        info = cover_autofill.analyze_template(save_upload(cover_file.getvalue(), ".docx"))
        with st.expander("How I read your cover page", expanded=True):
            if info["recognised"]:
                for label, field in info["recognised"]:
                    st.write(f"✅ **{label}** → fills with the BOQ's **{field}**")
            else:
                st.warning("I couldn't find labels like 'Description', 'Make', 'Model' or 'BOQ Item No.' "
                           "- add them (or use {Description}, {Make}, {Model no.}, {S.No} placeholders).")
            if info["other"]:
                st.write("➖ Left exactly as you typed it: " + ", ".join(f"*{o}*" for o in info["other"]))
    highlight = st.checkbox("Highlight the BOQ description / model in the TDS", value=True)

    with st.expander("AI judge (optional) - makes matching much more reliable"):
        st.caption("Without it, the app ranks datasheets by their content and you confirm the doubtful lines. "
                   "With it, an AI model reads the BOQ line and the shortlisted datasheets (short excerpts only), "
                   "picks the right one - or says none fits - and reads the model/make off the datasheet. "
                   "Needs your own API key; usage is billed/limited by that provider.")
        prov = st.selectbox("Provider", ["None (free ranking only)", "Anthropic", "OpenAI-compatible (other services)"],
                            index={"anthropic": 1, "openai": 2}.get(secret("LLM_PROVIDER", ""), 0))
        api_key = st.text_input("API key", type="password", value=secret("LLM_API_KEY"))
        model_name = st.text_input("Model", value=secret("LLM_MODEL", "claude-haiku-5-5" if prov == "Anthropic" else ""))
        base_url = st.text_input("Base URL (OpenAI-compatible only)", value=secret("LLM_BASE_URL"))
    min_conf = st.select_slider("Pre-tick 'Use' for matches with confidence at least", ["High", "Medium", "Low"], value="Medium")

judge = None
if not prov.startswith("None") and api_key and model_name:
    judge = llm_judge.Judge("anthropic" if prov == "Anthropic" else "openai", api_key, model_name, base_url=base_url)

# ---- 3. analyze --------------------------------------------------------------------------
st.header("3 · Analyze")
ready = bool(boq_file and cover_file)
if source.startswith("SharePoint"):
    ready = ready and "token" in st.session_state and bool(folder_url)
elif source.startswith("Folder"):
    ready = ready and bool(local_dir) and Path(local_dir.strip().strip('"')).is_dir()
else:
    ready = ready and bool(db_zip)

b1, b2 = st.columns([1, 5])
if b1.button("Re-read repository"):
    clear_repo_cache()
if b2.button("Analyze BOQ & pick TDS files", type="primary", disabled=not ready):
    try:
        status, bar = st.empty(), st.progress(0.0)
        # repository -> library of readable TDS files
        if source.startswith("SharePoint"):
            key = (folder_url, st.session_state["token"][:20])
            if st.session_state.get("repo_key") != key:
                status.write("Indexing the SharePoint folder...")
                spp = sp.SharePointTDSProvider(st.session_state["token"], folder_url, progress_cb=status.write)
                lib = spp.build_library(status.write)
                if lib.count == 0:
                    raise ValueError("No PDF files found in that SharePoint folder (or no access).")
                st.session_state["library"] = lib
                st.session_state["repo_key"] = key
        elif source.startswith("Folder"):
            root = Path(local_dir.strip().strip('"'))
            key = ("dir", str(root))
            if st.session_state.get("repo_key") != key:
                status.write("Reading the TDS folder...")
                st.session_state["library"] = read_repository(root, status.write)
                st.session_state["repo_key"] = key
        else:
            key = ("zip", tuple((f.name, f.size) for f in db_zip))
            if st.session_state.get("repo_key") != key:
                status.write("Reading the uploaded repository...")
                st.session_state["library"] = read_repository(zip_to_dir(db_zip), status.write)
                st.session_state["repo_key"] = key
        library = st.session_state["library"]
        status.write(f"{library.count} TDS files read. Matching the BOQ...")

        boq_path = save_upload(boq_file.getvalue(), ".xlsx")
        cover_path = save_upload(cover_file.getvalue(), ".docx")
        _, items = tool.read_boq(boq_path)

        def on_progress(done, total, msg):
            bar.progress(min(done / max(total, 1), 1.0))
            status.write(msg)

        st.session_state["plan"] = library.plan(items, judge=judge, min_conf=min_conf, progress_cb=on_progress)
        st.session_state["paths"] = (boq_path, cover_path)
        st.session_state.pop("result", None)
        bar.progress(1.0)
        status.write("Done - review the matches below.")
    except Exception as e:
        st.error(f"Something went wrong: {e}")

# ---- 4. review + generate ----------------------------------------------------------------
plan = st.session_state.get("plan")
if plan:
    library = st.session_state["library"]
    st.header("4 · Review & generate")
    n_use = sum(1 for e in plan if e["use"])
    st.caption(f"{n_use} of {len(plan)} lines have a TDS selected. Check the **Match**, tick **Use** for lines you agree with, "
               "and fix **Model no.** / **Make** if needed. Lines without 'Use' get a cover page only.")
    df, options = plan_to_frame(plan, library)
    edited = st.data_editor(
        df, hide_index=True, width="stretch", key="review",
        column_config={
            "Use": st.column_config.CheckboxColumn(width="small"),
            "Match": st.column_config.SelectboxColumn(options=[NONE_LABEL] + options, width="large"),
            "Description": st.column_config.TextColumn(width="medium"),
        },
        disabled=["S.No", "Description", "Confidence", "Decided by", "Why / other candidates"])

    if st.button("Generate package", type="primary"):
        try:
            status, bar = st.empty(), st.progress(0.0)
            final = apply_edits(plan, edited, library)
            boq_path, cover_path = st.session_state["paths"]

            def on_progress(done, total, msg):
                bar.progress(min(done / max(total, 1), 1.0))
                status.write(msg)

            zip_path, summary, out_dir = build_package(boq_path, cover_path, final, highlight, on_progress)
            bar.progress(1.0)
            status.write("Done.")
            st.session_state["result"] = {"zips": zip_path, "summary": summary}
        except Exception as e:
            st.error(f"Something went wrong: {e}")

res = st.session_state.get("result")
if res:
    s = res["summary"]
    st.success(f"{s['items']} lines processed · {s['items'] - s['missing']} with a TDS · {s['missing']} cover page only")
    zips = res["zips"]
    if len(zips) > 1:
        st.info(f"The output is large, so it is split into {len(zips)} zip files (each has its own folders; "
                "the logs are repeated in every part). Download all parts.")
    for i, zp in enumerate(zips, 1):
        with open(zp, "rb") as f:
            st.download_button(f"⬇️ Download {Path(zp).name}  ({Path(zp).stat().st_size / 1e6:,.0f} MB)", f,
                               file_name=Path(zp).name, mime="application/zip", type="primary", key=f"dl{i}")

# ---- diagnostics (helps when something fails on the hosted server) -------------------------------------------------
with st.expander("Diagnostics"):
    import importlib.metadata as md
    def _v(n):
        try:
            return md.version(n)
        except Exception:
            return "not installed"
    free_gb = shutil.disk_usage(tempfile.gettempdir()).free / 1e9
    st.code(f"python {platform.python_version()} | streamlit {st.__version__} | pymupdf {_v('pymupdf')} | "
            f"python-docx {_v('python-docx')} | reportlab {_v('reportlab')}\n"
            f"hosted(Streamlit Cloud): {IS_CLOUD} | free temp disk: {free_gb:,.1f} GB")
