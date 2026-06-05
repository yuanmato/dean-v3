"""
Ross Area & Faculty Toolkit — Streamlit app (v4)
================================================

Reversed mapping: Stage 1 lists each school's OWN departments and maps each to Ross area(s).
Each department gets a "Check faculty individually" toggle (ON by default when it maps to 2+
Ross areas; OFF for a unique match; you can flip either). Stage 2 honours the toggle: flagged
departments are classified per-faculty by journals; unique ones inherit their single area.

Deploy (Streamlit Community Cloud): put streamlit_app.py + ross_toolkit.py + requirements.txt
in a repo; add the secret OPENAI_API_KEY in the app dashboard; point the app at streamlit_app.py.
"""

import os, io, concurrent.futures as cf
import pandas as pd
import streamlit as st
from openai import OpenAI

from ross_toolkit import (
    MATCH_SYSTEM, VERIFY_SYSTEM, COVERAGE_SYSTEM,
    ROSS_AREAS, ROSS_SHORT, BUILTIN, PRICE, WEB_SEARCH_PRICE_PER_1K,
    DEFAULT_MODEL, FLAGSHIP, CHEAP_MODEL, CostMeter, call_model_json,
    parse_departments, default_force_check, ross_coverage, merge_coverage,
    normalize_area, normalize_to_kind, normalize_subfield, valid_ross,
    is_no_match, LADDER, fetch_department_faculty,
)

# --------------------------------------------------------------------------- #
#  Palette                                                                     #
# --------------------------------------------------------------------------- #
NAVY, MAIZE, GREEN, AMBER, GREY, VIOLET, RED = (
    "#00274C", "#FFCB05", "#1f7a4d", "#b07d00", "#9aa3ad", "#6b3fa0", "#b23b3b")
CONF_BG = {"high": "#eef6ef", "medium": "#fbf3df", "low": "#f4f1f7"}
RANK_COLOR = {"Professor": "#00274C", "Associate Professor": "#1f7a8c", "Assistant Professor": "#5a7a2e"}
RANK_ORDER = {"Professor": 0, "Associate Professor": 1, "Assistant Professor": 2}

st.set_page_config(page_title="Ross Area & Faculty Toolkit", page_icon="🎓", layout="wide")
st.markdown("""
<style>
  .stApp { background:#FAF7EF; }
  h1,h2,h3 { font-family:Georgia,serif !important; color:#00274C; }
  .rt-banner { background:#00274C; border-bottom:4px solid #FFCB05; border-radius:12px; padding:20px 24px; margin-bottom:8px; }
  .rt-banner .t { font-family:Georgia,serif; font-size:26px; font-weight:700; color:#fff; line-height:1.15; }
  .rt-banner .s { color:#c9d6e5; font-size:14px; margin-top:4px; }
  .rt-pill { display:inline-block; background:#00274C; color:#FFCB05; font-weight:700; font-size:11px; border-radius:5px; padding:2px 6px; margin-right:4px; }
  .rt-area { color:#00274C; font-weight:600; }
  .rt-dot { display:inline-block; width:8px; height:8px; border-radius:50%; margin-right:6px; vertical-align:middle; }
  .stButton>button { font-family:Georgia,serif; font-weight:600; border-radius:8px; }
  div[data-testid="stMetricValue"] { color:#00274C; }
</style>
""", unsafe_allow_html=True)

# --------------------------------------------------------------------------- #
#  Session state                                                              #
# --------------------------------------------------------------------------- #
ss = st.session_state
ss.setdefault("page", "match")
ss.setdefault("match_schools", [])             # [{ranking,university,school}]
ss.setdefault("match_results", {})             # (uni,school) -> {"status","deps","error"}
ss.setdefault("carry", None)                   # [{ranking,university,school,departments:[...]}]
ss.setdefault("fac_schools", [])               # same shape as carry
ss.setdefault("fac_results", {})               # (uni,school,dept) -> {"status","data","error"}
ss.setdefault("meter", CostMeter(DEFAULT_MODEL))

# --------------------------------------------------------------------------- #
#  Helpers                                                                    #
# --------------------------------------------------------------------------- #
def get_api_key():
    key = os.environ.get("OPENAI_API_KEY", "")
    try:
        if "OPENAI_API_KEY" in st.secrets:
            key = st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass
    return key

def get_client():
    k = get_api_key()
    return OpenAI(api_key=k) if k else None

def read_uploaded(uploaded):
    data = uploaded.getvalue()
    if (uploaded.name or "").lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(io.BytesIO(data))
    return pd.read_csv(io.BytesIO(data))

def pick(cols, *pats):
    import re
    for p in pats:
        for c in cols:
            if re.search(p, str(c), re.I):
                return c
    return None

def df_to_schools(df):
    cols = list(df.columns)
    rk, uk = pick(cols, r"rank"), pick(cols, r"univ")
    sk = pick(cols, r"business\s*school", r"^school$", r"b-?school", r"school")
    out = []
    for _, r in df.iterrows():
        s = {"ranking": str(r[rk]).strip() if rk else "", "university": str(r[uk]).strip() if uk else "",
             "school": str(r[sk]).strip() if sk else ""}
        if s["university"] or s["school"]:
            out.append(s)
    return out

def df_to_carry(df):
    """Read a v4 Stage-1 CSV into the carry/department structure (with force_check)."""
    cols = list(df.columns)
    dept_c = pick(cols, r"school\s*department", r"department.*own", r"^department$", r"unit")
    rk, uk = pick(cols, r"rank"), pick(cols, r"univ")
    sk = pick(cols, r"business\s*school", r"^school$", r"b-?school", r"school$")
    rossk = pick(cols, r"ross\s*area"); ck = pick(cols, r"check\s*fac", r"individual")
    urlk = pick(cols, r"url|source"); confk = pick(cols, r"conf")
    tkk = pick(cols, r"t&o\s*subfield", r"t.o\s*subfield", r"subfield")
    if not dept_c:
        raise ValueError("This doesn't look like a v4 Stage-1 CSV (no 'School Department' column). "
                         "Re-export from the Match step, or run Match here.")
    by = {}
    for _, r in df.iterrows():
        uni = str(r[uk]).strip() if uk else ""; sch = str(r[sk]).strip() if sk else ""
        dept = str(r[dept_c]).strip() if dept_c else ""
        if not dept or not (uni or sch):
            continue
        e = by.setdefault((uni, sch), {"ranking": str(r[rk]).strip() if rk else "",
                                       "university": uni, "school": sch, "departments": []})
        ross = valid_ross(str(r[rossk]) if rossk and pd.notna(r[rossk]) else "")
        to_kind = normalize_to_kind(str(r[tkk])) if (tkk and pd.notna(r[tkk]) and "Technology and Operations" in ross) else ""
        force = None
        if ck and pd.notna(r[ck]):
            force = str(r[ck]).strip().lower() in ("yes", "true", "1", "on")
        e["departments"].append({"name": dept, "ross_areas": ross, "to_kind": to_kind,
                                 "url": str(r[urlk]).strip() if urlk and pd.notna(r[urlk]) else "",
                                 "confidence": str(r[confk]).strip().lower() if confk and pd.notna(r[confk]) else "",
                                 "notes": "", "force_check": force if force is not None else default_force_check(ross, to_kind)})
    return list(by.values())

def run_batch(items, fn, concurrency):
    """items: [(key, payload)]; fn(payload)->data (NO st.* inside). Shows progress."""
    results, total, done = {}, len(items), 0
    bar = st.progress(0.0); status = st.empty()
    with cf.ThreadPoolExecutor(max_workers=max(1, concurrency)) as ex:
        futs = {ex.submit(fn, p): k for k, p in items}
        for fut in cf.as_completed(futs):
            k = futs[fut]
            try:
                results[k] = ("done", fut.result())
            except Exception as e:
                results[k] = ("error", str(e))
            done += 1
            bar.progress(done / total)
            status.markdown(f"Working… **{done} / {total}**  ·  ~${ss.meter.dollars():.2f} so far")
    bar.empty(); status.empty()
    return results

def conf_dot(c):
    col = {"high": GREEN, "medium": AMBER, "low": GREY}.get(c, GREY)
    return f"<span class='rt-dot' style='background:{col}'></span>{c}" if c else ""

def badges(areas, to_kind=""):
    if not areas:
        return f"<span style='color:{GREY};font-style:italic'>No Ross equivalent</span>"
    out = []
    for a in areas:
        label = ROSS_SHORT.get(a, a)
        if a == "Technology and Operations" and to_kind:
            label += f"·{to_kind}"
        out.append(f"<span class='rt-pill'>{label}</span>")
    return " ".join(out)

def group_faculty_by_area(faculty):
    g = {a: [] for a in ROSS_AREAS}; g["Unclassified"] = []
    for f in faculty:
        a = f.get("ross_area") if f.get("ross_area") in ROSS_AREAS else "Unclassified"
        g[a].append(f)
    out = []
    for a in ROSS_AREAS + ["Unclassified"]:
        fl = g[a]
        if fl:
            fl.sort(key=lambda f: (RANK_ORDER.get(f.get("rank"), 9),
                                   (f.get("name", "").split()[-1].lower() if f.get("name") else "")))
            out.append((a, fl))
    return out

model = None  # set in sidebar

# ---- task builders (no st.* inside) ----
def make_match_task(client, model, use_fetch, meter):
    def task(s):
        user = (f"Target school: {s['university']} — {s['school']}.\n"
                "Enumerate this school's OWN academic units and map EACH to the Ross area(s) it fits, by the journals "
                "its faculty publish in. A unit may map to one Ross area, several, or none. Apply the "
                "Operations/Information-Systems rule, set to_kind for T&O units, and run the mandatory per-area coverage "
                "check before returning. Return ONLY the JSON.")
        parsed = call_model_json(client, meter, system=MATCH_SYSTEM, user=user, model=model,
                                  use_web_fetch=use_fetch, max_tokens=3000)
        deps = parse_departments(parsed)
        gaps = [a for a, ok in ross_coverage(deps).items() if not ok]
        if gaps:  # immediate automatic re-check for missing Ross areas
            dep_list = "; ".join(d["name"] for d in deps) or "(none found)"
            cu = (f"School: {s['university']} — {s['school']}\nDepartments already found: {dep_list}\n"
                  f"Ross areas still UNMAPPED: {', '.join(gaps)}\n"
                  "For each unmapped area, find the department whose faculty cover it (often a combined unit, e.g. Strategy "
                  "inside a Management department) and report present=true with that department's name, or present=false if "
                  "the school truly lacks it. Judge by journals. Return ONLY the JSON.")
            try:
                cov = call_model_json(client, meter, system=COVERAGE_SYSTEM, user=cu, model=model,
                                       use_web_fetch=use_fetch, max_tokens=1500)
                deps, _absent = merge_coverage(deps, cov, gaps)
            except Exception:
                pass
        return deps
    return task

def make_faculty_task(client, model, use_fetch, meter, recheck=False, prior=None):
    def task(payload):
        s, d = payload
        return fetch_department_faculty(client, meter, model, use_fetch, s, d, recheck=recheck, prior=prior)
    return task

# --------------------------------------------------------------------------- #
#  Sidebar                                                                    #
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.markdown("### Settings")
    model_choice = st.selectbox("Model", [DEFAULT_MODEL, FLAGSHIP, CHEAP_MODEL], index=0,
                                help="gpt-5.4 = balanced default · gpt-5.5 = most capable (pricier) · gpt-5.4-mini = cheapest.")
    model = (st.text_input("…or exact model id", value=model_choice,
                           help="Override if your OpenAI account uses a different model string.").strip() or model_choice)
    ss.meter.model = model
    use_fetch = st.toggle("Fetch directory pages (reads full pages)", value=True)
    concurrency = st.slider("Parallel lookups", 1, 6, 3)
    ss["use_fetch"], ss["concurrency"] = use_fetch, concurrency

    st.markdown("✅ API key detected" if get_api_key() else
                "⚠️ **No API key.** Add `OPENAI_API_KEY` in Streamlit secrets (or env var).")
    st.divider()
    st.markdown("### Cost this session")
    p = PRICE.get(model, {"in": 5.0, "out": 30.0})
    st.metric("Estimated spend", f"${ss.meter.dollars():.2f}")
    st.caption(f"{ss.meter.calls} calls · {ss.meter.in_tok:,} in · {ss.meter.out_tok:,} out · {ss.meter.searches} searches\n\n"
               f"Rate: ${p['in']}/M in · ${p['out']}/M out · ${WEB_SEARCH_PRICE_PER_1K}/1k searches")
    if st.button("Reset everything"):
        for k in ["match_schools", "match_results", "carry", "fac_schools", "fac_results"]:
            ss[k] = [] if "schools" in k else ({} if "results" in k else None)
        ss.meter = CostMeter(model); st.rerun()

# --------------------------------------------------------------------------- #
#  Header + nav                                                               #
# --------------------------------------------------------------------------- #
st.markdown("<div class='rt-banner'><div class='t'>🎓 Ross Area &amp; Faculty Toolkit</div>"
            "<div class='s'>List each school's departments and map them to Ross's seven areas by research field, "
            "then pull each department's complete current faculty — classified individually where it matters.</div></div>",
            unsafe_allow_html=True)
c1, c2, _ = st.columns([1, 1, 4])
if c1.button("① Match Areas", use_container_width=True, type=("primary" if ss.page == "match" else "secondary")):
    ss.page = "match"; st.rerun()
if c2.button("② Find Faculty", use_container_width=True, type=("primary" if ss.page == "faculty" else "secondary")):
    ss.page = "faculty"; st.rerun()
st.write("")

# =========================================================================== #
#  PAGE 1 — AREA MATCHER (reversed)                                            #
# =========================================================================== #
def page_match():
    client = get_client()
    st.subheader("1 · Load the school list")
    cc = st.columns([1, 1])
    if cc[0].button("Load built-in US News Top 50", use_container_width=True):
        ss.match_schools = [{"ranking": r, "university": u, "school": s} for r, u, s in BUILTIN]; ss.match_results = {}
    up = cc[1].file_uploader("…or upload Ranking / University / Business School (.xlsx / .csv)",
                             type=["xlsx", "xls", "csv"])
    if up is not None:
        try:
            ss.match_schools = df_to_schools(read_uploaded(up)); ss.match_results = {}
        except Exception as e:
            st.error(f"Could not read that file: {e}")
    if not ss.match_schools:
        st.info("Load a list to begin."); return
    st.success(f"{len(ss.match_schools)} schools loaded.")

    st.subheader("2 · Map each school's departments to Ross areas")
    disabled = client is None
    res = ss.match_results
    not_done = [s for s in ss.match_schools
                if (res.get((s["university"], s["school"]), {}).get("status")) != "done"]
    b1, b2 = st.columns([1, 1])
    if b1.button(f"▶ Map {len(not_done)} schools" + (" (all)" if len(not_done) == len(ss.match_schools) else " (remaining)"),
                 type="primary", disabled=(disabled or not not_done)):
        task = make_match_task(client, model, ss.use_fetch, ss.meter)
        items = [((s["university"], s["school"]), s) for s in not_done]
        raw = run_batch(items, task, ss.concurrency)
        for k, (st_, val) in raw.items():
            ss.match_results[k] = ({"status": "done", "deps": val, "error": ""} if st_ == "done"
                                   else {"status": "error", "deps": None, "error": val})
        st.rerun()
    b2.caption("Or expand any school below and run it on its own — you don't have to do all 50 at once.")
    if disabled:
        st.warning("Add an API key (sidebar) to enable matching.")
    with st.expander("How does this work?"):
        st.markdown("For each school, the model lists the school's **own** departments and maps each to the Ross area(s) it "
                    "fits — judged by the **journals its faculty publish in**, not the department name. A department can "
                    "map to several areas (a *Management* group spanning OB and strategy → both M&O and Strategy) or to "
                    "none. **If any of the 7 Ross areas is left unmapped, the model immediately re-checks** — that field is "
                    "often hiding inside a combined unit (Strategy inside *Management*, Economics inside a policy group) — "
                    "and either attaches it or confirms the school truly lacks it. Operations and Information Systems both "
                    "map to **Technology & Operations**; a combined OM+IS unit is auto-flagged so Stage 2 tags each person "
                    "**OM** or **IS** by their journals. Each department's **Check faculty individually** switch is on by "
                    "default when the match isn't unique (or it's a combined OM+IS unit); flip it on for any match you don't trust.")

    st.subheader("3 · Results")
    done = sum(1 for v in res.values() if v["status"] == "done")
    failed = [k for k, v in res.items() if v["status"] == "error"]
    n_deps = sum(len(v["deps"]) for v in res.values() if v["status"] == "done" and v["deps"])
    n_flagged = sum(1 for v in res.values() if v["status"] == "done" and v["deps"]
                    for d in v["deps"] if d["force_check"])
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Schools matched", f"{done}/{len(ss.match_schools)}")
    m2.metric("Departments", n_deps)
    m3.metric("Flagged for per-faculty", n_flagged)
    m4.metric("Spend", f"${ss.meter.dollars():.2f}")

    if failed and not disabled and st.button(f"↻ Retry {len(failed)} failed"):
        task = make_match_task(client, model, ss.use_fetch, ss.meter)
        by = {(s["university"], s["school"]): s for s in ss.match_schools}
        items = [(k, by[k]) for k in failed if k in by]
        raw = run_batch(items, task, ss.concurrency)
        for k, (st_, val) in raw.items():
            ss.match_results[k] = ({"status": "done", "deps": val, "error": ""} if st_ == "done"
                                   else {"status": "error", "deps": None, "error": val})
        st.rerun()

    view = st.radio("View", ["Departments", "Matrix"], horizontal=True, label_visibility="collapsed")
    flt = st.text_input("Filter", placeholder="Filter schools…", label_visibility="collapsed")
    schools = [s for s in ss.match_schools
               if flt.lower() in (s["university"] + " " + s["school"]).lower()] if flt else ss.match_schools

    if view == "Matrix":
        render_matrix(schools)
    else:
        for s in schools:
            render_school(client, s)

    st.divider()
    st.markdown("#### Download results")
    csv1 = _stage1_csv()
    n_rows = max(0, csv1.decode("utf-8").count("\n") - 1)  # minus header
    d1, d2 = st.columns([1, 1])
    if n_rows > 0:
        d1.download_button(f"⬇ Download matches CSV · {n_rows} departments", csv1,
                           "ross_area_matches.csv", "text/csv", type="primary", use_container_width=True)
    else:
        d1.caption("Match at least one school to enable the CSV download.")
    if d2.button("Send to Faculty step →", type="primary", disabled=(done == 0), use_container_width=True):
        ss.carry = _carry_from_results()
        ss.fac_schools = [dict(s, departments=[dict(d) for d in s["departments"]]) for s in ss.carry]
        ss.fac_results = {}
        ss.page = "faculty"; st.rerun()
    st.caption("This CSV (one row per department, with the Ross mapping, the *Check faculty individually* flag and T&O "
               "subfield) is also exactly what the Find Faculty step reads — so you can save it and re-upload it later "
               "instead of re-running the matching.")

def _one_school(client, s):
    key = (s["university"], s["school"])
    task = make_match_task(client, model, ss.use_fetch, ss.meter)
    with st.spinner(f"Mapping {s['school']}…"):
        try:
            ss.match_results[key] = {"status": "done", "deps": task(s), "error": ""}
        except Exception as e:
            ss.match_results[key] = {"status": "error", "deps": None, "error": str(e)}

def render_school(client, s):
    key = (s["university"], s["school"])
    r = ss.match_results.get(key, {"status": "idle", "deps": None, "error": ""})
    head = f"{s.get('ranking', '—')}  ·  {s['school']} — {s['university']}"
    head += {"done": "  ✓", "error": "  ⚠"}.get(r["status"], "")
    with st.expander(head, expanded=False):
        if r["status"] in ("idle", "error"):
            label = "↻ Retry this school" if r["status"] == "error" else "▶ Match this school"
            if st.button(label, key=f"one_{s['university']}_{s['school']}", disabled=(client is None)):
                _one_school(client, s); st.rerun()
        if r["status"] == "error":
            st.error(r["error"] or "Lookup failed."); return
        deps = r["deps"]
        if r["status"] != "done" or deps is None:
            st.caption("Not matched yet — click **Match this school** above to run just this one."); return
        gaps = [ROSS_SHORT[a] for a, ok in ross_coverage(deps).items() if not ok]
        if gaps:
            st.caption(f"No department mapped to **{', '.join(gaps)}** even after the automatic coverage re-check — "
                       "the school appears to genuinely lack these. Use *Re-check mapping* / *Re-match* if you disagree.")
        for i, d in enumerate(deps):
            left, mid, right = st.columns([3, 1.4, 1.2])
            src = f" · [source]({d['url']})" if d.get("url") else ""
            note = f"<br><span style='color:#6b7787;font-size:12px'>{d['notes']}</span>" if d.get("notes") else ""
            chk = f" · <span style='color:{VIOLET}'>✔ {d['checked']}</span>" if d.get("checked") else ""
            left.markdown(f"**{d['name']}** {conf_dot(d.get('confidence'))}{src}<br>{badges(d['ross_areas'], d.get('to_kind',''))}{chk}{note}",
                          unsafe_allow_html=True)
            tkey = f"force_{s['university']}_{s['school']}_{i}"
            d["force_check"] = mid.toggle("Check faculty individually", value=d["force_check"], key=tkey)
            if right.button("Re-check mapping", key=f"vrf_{s['university']}_{s['school']}_{i}", disabled=(client is None)):
                _verify_dept(client, s, d); st.rerun()
        if st.button("↻ Re-match this school", key=f"rm_{s['university']}_{s['school']}", disabled=(client is None)):
            _one_school(client, s); st.rerun()

def _verify_dept(client, s, d):
    user = (f"School: {s['university']} — {s['school']}\nUnit: {d['name']}\n"
            f"Reviewer flagged the current Ross mapping as possibly wrong: "
            f"\"{' ; '.join(d['ross_areas']) if d['ross_areas'] else 'No Ross equivalent'}\"\n"
            "Re-investigate ONLY this unit, by faculty research field / journals, applying the "
            "Operations/Information-Systems rule. Return the corrected JSON.")
    with st.spinner(f"Re-checking {d['name']}…"):
        try:
            c = call_model_json(client, ss.meter, system=VERIFY_SYSTEM, user=user, model=model,
                                 use_web_fetch=True, max_tokens=1300)
            d["ross_areas"] = valid_ross(c.get("ross_areas", []))
            d["url"] = c.get("url") or d.get("url", "")
            d["confidence"] = (c.get("confidence") or d.get("confidence", "")).lower()
            d["notes"] = c.get("notes") or d.get("notes", "")
            d["checked"] = c.get("checked", "")
            # the toggle keeps its current state; flip it manually if the new mapping warrants
        except Exception as e:
            st.error(f"Re-check failed: {e}")

def render_matrix(schools):
    data, conf, idx = [], [], []
    for s in schools:
        r = ss.match_results.get((s["university"], s["school"]), {"status": "idle", "deps": None})
        deps = r.get("deps") or []
        cell = {a: [] for a in ROSS_AREAS}
        for d in deps:
            for a in d["ross_areas"]:
                label = d["name"]
                if a == "Technology and Operations" and d.get("to_kind"):
                    label += f" ({d['to_kind']})"
                cell[a].append(label)
        row = {ROSS_SHORT[a]: (" ; ".join(cell[a]) if cell[a] else ("—" if r["status"] == "done" else
               ("…" if r["status"] in ("running",) else ""))) for a in ROSS_AREAS}
        crow = {ROSS_SHORT[a]: (r["status"] == "done" and not cell[a]) for a in ROSS_AREAS}  # True = gap
        data.append(row); conf.append(crow); idx.append(s["school"])
    if not data:
        st.info("Nothing to show."); return
    df = pd.DataFrame(data, index=idx); gdf = pd.DataFrame(conf, index=idx)
    styles = gdf.copy()
    for col in styles.columns:
        styles[col] = gdf[col].map(lambda gap: f"background-color:{'#fbecec' if gap else '#eef6ef'}")
    st.dataframe(df.style.apply(lambda _: styles, axis=None), use_container_width=True,
                 height=min(640, 70 + 35 * len(df)))
    st.caption("Each cell lists the school's department(s) mapped to that Ross area. 🟥 = gap (no department mapped). "
               "Columns: " + " · ".join(f"{ROSS_SHORT[a]}={a}" for a in ROSS_AREAS))

def _carry_from_results():
    out = []
    for s in ss.match_schools:
        r = ss.match_results.get((s["university"], s["school"]))
        if not r or r["status"] != "done" or not r["deps"]:
            continue
        out.append({"ranking": s["ranking"], "university": s["university"], "school": s["school"],
                    "departments": [dict(d) for d in r["deps"]]})
    return out

def _stage1_csv():
    rows = []
    for s in ss.match_schools:
        r = ss.match_results.get((s["university"], s["school"]))
        deps = r["deps"] if (r and r["status"] == "done") else None
        for d in (deps or []):
            rows.append({"Ranking": s["ranking"], "University": s["university"], "Business School": s["school"],
                         "School Department": d["name"],
                         "Ross Area(s)": " ; ".join(d["ross_areas"]) if d["ross_areas"] else "No Ross equivalent",
                         "Unique Match": "yes" if len(d["ross_areas"]) == 1 else "no",
                         "Check Faculty Individually": "yes" if d["force_check"] else "no",
                         "T&O Subfield": d.get("to_kind", ""),
                         "Confidence": d.get("confidence", ""), "Source URL": d.get("url", ""),
                         "Notes": d.get("notes", "")})
    return pd.DataFrame(rows).to_csv(index=False).encode("utf-8")

# =========================================================================== #
#  PAGE 2 — FACULTY FINDER                                                     #
# =========================================================================== #
def page_faculty():
    client = get_client()
    st.subheader("1 · Load matched departments")
    if ss.fac_schools:
        nd = sum(len(s["departments"]) for s in ss.fac_schools)
        st.success(f"Using {len(ss.fac_schools)} schools · {nd} departments from the Match step. Ready to find faculty below.")
        cc = st.columns([1, 1])
        if ss.carry and cc[0].button("↻ Reload latest Match-step results"):
            ss.fac_schools = [dict(s, departments=[dict(d) for d in s["departments"]]) for s in ss.carry]
            ss.fac_results = {}; st.rerun()
        up = cc[1].file_uploader("Or upload a different Stage-1 CSV", type=["csv", "xlsx", "xls"], label_visibility="collapsed")
        if up is not None:
            try:
                ss.fac_schools = df_to_carry(read_uploaded(up)); ss.fac_results = {}; st.rerun()
            except Exception as e:
                st.error(f"Could not read that file: {e}")
    else:
        if ss.carry:
            st.info("Results from the Match step are ready in this session (toggle states included).")
            if st.button("Use Match-step results →", type="primary"):
                ss.fac_schools = [dict(s, departments=[dict(d) for d in s["departments"]]) for s in ss.carry]
                ss.fac_results = {}; st.rerun()
        up = st.file_uploader("…or upload the Stage-1 CSV (ross_area_matches.csv)", type=["csv", "xlsx", "xls"])
        if up is not None:
            try:
                ss.fac_schools = df_to_carry(read_uploaded(up)); ss.fac_results = {}; st.rerun()
            except Exception as e:
                st.error(f"Could not read that file: {e}")
        st.info("Load Match-step results or a Stage-1 CSV to begin."); return

    include_unmapped = st.checkbox("Also pull faculty for 'No Ross equivalent' departments", value=False)
    jobs = [(s, d) for s in ss.fac_schools for d in s["departments"]
            if d["ross_areas"] or include_unmapped]
    n_classify = sum(1 for _s, d in jobs if d["force_check"])
    st.success(f"{len(ss.fac_schools)} schools · {len(jobs)} departments "
               f"({n_classify} will classify faculty individually, {len(jobs) - n_classify} inherit their area).")

    st.subheader("2 · Find current faculty")
    if not ss.use_fetch:
        st.warning("Page fetching is off — rosters may be less complete. Turn it on in the sidebar for best results.")
    disabled = client is None
    fr = ss.fac_results
    not_done = [(s, d) for s, d in jobs
                if (fr.get((s["university"], s["school"], d["name"]), {}).get("status")) != "done"]
    if st.button(f"▶ Find faculty · {len(not_done)} departments" + (" (all)" if len(not_done) == len(jobs) else " (remaining)"),
                 type="primary", disabled=(disabled or not not_done)):
        task = make_faculty_task(client, model, ss.use_fetch, ss.meter)
        items = [((s["university"], s["school"], d["name"]), (s, d)) for s, d in not_done]
        raw = run_batch(items, task, ss.concurrency)
        for k, (st_, val) in raw.items():
            ss.fac_results[k] = ({"status": "done", "data": val, "error": ""} if st_ == "done"
                                 else {"status": "error", "data": None, "error": val})
        st.rerun()
    if disabled:
        st.warning("Add an API key (sidebar) to enable lookups.")
    st.caption("Only ladder ranks (Assistant / Associate / Full Professor) are kept. Departments **flagged** in Step 1 are "
               "classified person-by-person by journals; unique-match departments inherit their single Ross area. "
               "You can also expand any single department below and run it on its own.")

    st.subheader("3 · Faculty by area")
    done = sum(1 for v in ss.fac_results.values() if v["status"] == "done")
    failed = [k for k, v in ss.fac_results.items() if v["status"] == "error"]
    total_fac = sum(len(v["data"]["faculty"]) for v in ss.fac_results.values() if v["status"] == "done" and v["data"])
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Departments done", f"{done}/{len(jobs)}"); m2.metric("Faculty found", total_fac)
    m3.metric("Failed", len(failed)); m4.metric("Spend", f"${ss.meter.dollars():.2f}")

    if failed and not disabled and st.button(f"↻ Retry {len(failed)} failed"):
        task = make_faculty_task(client, model, ss.use_fetch, ss.meter)
        by = {(s["university"], s["school"], d["name"]): (s, d) for s, d in jobs}
        items = [(k, by[k]) for k in failed if k in by]
        raw = run_batch(items, task, ss.concurrency)
        for k, (st_, val) in raw.items():
            ss.fac_results[k] = ({"status": "done", "data": val, "error": ""} if st_ == "done"
                                 else {"status": "error", "data": None, "error": val})
        st.rerun()

    fdf = _faculty_df(include_unmapped)
    st.markdown("#### Download results")
    if not fdf.empty:
        partial = "" if done == len(jobs) else f"  (partial — {done} of {len(jobs)} departments done so far)"
        st.download_button(f"⬇ Download master roster CSV · {len(fdf)} faculty{partial}",
                           fdf.to_csv(index=False).encode("utf-8"),
                           "ross_faculty_by_area.csv", "text/csv", type="primary", use_container_width=True)
        st.caption("One row per faculty member: school, department, rank, Ross area, T&O subfield (OM/IS), and the journal "
                   "evidence. You can download partial results at any time and again once more departments finish.")
        counts = fdf[fdf["Faculty Ross Area"] != ""]["Faculty Ross Area"].value_counts()
        if not counts.empty:
            st.bar_chart(counts)
        with st.expander(f"Preview full roster table · {len(fdf)} faculty", expanded=False):
            st.dataframe(fdf, use_container_width=True, height=460)
    else:
        st.info("No results to download yet. Run at least one department — use **Find faculty** above, or expand a "
                "single department in the detailed view below and run it on its own. The download button appears here as "
                "soon as any department finishes.")

    st.markdown("##### Detailed view — run a whole school, or expand a single department")
    ranks = st.multiselect("Ranks", list(RANK_ORDER), default=list(RANK_ORDER))
    names = [f"{s.get('ranking', '—')} · {s['school']}" for s in ss.fac_schools]
    choice = st.selectbox("School", names)
    school = ss.fac_schools[names.index(choice)]

    school_jobs = [d for d in school["departments"] if d["ross_areas"] or include_unmapped]
    school_not_done = [d for d in school_jobs
                       if (ss.fac_results.get((school["university"], school["school"], d["name"]), {}).get("status")) != "done"]
    rc1, rc2 = st.columns([1, 1])
    if rc1.button(f"▶ Find faculty for all {len(school_not_done)} departments in {school['school']}"
                  + ("" if len(school_not_done) == len(school_jobs) else " (remaining)"),
                  disabled=(client is None or not school_not_done), use_container_width=True):
        task = make_faculty_task(client, model, ss.use_fetch, ss.meter)
        items = [((school["university"], school["school"], d["name"]), (school, d)) for d in school_not_done]
        raw = run_batch(items, task, ss.concurrency)
        for k, (st_, val) in raw.items():
            ss.fac_results[k] = ({"status": "done", "data": val, "error": ""} if st_ == "done"
                                 else {"status": "error", "data": None, "error": val})
        st.rerun()
    if not fdf.empty:
        sdf = fdf[(fdf["University"] == school["university"]) & (fdf["Business School"] == school["school"])]
        if not sdf.empty:
            rc2.download_button(f"⬇ Download {school['school']} roster · {len(sdf)} faculty",
                                sdf.to_csv(index=False).encode("utf-8"),
                                f"roster_{school['school'].replace(' ', '_')}.csv", "text/csv", use_container_width=True)
    for d in school["departments"]:
        if d["ross_areas"] or include_unmapped:
            render_dept(client, school, d, ranks)

def render_dept(client, s, d, ranks):
    key = (s["university"], s["school"], d["name"])
    r = ss.fac_results.get(key, {"status": "idle", "data": None})
    flag = "classified" if d["force_check"] else "inherited"
    head = f"{d['name']}  ·  [{', '.join(ROSS_SHORT.get(a, a) for a in d['ross_areas']) or 'no Ross area'}]  ·  {flag}"
    if r["status"] == "done" and r["data"]:
        head += f"  ·  {len(r['data']['faculty'])} faculty  ·  {'complete' if r['data']['complete'] else 'partial'}"
    with st.expander(head, expanded=False):
        if r["status"] == "error":
            st.error(r["error"] or "Lookup failed.")
        if r["status"] in ("idle", "error") and client is not None:
            label = "↻ Retry this department" if r["status"] == "error" else "▶ Find faculty"
            if st.button(label, key=f"find_{key}"):
                _one_dept(client, s, d, recheck=False); st.rerun()
        data = r.get("data")
        if r["status"] == "done" and data:
            for area, fl in group_faculty_by_area(data["faculty"]):
                shown = [f for f in fl if f.get("rank") in ranks]
                if not shown:
                    continue
                area_badge = badges([area]) if area in ROSS_AREAS else "<span class='rt-pill'>?</span>"
                st.markdown(f"{area_badge} "
                            f"<span class='rt-area'>{area}</span> <span style='color:#6b7787'>({len(shown)})</span>",
                            unsafe_allow_html=True)
                for f in shown:
                    dot = RANK_COLOR.get(f.get("rank"), GREY)
                    sub = f"  <span class='rt-pill'>{f.get('subfield')}</span>" if f.get("subfield") else ""
                    ev = f"  —  *{f.get('field_evidence')}*" if f.get("field_evidence") else ""
                    st.markdown(f"<span class='rt-dot' style='background:{dot}'></span>**{f.get('name')}**{sub} "
                                f"<span style='color:#6b7787;font-size:12px'>{f.get('title','')}{ev}</span>",
                                unsafe_allow_html=True)
            if data.get("sources"):
                st.caption("Sources: " + "  ".join(f"[{i+1}]({u})" for i, u in enumerate(data["sources"][:5])))
            if data.get("notes"):
                st.caption(data["notes"])
            if client is not None and st.button("🔄 Recheck (scan for anyone missed)", key=f"rc_{key}"):
                _one_dept(client, s, d, recheck=True); st.rerun()

def _one_dept(client, s, d, recheck):
    key = (s["university"], s["school"], d["name"])
    prior = None
    if recheck:
        cur = ss.fac_results.get(key)
        prior = cur["data"]["faculty"] if cur and cur.get("data") else None
    task = make_faculty_task(client, model, ss.use_fetch, ss.meter, recheck=recheck, prior=prior)
    with st.spinner("Reading the faculty directory…"):
        try:
            ss.fac_results[key] = {"status": "done", "data": task((s, d)), "error": ""}
        except Exception as e:
            ss.fac_results[key] = {"status": "error", "data": None, "error": str(e)}

def _faculty_df(include_unmapped):
    rows = []
    for s in ss.fac_schools:
        for d in s["departments"]:
            if not d["ross_areas"] and not include_unmapped:
                continue
            r = ss.fac_results.get((s["university"], s["school"], d["name"]))
            if not r or r["status"] != "done" or not r["data"]:
                continue
            data = r["data"]
            dept_ross = " ; ".join(d["ross_areas"]) if d["ross_areas"] else "No Ross equivalent"
            comp = "yes" if data["complete"] else "no"
            clf = "yes" if d["force_check"] else "no"
            for f in data["faculty"]:
                rows.append({"Ranking": s["ranking"], "University": s["university"], "Business School": s["school"],
                             "School Department": d["name"], "Dept Ross Area(s)": dept_ross,
                             "Checked Individually": clf, "Faculty Name": f.get("name", ""), "Rank": f.get("rank", ""),
                             "Faculty Ross Area": f.get("ross_area", ""), "T&O Subfield": f.get("subfield", ""),
                             "Field Evidence": f.get("field_evidence", ""), "Complete": comp})
    return pd.DataFrame(rows)

# --------------------------------------------------------------------------- #
#  Route                                                                      #
# --------------------------------------------------------------------------- #
if ss.page == "match":
    page_match()
else:
    page_faculty()
st.divider()
st.caption("Matches are judged by faculty research field / journals; faculty lists come from official directories. "
           "Verify anything high-stakes against each school's pages (source links provided throughout).")
