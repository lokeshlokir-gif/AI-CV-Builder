#!/usr/bin/env python3
"""AI CV Builder — Streamlit Web App
Run locally: streamlit run streamlit_app.py
Deploy free: https://streamlit.io/cloud
"""
import streamlit as st
import json, re, time, io, ssl
import urllib.request, urllib.error
from datetime import datetime

# ═══ SSL ═══
def _get_ssl():
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where()), "certifi"
    except: pass
    try:
        ctx = ssl.create_default_context()
        urllib.request.urlopen("https://www.google.com", timeout=5, context=ctx)
        return ctx, "system"
    except: pass
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx, "unverified"

SSL_CTX, SSL_M = _get_ssl()

# ═══ Optional imports ═══
HAS_PDF = False
try:
    from pypdf import PdfReader
    HAS_PDF = True
except:
    try:
        from PyPDF2 import PdfReader
        HAS_PDF = True
    except: pass

HAS_DOCX = False
try:
    from docx import Document
    HAS_DOCX = True
except: pass

HAS_REQ = False
try:
    import requests as rq
    HAS_REQ = True
except: pass

# ═══ Models ═══
GEMINI_MODELS = [
    ("gemini-2.5-flash", "2.5 Flash (FREE)"),
    ("gemini-2.5-flash-lite", "2.5 Flash-Lite"),
    ("gemini-2.5-pro", "2.5 Pro"),
    ("gemini-2.0-flash", "2.0 Flash"),
]
OPENAI_MODELS = [
    ("gpt-4o", "GPT-4o (Best)"),
    ("gpt-4o-mini", "GPT-4o Mini (Cheap)"),
    ("gpt-4-turbo", "GPT-4 Turbo"),
    ("gpt-3.5-turbo", "GPT-3.5 Turbo"),
]
FB_ = [m[0] for m in GEMINI_MODELS]
MR = 3
RDL = 8

# ═══ Domains for Mock Interview ═══
DOMAINS = [
    "General", "AI / Machine Learning", "Data Science / Analytics", "Cyber Security",
    "Cloud / DevOps", "Software Engineering", "Web Development", "Mobile Development",
    "Database / SQL", "Networking", "Project Management", "Agile / Scrum",
    "Finance / Banking", "Law / Legal", "Healthcare / Medical", "Marketing / Digital",
    "Sales / Business Dev", "HR / Recruitment", "Education / Teaching",
    "Agriculture / Environment", "Supply Chain / Logistics", "Manufacturing / Engineering"
]

# ═══ Regions ═══
RD_ = {}
RD_["\U0001f1ec\U0001f1e7 United Kingdom"] = dict(code="UK", term="CV", pages="2p A4", photo="No", pn=False, pi="Name,phone,email,LinkedIn", sec="Profile,Key Skills,Experience,Education,Additional", sty="Achievement-focused,UK English", sp="Profile essential.", ats="ATS critical.", av="Photos,DOB")
RD_["\U0001f1fa\U0001f1f8 United States"] = dict(code="US", term="Resume", pages="1-2p", photo="No", pn=False, pi="Name,phone,email,LinkedIn", sec="Summary,Competencies,Experience,Education,Certs", sty="STAR method,US English", sp="US English.", ats="ATS CRITICAL.", av="Photos,DOB,SSN")
RD_["\U0001f1e8\U0001f1e6 Canada"] = dict(code="CA", term="Resume", pages="1-2p", photo="No", pn=False, pi="Name,phone,email,LinkedIn", sec="Summary,Qualifications,Experience,Education", sty="Achievement-focused", sp="Volunteer valued.", ats="ATS required.", av="Photos,DOB")
RD_["\U0001f1e9\U0001f1ea Germany"] = dict(code="DE", term="Lebenslauf", pages="2-3p A4", photo="Yes", pn=True, pi="Name,address,phone,email,DOB", sec="Personliche Daten,Berufserfahrung,Ausbildung,Kenntnisse", sty="Formal,chronological", sp="Photo EXPECTED.", ats="Growing.", av="Gaps")
RD_["\U0001f1eb\U0001f1f7 France"] = dict(code="FR", term="CV", pages="1-2p A4", photo="Common", pn=True, pi="Name,address,phone,email", sec="Etat Civil,Experience,Formation,Competences", sty="Formal,CEFR", sp="CEFR levels.", ats="Growing.", av="Casual")
RD_["\U0001f1ea\U0001f1f8 Spain"] = dict(code="ES", term="CV", pages="1-2p A4", photo="Expected", pn=True, pi="Name,address,phone,email,DOB", sec="Datos,Experiencia,Formacion,Idiomas", sty="Europass", sp="Europass.", ats="Europass.", av="Missing data")
RD_["\U0001f1ee\U0001f1f9 Italy"] = dict(code="IT", term="CV", pages="2-3p A4", photo="Expected", pn=True, pi="Name,address,phone,email,DOB", sec="Dati,Esperienza,Istruzione,Competenze", sty="Europass", sp="Privacy MANDATORY.", ats="Europass.", av="No privacy")
RD_["\U0001f1ee\U0001f1f3 India"] = dict(code="IN", term="Resume/CV", pages="2-3p A4", photo="Optional", pn=False, pi="Name,phone,email,LinkedIn", sec="Objective,Skills,Experience,Education,Projects", sty="Education detailed", sp="Declaration.", ats="Growing.", av="Excessive personal")
RD_["\U0001f1e6\U0001f1ea UAE / Gulf"] = dict(code="AE", term="CV", pages="2-4p A4", photo="Recommended", pn=True, pi="Name,phone,email,nationality,visa,DOB", sec="Summary,Experience,Education,Skills,Languages", sty="Detailed", sp="Nationality+visa MUST.", ats="Less ATS.", av="Missing nationality")
RD_["\U0001f1ef\U0001f1f5 Japan"] = dict(code="JP", term="Rirekisho", pages="1-2p", photo="Required", pn=True, pi="Name,address,phone,email,DOB", sec="Personal,Education,Work,Qualifications,Motivation", sty="STRICT template", sp="Rirekisho.", ats="Template.", av="Free-format")
RD_["\U0001f1e8\U0001f1f3 China"] = dict(code="CN", term="Resume", pages="1p A4", photo="Expected", pn=True, pi="Name,phone,email,DOB", sec="Personal,Education,Experience,Skills", sty="Concise", sp="985/211.", ats="Growing.", av="Multi-page")
RD_["\U0001f1e6\U0001f1fa Australia"] = dict(code="AU", term="Resume/CV", pages="2-3p A4", photo="No", pn=False, pi="Name,phone,email,LinkedIn", sec="Summary,Skills,Employment,Education,Referees", sty="Achievement-focused", sp="Referees EXPECTED.", ats="ATS recommended.", av="Photos")
RD_["\U0001f1f8\U0001f1ec Singapore"] = dict(code="SG", term="Resume/CV", pages="1-2p A4", photo="Optional", pn=False, pi="Name,phone,email,nationality", sec="Summary,Skills,Experience,Education,Certs", sty="Concise", sp="Work pass.", ats="ATS used.", av="Missing pass")
RD_["\U0001f1ea\U0001f1fa EU (Europass)"] = dict(code="EU", term="Europass CV", pages="2-3p A4", photo="Optional", pn=False, pi="Name,address,phone,email", sec="Personal,Experience,Education,Skills(CEFR)", sty="Standardised", sp="EU format.", ats="Machine-readable.", av="Non-standard")
RL = list(RD_.keys())

# ═══ AI Engine ═══
def _post(url, data, headers=None, to=120):
    b = json.dumps(data).encode()
    hdrs = {"Content-Type": "application/json"}
    if headers: hdrs.update(headers)
    r = urllib.request.Request(url, data=b, headers=hdrs, method="POST")
    try:
        with urllib.request.urlopen(r, timeout=to, context=SSL_CTX) as rp:
            return rp.status, rp.read().decode()
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode()
    except Exception as e:
        return 0, str(e)

def _gem(pr, key, model, to=120):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
    pl = {"contents": [{"parts": [{"text": pr}]}], "generationConfig": {"temperature": 0.7, "maxOutputTokens": 8192}}
    if HAS_REQ:
        try:
            r = rq.post(url, json=pl, timeout=to, headers={"Content-Type": "application/json"})
            st_code, bd = r.status_code, r.text
        except:
            try:
                r = rq.post(url, json=pl, timeout=to, headers={"Content-Type": "application/json"}, verify=False)
                st_code, bd = r.status_code, r.text
            except:
                st_code, bd = _post(url, pl, to=to)
    else:
        st_code, bd = _post(url, pl, to=to)
    if st_code == 200:
        res = json.loads(bd) if isinstance(bd, str) else bd
        return True, res["candidates"][0]["content"]["parts"][0]["text"]
    elif st_code == 429:
        return False, f"RATE_LIMITED|{model}"
    elif st_code == 404:
        return False, f"MODEL_NOT_FOUND|{model}"
    else:
        return False, f"[Error {st_code}] {bd[:200]}"

def _openai(pr, key, model, to=120):
    url = "https://api.openai.com/v1/chat/completions"
    pl = {"model": model, "messages": [{"role": "user", "content": pr}], "temperature": 0.7, "max_tokens": 4096}
    hdrs = {"Content-Type": "application/json", "Authorization": f"Bearer {key}"}
    if HAS_REQ:
        try:
            r = rq.post(url, json=pl, timeout=to, headers=hdrs)
            st_code, bd = r.status_code, r.text
        except:
            st_code, bd = _post(url, pl, headers=hdrs, to=to)
    else:
        st_code, bd = _post(url, pl, headers=hdrs, to=to)
    if st_code == 200:
        res = json.loads(bd) if isinstance(bd, str) else bd
        return True, res["choices"][0]["message"]["content"]
    elif st_code == 401:
        return False, "Invalid OpenAI key."
    else:
        return False, f"[Error {st_code}] {bd[:200]}"

def ai_call(pr, key, prov="gemini", model="gemini-2.5-flash"):
    try:
        if prov == "openai":
            ok, t = _openai(pr, key, model)
            return t
        for a in range(1, MR + 1):
            ok, t = _gem(pr, key, model)
            if ok: return t
            if "RATE_LIMITED" not in t and "MODEL_NOT_FOUND" not in t: return t
            if "MODEL_NOT_FOUND" in t: break
            if a < MR: time.sleep(RDL)
        for fb in FB_:
            if fb == model: continue
            ok, t = _gem(pr, key, fb)
            if ok: return f"[Used:{fb}]\n\n{t}"
            if "RATE_LIMITED" not in t: return t
            time.sleep(2)
        return "\u274c All models failed. Wait a minute or try a new key."
    except Exception as e:
        return f"[Error] {e}"

# ═══ File Reader ═══
def read_file(uploaded):
    if uploaded is None: return ""
    name = uploaded.name.lower()
    if name.endswith(".pdf"):
        if not HAS_PDF: return "Install pypdf"
        reader = PdfReader(io.BytesIO(uploaded.read()))
        return "\n".join(p.extract_text() or "" for p in reader.pages)
    elif name.endswith(".docx"):
        if not HAS_DOCX: return "Install python-docx"
        doc = Document(io.BytesIO(uploaded.read()))
        return "\n".join(p.text for p in doc.paragraphs)
    elif name.endswith(".txt"):
        return uploaded.read().decode("utf-8", errors="ignore")
    return ""

# ═══ Prompts ═══
NO_PRE = "RULES: 1)NO FABRICATION 2)Missing='[Not provided]' 3)EXPERIENCE=sum jobs (exclude gaps) 4)Only stated numbers 5)Only stated skills 6)Name first, no preamble 7)ACTUAL years 8)Enhance wording only 9)Region format"
def _rb(rn, ri): return f"TARGET:{rn}|{ri['term']}|{ri['pages']}|Photo:{ri['photo']}\nPersonal:{ri['pi']}\nSections:{ri['sec']}\nStyle:{ri['sty']}\nATS:{ri['ats']}\nAVOID:{ri['av']}"
def p_gen(jd, info, rn, ri): return f"Expert CV writer for {rn}.\n{_rb(rn, ri)}\n{NO_PRE}\nMatch JD keywords where experience aligns.\n---JD---\n{jd}\n---INFO---\n{info}"
def p_cmp(cv, jd, rn, ri): return f"Expert analyst for {rn}.\n{_rb(rn, ri)}\nJD MATCH X/100 | COMPLIANCE X/100 | MATCHES | GAPS | SUGGESTIONS\n---JD---\n{jd}\n---CV---\n{cv}"
def p_imp_cmp(cv, an, jd, rn, ri, kw="", kws="experience"):
    r = ("\nKEYWORD(ATS): MUST place EVERY keyword in Experience/Summary/Skills. Weave naturally." if kws == "force_all" else "\nKEYWORD: Only where experience supports. Weave naturally.")
    return f"Expert CV writer for {rn}.\n{_rb(rn, ri)}\n{NO_PRE}\nKeep ALL info. Fix issues.{r}\n---CV---\n{cv}\n---JD---\n{jd}\n---ANALYSIS---\n{an}\n---KEYWORDS---\n{kw}"
def p_cover(cv, jd, rn, ri): return f"Cover letter for {rn}.\n{NO_PRE}\nFrom CV+JD. **Bold** matches. ~350 words. KEY MATCHES section at end.\n---CV---\n{cv}\n---JD---\n{jd}\n---REGION---\n{_rb(rn, ri)}"
def p_ana(cv, rn, ri): return f"Reviewer for {rn}.\n{_rb(rn, ri)}\nSCORE X/100 | COMPLIANCE X/100 | STRENGTHS | IMPROVEMENTS | RECOMMENDATIONS\n---CV---\n{cv}"
def p_imp(cv, an, rn, ri): return f"CV writer for {rn}.\n{_rb(rn, ri)}\n{NO_PRE}\nFix ALL issues.\n---CV---\n{cv}\n---ANALYSIS---\n{an}"
def p_int(jd): return f"Interview coach.\nGenerate 15-20 questions from JD.\nTECHNICAL(10) + BEHAVIOURAL(5) + SITUATIONAL(5)\nFor EACH: question | why asked | answer framework (STAR) | key points\nBasic to advanced.\n---JD---\n{jd}"
def p_mock_jd(jd, n=15): return f"Interview coach.\nGenerate {n} questions from JD.\n60% Technical + 25% Behavioural + 15% Situational\nFormat:\nQ1: [question]\nQ2: [question]\n...Only questions.\n---JD---\n{jd}"
def p_mock_dom(dom, n=15): return f"Interview coach for {dom}.\nGenerate {n} questions.\n60% Technical + 25% Behavioural + 15% Situational\nBasic to advanced.\nFormat:\nQ1: [question]\nQ2: [question]\n...Only questions."
def p_mock_eval(q, a): return f"Interview coach. Evaluate:\nQuestion: {q}\nAnswer: {a}\n\nProvide:\nSCORE X/10 | STRENGTHS | WEAKNESSES | IDEAL ANSWER (STAR) | 3 TIPS | CONFIDENCE: Low/Med/High"
def p_coach(t, c): return f"Career coach: {t}\nContext: {c}\nInsights, Action plan, Pitfalls, Resources, Timeline"
def p_multi(cv, jds, rn, ri): return f"Analyst for {rn}.\n{_rb(rn, ri)}\nCompare CV vs MULTIPLE JDs.\nFor EACH: Title | Match X/100 | Matches | Gaps | Keywords\nOVERALL: best fit + why.\n---CV---\n{cv}\n\n{jds}"

INFO_T = """\u26a0\ufe0f Only provide REAL information.
Full Name:
Email:
Phone:
Location:
LinkedIn:

--- WORK EXPERIENCE ---
### Job 1:
Title:
Company:
From:
To:
What you did:

### Job 2:
Title:
Company:
From:
To:
What you did:

--- EDUCATION ---
Qualification:
Institution:
Year:

--- SKILLS ---
Technical:
Soft Skills:
Languages:

--- CERTIFICATIONS ---
Name:
Year:
"""

# ═══════════════ STREAMLIT APP ═══════════════
st.set_page_config(
    page_title="AI CV Builder",
    page_icon="\U0001f4c4",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Sidebar
with st.sidebar:
    st.markdown("# \U0001f4c4 AI CV Builder")
    st.markdown("---")

    # Navigation
    page = st.radio(
        "**Navigation**",
        ["\U0001f3e0 Home", "\U0001f4dd Generate CV", "\U0001f50d CV vs JD",
         "\U0001f4ca CV Analysis", "\U0001f4d1 Multi-JD Compare",
         "\U0001f3a4 Interview Prep", "\U0001f399\ufe0f Mock Interview",
         "\U0001f9d1\u200d\U0001f4bc Coaching", "\u2699\ufe0f Settings"],
        label_visibility="collapsed"
    )

    st.markdown("---")
    st.markdown("**\U0001f916 AI Provider**")
    provider = st.radio("Provider", ["Gemini (FREE)", "OpenAI (Paid)"], label_visibility="collapsed")
    prov = "openai" if "OpenAI" in provider else "gemini"
    models = OPENAI_MODELS if prov == "openai" else GEMINI_MODELS
    model_display = st.selectbox("Model", [f"{m[0]} ({m[1]})" for m in models])
    model_id = model_display.split(" (")[0]

    api_key = st.text_input("\U0001f511 API Key", type="password", placeholder="Paste your API key")

    if prov == "gemini":
        st.markdown("\U0001f517 [Get FREE Gemini key](https://aistudio.google.com/apikey)")
    else:
        st.markdown("\U0001f517 [Get OpenAI key](https://platform.openai.com/api-keys)")

    st.markdown("---")
    st.markdown("**\U0001f30d Target Region**")
    region = st.selectbox("Region", RL, label_visibility="collapsed")

    st.caption(f"SSL: {SSL_M}")

# Helper
def call_ai(prompt):
    if not api_key:
        st.error("\u26a0\ufe0f Please enter your API key in the sidebar.")
        return None
    with st.spinner("\U0001f916 AI is thinking..."):
        return ai_call(prompt, api_key, prov, model_id)

# ═══════════════ HOME ═══════════════
if page == "\U0001f3e0 Home":
    st.title("\U0001f4c4 AI CV Builder")
    st.markdown("### Build, analyse & optimise your CV with AI")
    st.markdown("**Gemini (FREE) + OpenAI** \u2022 14 Regions \u2022 Anti-hallucination \u2022 Mobile-friendly")
    st.markdown("---")

    col1, col2 = st.columns(2, gap="medium")
    with col1:
        with st.container(border=True):
            st.markdown("### \U0001f4dd Generate CV")
            st.write("Create a professional, region-formatted CV from a job description. Download as text.")
        with st.container(border=True):
            st.markdown("### \U0001f4ca CV Analysis")
            st.write("AI scores your CV for ATS compliance & region fit. Get improvements + improved CV.")
        with st.container(border=True):
            st.markdown("### \U0001f3a4 Interview Prep")
            st.write("15-20 tailored questions with answer frameworks (STAR) for your specific JD.")
        with st.container(border=True):
            st.markdown("### \U0001f9d1\u200d\U0001f4bc Career Coaching")
            st.write("Personalised advice on 10 career topics: salary negotiation, job search, networking and more.")

    with col2:
        with st.container(border=True):
            st.markdown("### \U0001f50d CV vs JD")
            st.write("Compare your CV against a JD. Get gap analysis, improved CV with keywords, and a tailored cover letter.")
        with st.container(border=True):
            st.markdown("### \U0001f4d1 Multi-JD Compare")
            st.write("Compare your CV against multiple JDs at once. AI ranks best fit with match scores per role.")
        with st.container(border=True):
            st.markdown("### \U0001f399\ufe0f Mock Interview")
            st.write("Practice with AI-generated questions from a JD or 22 domains. Get STAR feedback + score on your answers.")
        with st.container(border=True):
            st.markdown("### \U0001f30d 14 Regions Supported")
            st.write("UK, US, Canada, Germany, France, Spain, Italy, India, UAE, Japan, China, Australia, Singapore, EU.")

    st.markdown("---")
    st.info("\U0001f4a1 **Quick Start:** Pick provider \u2192 paste API key \u2192 select region \u2192 use any tab from the sidebar!")

# ═══════════════ GENERATE CV ═══════════════
elif page == "\U0001f4dd Generate CV":
    st.title("\U0001f4dd Generate CV from Job Description")
    st.caption("Paste a JD + fill your info \u2192 AI generates a region-formatted CV.")

    jd_gen = st.text_area("\U0001f4cb Job Description", height=200, key="gen_jd")
    info_gen = st.text_area("\U0001f464 Your Information", height=400, key="gen_info", value=INFO_T)

    if st.button("\u2728 Generate CV", type="primary", use_container_width=True):
        if not jd_gen.strip():
            st.warning("Please paste a Job Description.")
        elif not info_gen.strip():
            st.warning("Please fill in your information.")
        else:
            result = call_ai(p_gen(jd_gen, info_gen, region, RD_[region]))
            if result:
                st.session_state["gen_result"] = result

    if "gen_result" in st.session_state:
        st.markdown("### \u2728 Generated CV")
        st.text_area("Result", st.session_state["gen_result"], height=500, key="gen_out", label_visibility="collapsed")
        st.download_button("\U0001f4be Download CV (.txt)", st.session_state["gen_result"], "generated_cv.txt", use_container_width=True)

# ═══════════════ CV vs JD ═══════════════
elif page == "\U0001f50d CV vs JD":
    st.title("\U0001f50d CV vs Job Description")
    st.caption("Compare \u2192 improve with keywords \u2192 generate cover letter.")

    col1, col2 = st.columns(2, gap="medium")
    with col1:
        st.markdown("### \U0001f4c4 Your CV")
        cv_file = st.file_uploader("Upload CV (PDF/DOCX/TXT)", type=["pdf", "docx", "txt"], key="cmp_cv_file")
        cv_text_input = st.text_area("Or paste CV text", height=300, key="cmp_cv_input")
        cv_text = read_file(cv_file) if cv_file else cv_text_input

    with col2:
        st.markdown("### \U0001f4cb Job Description")
        jd_cmp = st.text_area("Paste JD", height=300, key="cmp_jd", label_visibility="collapsed")

    if st.button("\U0001f50d Compare CV vs JD", type="primary", use_container_width=True):
        if not cv_text.strip() or not jd_cmp.strip():
            st.warning("Please provide both CV and JD.")
        else:
            result = call_ai(p_cmp(cv_text, jd_cmp, region, RD_[region]))
            if result:
                st.session_state["cmp_result"] = result
                st.session_state["cmp_cv"] = cv_text
                st.session_state["cmp_jd"] = jd_cmp

    if "cmp_result" in st.session_state:
        st.markdown("### \U0001f4ca Comparison Result")
        st.text_area("Result", st.session_state["cmp_result"], height=400, key="cmp_out", label_visibility="collapsed")

        st.markdown("---")
        st.markdown("### \u2728 Improve CV with Keywords")

        col_k1, col_k2 = st.columns(2)
        with col_k1:
            kw_strength = st.radio("Keyword Strength", ["Experience-based", "Force ALL"], key="kw_str", horizontal=True)
        with col_k2:
            kw_mode = st.radio("Keywords from", ["Auto from JD", "Manual"], key="kw_mode", horizontal=True)

        manual_kw = ""
        if kw_mode == "Manual":
            manual_kw = st.text_input("Enter keywords (comma-separated)", placeholder="Docker, Kubernetes, CI/CD, Agile")

        if st.button("\u2728 Improve CV", type="primary", use_container_width=True):
            strength = "force_all" if kw_strength == "Force ALL" else "experience"
            kw_str = manual_kw if kw_mode == "Manual" else ""
            result = call_ai(p_imp_cmp(
                st.session_state["cmp_cv"],
                st.session_state["cmp_result"],
                st.session_state["cmp_jd"],
                region, RD_[region], kw_str, strength
            ))
            if result:
                st.session_state["cmp_improved"] = result

        if "cmp_improved" in st.session_state:
            st.markdown("#### \u2728 Improved CV")
            st.text_area("Improved", st.session_state["cmp_improved"], height=500, key="cmp_imp_out", label_visibility="collapsed")
            st.download_button("\U0001f4be Download Improved CV", st.session_state["cmp_improved"], "improved_cv.txt", use_container_width=True, key="dl_imp")

        st.markdown("---")
        st.markdown("### \U0001f4e8 Cover Letter")

        if st.button("\U0001f4e8 Generate Cover Letter", type="primary", use_container_width=True):
            cv_src = st.session_state.get("cmp_improved", st.session_state.get("cmp_cv", ""))
            result = call_ai(p_cover(cv_src, st.session_state["cmp_jd"], region, RD_[region]))
            if result:
                st.session_state["cmp_cl"] = result

        if "cmp_cl" in st.session_state:
            st.text_area("Cover Letter", st.session_state["cmp_cl"], height=400, key="cmp_cl_out", label_visibility="collapsed")
            st.download_button("\U0001f4be Download Cover Letter", st.session_state["cmp_cl"], "cover_letter.txt", use_container_width=True, key="dl_cl")

# ═══════════════ CV ANALYSIS ═══════════════
elif page == "\U0001f4ca CV Analysis":
    st.title("\U0001f4ca CV Analysis")
    st.caption("Upload CV \u2192 AI scores ATS compliance \u2192 improve \u2192 export.")

    cv_ana_file = st.file_uploader("Upload CV", type=["pdf", "docx", "txt"], key="ana_file")
    cv_ana_input = st.text_area("Or paste CV", height=300, key="ana_cv_input")
    cv_ana = read_file(cv_ana_file) if cv_ana_file else cv_ana_input

    st.info("**Score guide:** 90-100 Excellent | 75-89 Good | 60-74 Average | <60 Needs work")

    if st.button("\U0001f4ca Analyse CV", type="primary", use_container_width=True):
        if not cv_ana.strip():
            st.warning("Please provide a CV.")
        else:
            result = call_ai(p_ana(cv_ana, region, RD_[region]))
            if result:
                st.session_state["ana_result"] = result
                st.session_state["ana_cv"] = cv_ana

    if "ana_result" in st.session_state:
        st.markdown("### \U0001f4ca Analysis")
        st.text_area("Analysis", st.session_state["ana_result"], height=400, key="ana_out", label_visibility="collapsed")

        st.markdown("---")
        if st.button("\u2728 Improve CV", type="primary", use_container_width=True):
            result = call_ai(p_imp(st.session_state["ana_cv"], st.session_state["ana_result"], region, RD_[region]))
            if result:
                st.session_state["ana_improved"] = result

        if "ana_improved" in st.session_state:
            st.markdown("### \u2728 Improved CV")
            st.text_area("Improved", st.session_state["ana_improved"], height=500, key="ana_imp_out", label_visibility="collapsed")
            st.download_button("\U0001f4be Download Improved CV", st.session_state["ana_improved"], "improved_cv.txt", use_container_width=True, key="ana_dl")

# ═══════════════ MULTI-JD ═══════════════
elif page == "\U0001f4d1 Multi-JD Compare":
    st.title("\U0001f4d1 Multi-JD Compare")
    st.caption("Compare your CV against multiple JDs. AI ranks best fit.")

    cv_multi_file = st.file_uploader("Upload CV", type=["pdf", "docx", "txt"], key="multi_file")
    cv_multi_input = st.text_area("Or paste CV", height=200, key="multi_cv_input")
    cv_multi = read_file(cv_multi_file) if cv_multi_file else cv_multi_input

    n_jds = st.number_input("Number of JDs to compare", min_value=2, max_value=10, value=2, key="multi_n")

    jd_texts = []
    for i in range(int(n_jds)):
        jd_texts.append(st.text_area(f"\U0001f4cb JD {i+1}", height=150, key=f"multi_jd_{i}"))

    if st.button("\U0001f4d1 Compare All JDs", type="primary", use_container_width=True):
        if not cv_multi.strip():
            st.warning("Please provide a CV.")
        else:
            combined = "\n\n".join([f"--- JD {i+1} ---\n{t}" for i, t in enumerate(jd_texts) if t.strip()])
            if not combined.strip():
                st.warning("Please fill in at least one JD.")
            else:
                result = call_ai(p_multi(cv_multi, combined, region, RD_[region]))
                if result:
                    st.session_state["multi_result"] = result

    if "multi_result" in st.session_state:
        st.markdown("### \U0001f4ca Comparison Results")
        st.text_area("Results", st.session_state["multi_result"], height=500, key="multi_out", label_visibility="collapsed")
        st.download_button("\U0001f4be Download Results", st.session_state["multi_result"], "multi_jd_results.txt", use_container_width=True)

# ═══════════════ INTERVIEW PREP ═══════════════
elif page == "\U0001f3a4 Interview Prep":
    st.title("\U0001f3a4 Interview Preparation")
    st.caption("Paste JD \u2192 AI generates 15-20 questions with answer frameworks.")
    st.info("\U0001f4a1 Questions come from AI's knowledge of interview patterns (no web search).")

    jd_intv = st.text_area("\U0001f4cb Job Description", height=250, key="intv_jd")

    if st.button("\U0001f3af Generate Questions", type="primary", use_container_width=True):
        if not jd_intv.strip():
            st.warning("Please paste a JD.")
        else:
            result = call_ai(p_int(jd_intv))
            if result:
                st.session_state["intv_result"] = result

    if "intv_result" in st.session_state:
        st.markdown("### \U0001f3af Questions + Answer Frameworks")
        st.text_area("Questions", st.session_state["intv_result"], height=600, key="intv_out", label_visibility="collapsed")
        st.download_button("\U0001f4be Download Questions", st.session_state["intv_result"], "interview_questions.txt", use_container_width=True)

# ═══════════════ MOCK INTERVIEW ═══════════════
elif page == "\U0001f399\ufe0f Mock Interview":
    st.title("\U0001f399\ufe0f Mock Interview")
    st.caption("AI generates questions \u2192 you answer \u2192 AI evaluates with STAR + score.")

    with st.container(border=True):
        st.markdown("### \U0001f4cb Question Source")
        source = st.radio("Source", ["\U0001f4cb From Job Description", "\U0001f3af By Domain / Field"], horizontal=True, key="mock_src", label_visibility="collapsed")

        if "JD" in source:
            mock_jd = st.text_area("Paste JD", height=200, key="mock_jd")
        else:
            mock_domain = st.selectbox("Select Domain", DOMAINS, key="mock_dom")

        n_q = st.radio("Number of questions", ["10", "15", "20"], index=1, horizontal=True, key="mock_n")

    if st.button("\U0001f3af Generate Questions", type="primary", use_container_width=True):
        if "JD" in source:
            if not mock_jd.strip():
                st.warning("Please paste a JD.")
            else:
                result = call_ai(p_mock_jd(mock_jd, int(n_q)))
                if result:
                    st.session_state["mock_questions"] = result
        else:
            result = call_ai(p_mock_dom(mock_domain, int(n_q)))
            if result:
                st.session_state["mock_questions"] = result

    if "mock_questions" in st.session_state:
        st.markdown("### \U0001f4cb Generated Questions")
        st.text_area("Questions", st.session_state["mock_questions"], height=350, key="mock_q_out", label_visibility="collapsed")

        st.markdown("---")
        st.markdown("### \U0001f4dd Practice & Evaluate")
        st.caption("Copy a question from above (or type your own), write your answer, then evaluate.")

        mock_q = st.text_area("Question", height=80, key="mock_q_in")
        mock_a = st.text_area("Your Answer", height=200, key="mock_a_in")

        if st.button("\U0001f4ca Evaluate My Answer", type="primary", use_container_width=True):
            if not mock_q.strip() or not mock_a.strip():
                st.warning("Please provide both question and answer.")
            else:
                result = call_ai(p_mock_eval(mock_q, mock_a))
                if result:
                    st.session_state["mock_feedback"] = result

        if "mock_feedback" in st.session_state:
            st.markdown("### \U0001f4ca AI Feedback")
            st.text_area("Feedback", st.session_state["mock_feedback"], height=400, key="mock_fb_out", label_visibility="collapsed")
            st.download_button("\U0001f4be Download Feedback", st.session_state["mock_feedback"], "mock_feedback.txt", use_container_width=True)

# ═══════════════ COACHING ═══════════════
elif page == "\U0001f9d1\u200d\U0001f4bc Coaching":
    st.title("\U0001f9d1\u200d\U0001f4bc Career Coaching")
    st.caption("Pick a topic + describe your situation \u2192 personalised advice.")

    topic = st.selectbox("Topic", [
        "Career Change", "Salary Negotiation", "Skill Development", "Job Search",
        "Interview Confidence", "LinkedIn", "Networking", "Promotion",
        "Work-Life Balance", "Remote Work"
    ], key="coach_topic")

    context = st.text_area("Your situation/context", height=200, key="coach_ctx",
                           placeholder="Describe where you are now, what you want, and any specific challenges...")

    if st.button("\U0001f4a1 Get Advice", type="primary", use_container_width=True):
        if not context.strip():
            st.warning("Please describe your situation.")
        else:
            result = call_ai(p_coach(topic, context))
            if result:
                st.session_state["coach_result"] = result

    if "coach_result" in st.session_state:
        st.markdown("### \U0001f4a1 Personalised Advice")
        st.text_area("Advice", st.session_state["coach_result"], height=500, key="coach_out", label_visibility="collapsed")
        st.download_button("\U0001f4be Download Advice", st.session_state["coach_result"], "career_advice.txt", use_container_width=True)

# ═══════════════ SETTINGS ═══════════════
elif page == "\u2699\ufe0f Settings":
    st.title("\u2699\ufe0f Settings & Info")

    with st.container(border=True):
        st.markdown("### \U0001f916 AI Providers")
        st.markdown("- **Gemini (FREE):** https://aistudio.google.com/apikey")
        st.markdown("- **OpenAI (Paid):** https://platform.openai.com/api-keys")

    with st.container(border=True):
        st.markdown("### \U0001f319 Dark Mode")
        st.markdown("Click the **\u2261 menu** (top-right) \u2192 **Settings** \u2192 **Theme** to switch between Light/Dark/System.")

    with st.container(border=True):
        st.markdown("### \U0001f4f1 Use on Phone")
        st.markdown("**iPhone (Safari):** Share button \u2192 Add to Home Screen")
        st.markdown("**Android (Chrome):** \u22ee menu \u2192 Add to Home Screen / Install app")
        st.markdown("This makes the web app feel like a real installed app on your phone.")

    with st.container(border=True):
        st.markdown("### \U0001f399\ufe0f Mock Interview")
        st.markdown("AI generates 10-20 questions from JD or 22 domains. You answer, AI evaluates with STAR + score + tips.")

    with st.container(border=True):
        st.markdown("### \U0001f9e9 Keywords")
        st.markdown("In CV vs JD tab, choose 'Experience-based' (only adds where real experience supports) or 'Force ALL' (places every keyword for ATS).")

    with st.container(border=True):
        st.markdown("### \U0001f4e8 Cover Letter")
        st.markdown("Auto-generated from CV + JD. Key matches highlighted in **bold**.")

    st.markdown("---")
    st.caption(f"SSL: {SSL_M} | Built with Streamlit | AI CV Builder")
