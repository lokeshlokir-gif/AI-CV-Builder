
#!/usr/bin/env python3
"""AI CV Builder - Streamlit Web App
Region-aware CV (14 regions), Word/PDF export, clickable home cards,
130+ mock interview domains, experience-level tailoring, anti-repeat AI
with persistent history, mobile-safe sidebar, emoji-stripped exports,
Gemini + OpenAI + Claude providers, open My Library."""
import streamlit as st
import streamlit.components.v1 as components
import json, re, time, io, ssl, random, uuid
import urllib.request, urllib.error
from datetime import datetime

# ============================================================
# SSL SETUP
# ============================================================
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

# ============================================================
# OPTIONAL LIBRARIES
# ============================================================
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
    from docx.shared import Pt, RGBColor, Cm, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    HAS_DOCX = True
except: pass

HAS_RL = False
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib import colors
    HAS_RL = True
except: pass

HAS_REQ = False
try:
    import requests as rq
    HAS_REQ = True
except: pass

# ============================================================
# MODELS — Gemini + OpenAI + Claude
# ============================================================
GEMINI_MODELS = [
    ("gemini-2.5-flash", "2.5 Flash (FREE)"),
    ("gemini-2.5-flash-lite", "2.5 Flash-Lite"),
    ("gemini-2.5-pro", "2.5 Pro"),
    ("gemini-2.0-flash", "2.0 Flash"),
]
OPENAI_MODELS = [
    ("gpt-4o", "GPT-4o (Best)"),
    ("gpt-4o-mini", "GPT-4o Mini"),
    ("gpt-4-turbo", "GPT-4 Turbo"),
    ("gpt-3.5-turbo", "GPT-3.5 Turbo"),
]
CLAUDE_MODELS = [
    ("claude-opus-4-20250514", "Claude Opus 4 (Best)"),
    ("claude-sonnet-4-20250514", "Claude Sonnet 4"),
    ("claude-3-5-sonnet-20241022", "Claude 3.5 Sonnet"),
    ("claude-3-5-haiku-20241022", "Claude 3.5 Haiku (Fast)"),
]
FB_ = [m[0] for m in GEMINI_MODELS]
MR = 3
RDL = 8

# ============================================================
# MOCK INTERVIEW DOMAINS — 130+ now (Issue #5)
# Audited every category; added Computer Science Eng, Information
# Science Eng, IT, Biomedical, Robotics, Mining, Petroleum, Marine,
# Textile, Metallurgy, AYUSH (Ayurveda/Homeopathy/Unani), Nutrition,
# Animation/VFX/Game Design, Diploma/ITI/B.Tech/M.Tech, SSC, Railway,
# State PSC, Insurance, Retail, Hair Stylist, Tailoring, Chef, etc.
# ============================================================
DOMAIN_GROUPS = {
    "General": ["General"],
    "Technology": [
        "AI / Machine Learning", "Data Science / Analytics", "Cyber Security",
        "Cloud / DevOps", "Software Engineering", "Web Development",
        "Mobile Development", "Database / SQL", "Networking",
        "Blockchain / Web3", "AR / VR", "Embedded Systems",
    ],
    "Management & Business": [
        "Project Management", "Agile / Scrum", "Product Management",
        "Business Analyst", "Operations Management", "Strategy Consulting",
    ],
    "Finance & Sales": [
        "Finance / Banking", "Investment / Trading", "Accounting / CA",
        "Sales / Business Dev", "Marketing / Digital", "HR / Recruitment",
        "Supply Chain / Logistics", "Insurance", "Wealth Management",
    ],
    "Engineering": [
        "Computer Science Engineering (CSE)",
        "Information Science Engineering (ISE)",
        "Information Technology (IT)",
        "Electronics & Communication (ECE)",
        "Electrical & Electronics (EEE)",
        "Mechanical Engineering",
        "Civil Engineering",
        "Chemical Engineering",
        "Aerospace Engineering",
        "Automobile Engineering",
        "Industrial Engineering",
        "Biomedical Engineering",
        "Robotics & Mechatronics",
        "Mining Engineering",
        "Petroleum Engineering",
        "Marine Engineering",
        "Textile Engineering",
        "Metallurgy & Materials",
        "Biotechnology / Genetic Eng",
        "Architecture",
    ],
    "Medical & Healthcare": [
        "MBBS / Medicine", "BDS / Dentistry", "BPT / Physiotherapy",
        "Nursing (B.Sc)", "Pharma-B / Pharmacy", "Pharma-D",
        "Veterinary", "Public Health", "Medical Lab Technology",
        "Radiology", "Optometry",
        "AYUSH - BAMS (Ayurveda)", "AYUSH - BHMS (Homeopathy)",
        "AYUSH - BUMS (Unani)", "Yoga & Naturopathy",
        "Nutrition & Dietetics", "Occupational Therapy",
    ],
    "Education & Teaching": [
        "Teaching / School", "Higher Education / Lecturer",
        "Special Education", "Education Administration", "Tutoring",
        "B.Ed / D.Ed", "Early Childhood / Montessori",
    ],
    "Indian Education Streams": [
        "PUC PCMB", "PUC PCMC", "PUC Commerce", "PUC Arts",
        "Diploma / Polytechnic", "ITI",
        "BCA", "BBA / BBM", "B.A", "B.Com", "B.Sc",
        "B.Tech / B.E", "MBA", "M.Sc", "M.A", "M.Com",
        "M.Tech / M.E", "Ph.D / Research",
    ],
    "Law & Legal": [
        "Law / Legal (General)", "Corporate Law", "Criminal Law",
        "Civil Law", "IP Law", "Tax Law", "Constitutional Law",
        "Cyber Law",
    ],
    "Agriculture & Environment": [
        "Agriculture / Agronomy", "Horticulture", "Forestry",
        "Environmental Science", "Fisheries", "Dairy Technology",
        "Food Technology", "Sericulture",
    ],
    "Arts & Social Sciences": [
        "Geography", "History", "Psychology", "Sociology",
        "Political Science", "Economics", "Anthropology",
        "Philosophy", "Literature", "Linguistics",
        "Social Work (MSW)",
    ],
    "Creative & Media": [
        "Journalism", "Graphic Design", "UI/UX Design",
        "Photography", "Film Production", "Music",
        "Fashion Design", "Interior Design",
        "Animation", "VFX", "Game Design",
        "Content Creation / YouTube", "Copywriting",
    ],
    "Services & Hospitality": [
        "Hotel Management", "Travel & Tourism",
        "Aviation / Air Hostess", "Event Management",
        "Real Estate", "Retail Management", "Customer Service",
    ],
    "Government & Civil Services": [
        "UPSC / Civil Services", "State PSC",
        "Banking Exams (IBPS/SBI)", "SSC (CGL/CHSL)",
        "Railway Exams (RRB)", "Defence (NDA/CDS)",
        "Police / Constable",
    ],
    "Trade & Skilled": [
        "Plumbing / Electrical Trade", "Auto Mechanic",
        "Carpentry", "Welding", "Hair Stylist / Beautician",
        "Tailoring / Stitching", "Painting / Construction",
        "Chef / Cooking", "Driving / Heavy Vehicle",
    ],
}

EXP_LEVELS = [
    "🎓 Student / Fresher (0 yrs)",
    "👶 Junior (0-2 yrs)",
    "👨‍💻 Mid-level (2-5 yrs)",
    "🧑‍🔧 Senior (5-10 yrs)",
    "👑 Lead / Principal (10+ yrs)",
]

# ============================================================
# REGION DATA (14 regions — unchanged)
# ============================================================
RD_ = {}
RD_['🇬🇧 United Kingdom'] = dict(code='UK', term='CV', pages='2 pages A4', photo='NO photo - never include photo or DOB', pn=False, pi='Name, phone, email, LinkedIn (NO photo, NO DOB, NO nationality)', sec='Personal Profile, Key Skills, Work Experience, Education, Additional', sty='Achievement-focused with measurable results, UK English', sp='Personal Profile (3-4 lines) at top is ESSENTIAL.', ats='ATS-critical.', av='Photos, DOB, nationality, references', cp=(0, 51, 102))
RD_['🇺🇸 United States'] = dict(code='US', term='Resume', pages='1-2 pages Letter', photo='NO photo - never include photo or age', pn=False, pi='Name, phone, email, LinkedIn', sec='Summary, Core Competencies, Professional Experience, Education, Certifications', sty='STAR method with quantifiable results, US English', sp='Quantify everything with %, $, numbers.', ats='ATS-CRITICAL.', av='Photos, age, DOB, SSN', cp=(0, 51, 102))
RD_['🇨🇦 Canada'] = dict(code='CA', term='Resume', pages='1-2 pages', photo='NO photo', pn=False, pi='Name, phone, email, LinkedIn, city/province', sec='Summary, Qualifications, Experience, Education, Volunteer Experience', sty='Achievement-focused, bilingual EN/FR a bonus', sp='Volunteer experience is valued.', ats='ATS required.', av='Photos, DOB, SIN', cp=(139, 0, 0))
RD_['🇩🇪 Germany'] = dict(code='DE', term='Lebenslauf', pages='2-3 pages A4', photo='YES - include professional headshot photo placeholder at top-right (3.5x4.5cm)', pn=True, pi='Name, address, phone, email, DOB, nationality, marital status, photo placeholder', sec='Personliche Daten, Berufserfahrung, Ausbildung, Kenntnisse, Sprachen', sty='Formal, strict reverse-chronological, formal German tone', sp='Photo at top-right is EXPECTED. Include personal data (DOB, nationality).', ats='Growing ATS.', av='Casual tone, gaps', cp=(15, 52, 96))
RD_['🇫🇷 France'] = dict(code='FR', term='CV', pages='1-2 pages A4', photo='OPTIONAL - photo is common', pn=True, pi='Name, address, phone, email, age/DOB, optional photo', sec='Etat Civil, Experience Professionnelle, Formation, Competences, Langues (CEFR)', sty='Formal French CV with CEFR levels', sp='CEFR language levels mandatory.', ats='Growing ATS.', av='Overly casual', cp=(0, 35, 102))
RD_['🇪🇸 Spain'] = dict(code='ES', term='CV', pages='1-2 pages A4', photo='EXPECTED - small professional photo at top', pn=True, pi='Name, address, phone, email, DNI/NIE, DOB, photo placeholder', sec='Datos Personales, Experiencia, Formacion, Idiomas, Competencias', sty='Europass-style common', sp='Photo expected.', ats='Europass.', av='Missing personal data', cp=(170, 0, 0))
RD_['🇮🇹 Italy'] = dict(code='IT', term='CV', pages='2-3 pages A4', photo='EXPECTED - professional photo at top-right', pn=True, pi='Name, address, phone, email, DOB, nationality, photo placeholder', sec='Dati Personali, Esperienza, Istruzione, Competenze, Lingue', sty='Europass format with GDPR clause', sp='Photo expected. GDPR privacy clause MANDATORY at end.', ats='Europass.', av='Missing GDPR clause', cp=(0, 102, 51))
RD_['🇮🇳 India'] = dict(code='IN', term='Resume/CV', pages='2-3 pages A4', photo='OPTIONAL', pn=False, pi='Name, phone, email, LinkedIn, city', sec='Career Objective, Skills, Experience, Education (detailed), Projects, Declaration', sty='Education-focused, tech skills prominent', sp='Declaration statement at bottom is traditional. Education detailed.', ats='Growing ATS.', av='Missing declaration', cp=(255, 103, 0))
RD_['🇦🇪 UAE / Gulf'] = dict(code='AE', term='CV', pages='2-4 pages A4', photo='RECOMMENDED - professional photo at top', pn=True, pi='Name, phone, email, NATIONALITY (mandatory), VISA STATUS (mandatory), DOB, photo', sec='Personal Details, Career Objective, Summary, Experience, Education, Skills, Languages', sty='Detailed CV, nationality and visa status critical', sp='Nationality and visa status MANDATORY at top.', ats='Less ATS.', av='Missing nationality/visa', cp=(0, 100, 0))
RD_['🇯🇵 Japan'] = dict(code='JP', term='Rirekisho', pages='1-2 pages STRICT format', photo='REQUIRED - professional photo at top-right (3x4cm) in dedicated box', pn=True, pi='Name (kanji+furigana), address, phone, email, DOB, gender, photo', sec='Personal Info (with PHOTO BOX), Education (chronological), Work History (chronological), Qualifications, Motivation', sty='STRICT Rirekisho - chronological (oldest first), photo top-right, seal at bottom', sp='Photo REQUIRED at top-right. Chronological order (oldest first) MANDATORY. Mark a PHOTO BOX area clearly.', ats='Strict template.', av='Reverse-chronological order', cp=(153, 0, 0))
RD_['🇨🇳 China'] = dict(code='CN', term='Resume', pages='1 page A4', photo='EXPECTED - professional photo at top-right', pn=True, pi='Name, phone, email, DOB, gender, hukou, photo', sec='Personal Information, Education, Work Experience, Skills, Awards, Self-evaluation', sty='Concise one-pager, highlight 985/211 universities', sp='Photo expected. One page strictly.', ats='Growing.', av='Multi-page', cp=(204, 0, 0))
RD_['🇦🇺 Australia'] = dict(code='AU', term='Resume/CV', pages='2-3 pages A4', photo='NO photo', pn=False, pi='Name, phone, email, LinkedIn, city', sec='Career Summary, Key Skills, Employment History, Education, Referees', sty='Achievement-focused, named referees expected', sp='Referees with names and contact details EXPECTED at end.', ats='ATS recommended.', av='Photos, missing referees', cp=(0, 0, 128))
RD_['🇸🇬 Singapore'] = dict(code='SG', term='Resume/CV', pages='1-2 pages A4', photo='OPTIONAL', pn=False, pi='Name, phone, email, nationality, work pass status', sec='Summary, Skills, Experience, Education, Certifications, Languages', sty='Concise, work pass status important', sp='Work pass status (EP/SP/PR) is IMPORTANT.', ats='ATS widely used.', av='Missing work pass info', cp=(204, 0, 0))
RD_['🇪🇺 EU (Europass)'] = dict(code='EU', term='Europass CV', pages='2-3 pages A4', photo='OPTIONAL', pn=False, pi='Name, address, phone, email, nationality', sec='Personal Information, Work Experience, Education, Language Skills (CEFR), Digital Skills', sty='Standardised Europass format', sp='CEFR language levels required.', ats='Machine-readable.', av='Non-standard', cp=(0, 51, 153))
RL = list(RD_.keys())

# ============================================================
# STRIP EMOJIS / SPECIAL UNICODE FOR PDF/WORD SAFETY
# ============================================================
def strip_emojis(text):
    if not text:
        return ""
    text = re.sub(r'[\U0001F000-\U0001FFFF]', '', text)
    text = re.sub(r'[\U0001F1E6-\U0001F1FF]', '', text)
    text = re.sub(r'[\u2500-\u27BF]', '', text)
    text = re.sub(r'[\u2B00-\u2BFF]', '', text)
    text = re.sub(r'[\u2300-\u23FF]', '', text)
    text = re.sub(r'[\uFE00-\uFE0F]', '', text)
    text = text.replace('\u200d', '')
    for ch in '═━─━│┃┌┐└┘├┤┬┴┼╔╗╚╝╠╣╦╩╬':
        text = text.replace(ch, '')
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

# ============================================================
# EXTRACT CANDIDATE NAME FOR FILENAMES
# ============================================================
def extract_name(text):
    if not text:
        return "Candidate"
    SK = ["SUMMARY", "SKILL", "EXPERIENCE", "EDUCATION", "CERTIF", "PROFILE",
          "OBJECTIVE", "PERSONAL", "CAREER", "LANGUAGES", "ADDITIONAL",
          "DECLARATION", "REFERENCES", "QUALIFICATIONS", "WORK HISTORY",
          "MOTIVATION", "DEAR", "TO WHOM", "HIRING MANAGER", "COVER LETTER",
          "SCORE", "ANALYSIS", "FEEDBACK", "QUESTION", "ANSWER", "STRENGTH",
          "WEAKNESS", "IDEAL", "TIPS", "INSIGHT", "ACTION", "RECOMMENDATION"]
    raw = [l for l in strip_emojis(text).split("\n") if l.strip()]
    candidates = raw[:10] + raw[-5:]
    for line in candidates:
        c = re.sub(r"#+\s*", "", line).strip()
        c = c.strip("*").strip("- •▪").strip()
        c = re.sub(r"^(sincerely|regards|best regards|yours sincerely|"
                   r"yours truly|kind regards|warm regards),?\s*", "",
                   c, flags=re.I).strip()
        if not c:
            continue
        u = c.upper().replace(":", "")
        if any(k in u for k in SK):
            continue
        if c.lower().startswith(("here", "below", "i have", "this is",
                                  "based on", "dear", "to whom")):
            continue
        if len(c) < 2 or len(c) > 60:
            continue
        if not re.match(r"^[A-Za-z][A-Za-z\s\.\-']{1,}$", c):
            continue
        safe = re.sub(r"[^A-Za-z0-9_\-]+", "_", c).strip("_")
        if safe:
            return safe
    return "Candidate"

# ============================================================
# NETWORK / AI CALLS — Gemini + OpenAI + Claude (Bonus)
# ============================================================
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

def _gem(pr, key, model, to=120, temperature=0.7):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
    pl = {"contents": [{"parts": [{"text": pr}]}],
          "generationConfig": {"temperature": temperature,
                               "maxOutputTokens": 8192,
                               "topP": 0.95}}
    if HAS_REQ:
        try:
            r = rq.post(url, json=pl, timeout=to, headers={"Content-Type": "application/json"})
            sc, bd = r.status_code, r.text
        except:
            try:
                r = rq.post(url, json=pl, timeout=to, headers={"Content-Type": "application/json"}, verify=False)
                sc, bd = r.status_code, r.text
            except:
                sc, bd = _post(url, pl, to=to)
    else:
        sc, bd = _post(url, pl, to=to)
    if sc == 200:
        res = json.loads(bd) if isinstance(bd, str) else bd
        return True, res["candidates"][0]["content"]["parts"][0]["text"]
    elif sc == 429: return False, f"RATE_LIMITED|{model}"
    elif sc == 404: return False, f"MODEL_NOT_FOUND|{model}"
    else: return False, f"[Error {sc}] {bd[:200]}"

def _openai(pr, key, model, to=120, temperature=0.7):
    url = "https://api.openai.com/v1/chat/completions"
    pl = {"model": model, "messages": [{"role": "user", "content": pr}],
          "temperature": temperature, "max_tokens": 4096}
    hdrs = {"Content-Type": "application/json", "Authorization": f"Bearer {key}"}
    if HAS_REQ:
        try:
            r = rq.post(url, json=pl, timeout=to, headers=hdrs)
            sc, bd = r.status_code, r.text
        except:
            sc, bd = _post(url, pl, headers=hdrs, to=to)
    else:
        sc, bd = _post(url, pl, headers=hdrs, to=to)
    if sc == 200:
        res = json.loads(bd) if isinstance(bd, str) else bd
        return True, res["choices"][0]["message"]["content"]
    elif sc == 401: return False, "Invalid OpenAI key."
    else: return False, f"[Error {sc}] {bd[:200]}"

def _claude(pr, key, model, to=120, temperature=0.7):
    """Anthropic Claude API call."""
    url = "https://api.anthropic.com/v1/messages"
    pl = {"model": model, "max_tokens": 4096,
          "temperature": min(temperature, 1.0),
          "messages": [{"role": "user", "content": pr}]}
    hdrs = {"Content-Type": "application/json",
            "x-api-key": key,
            "anthropic-version": "2023-06-01"}
    if HAS_REQ:
        try:
            r = rq.post(url, json=pl, timeout=to, headers=hdrs)
            sc, bd = r.status_code, r.text
        except:
            sc, bd = _post(url, pl, headers=hdrs, to=to)
    else:
        sc, bd = _post(url, pl, headers=hdrs, to=to)
    if sc == 200:
        res = json.loads(bd) if isinstance(bd, str) else bd
        try:
            return True, res["content"][0]["text"]
        except:
            return False, f"[Parse Error] {str(res)[:300]}"
    elif sc == 401: return False, "Invalid Claude/Anthropic key."
    else: return False, f"[Error {sc}] {bd[:200]}"

def ai_call(pr, key, prov="gemini", model="gemini-2.5-flash", temperature=0.7):
    try:
        if prov == "openai":
            ok, t = _openai(pr, key, model, temperature=temperature)
            return t
        if prov == "claude":
            ok, t = _claude(pr, key, model, temperature=temperature)
            return t
        # Gemini with auto-fallback
        for a in range(1, MR + 1):
            ok, t = _gem(pr, key, model, temperature=temperature)
            if ok: return t
            if "RATE_LIMITED" not in t and "MODEL_NOT_FOUND" not in t: return t
            if "MODEL_NOT_FOUND" in t: break
            if a < MR: time.sleep(RDL)
        for fb in FB_:
            if fb == model: continue
            ok, t = _gem(pr, key, fb, temperature=temperature)
            if ok: return f"[Used:{fb}]\n\n{t}"
            if "RATE_LIMITED" not in t: return t
            time.sleep(2)
        return "All models failed."
    except Exception as e:
        return f"[Error] {e}"

# ============================================================
# FILE READER
# ============================================================
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

# ============================================================
# WORD EXPORT
# ============================================================
def make_docx(text, region):
    if not HAS_DOCX: return None
    text = strip_emojis(text)
    ri = RD_.get(region, {})
    cp = ri.get("cp", (0, 51, 102))
    has_photo = ri.get("pn", False)
    PC = RGBColor(*cp)
    doc = Document()
    for s in doc.sections:
        s.top_margin = Cm(1.5); s.bottom_margin = Cm(1.5)
        s.left_margin = Cm(2);  s.right_margin = Cm(2)
    lines = [l for l in text.split("\n") if l.strip()]
    SK = ["SUMMARY", "SKILL", "EXPERIENCE", "EDUCATION", "CERTIF", "PROFILE",
          "OBJECTIVE", "PERSONAL", "CAREER", "LANGUAGES", "ADDITIONAL",
          "DECLARATION", "REFERENCES", "QUALIFICATIONS", "WORK HISTORY", "MOTIVATION"]
    name = ""; cs = 0
    for idx, line in enumerate(lines):
        c = re.sub(r"#+\s*", "", line.strip()).strip("*").strip()
        if not c: continue
        u = c.upper().replace(":", "")
        if not name and not any(k in u for k in SK) and len(c) > 1 and \
           not c.lower().startswith(("here", "below", "i have", "this is", "based on")):
            name = c; cs = idx + 1; break
    if has_photo and name:
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        lc = table.rows[0].cells[0]; lc.width = Cm(12)
        p = lc.paragraphs[0]; r = p.add_run(name)
        r.font.size = Pt(22); r.font.bold = True; r.font.color.rgb = PC
        rc = table.rows[0].cells[1]; rc.width = Cm(4)
        pp = rc.paragraphs[0]
        pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pp.paragraph_format.space_before = Pt(20)
        pp.paragraph_format.space_after = Pt(20)
        pp.add_run("[ PHOTO ]\n3.5 x 4.5cm").font.size = Pt(10)
    else:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(name); r.font.size = Pt(24); r.font.bold = True; r.font.color.rgb = PC
    for i in range(cs, len(lines)):
        line = lines[i].strip()
        if not line: continue
        if line.lower().startswith(("here is", "below is", "i have", "this is", "based on")):
            continue
        c = re.sub(r"#+\s*", "", line).strip()
        cfc = c.strip("*").strip()
        if not cfc: continue
        u = cfc.upper().replace(":", "")
        is_h = False; ht = cfc
        if line.startswith("#"):
            is_h = True
        elif line.startswith("**") and line.endswith("**") and len(line) > 4:
            ht = line.strip("*").strip()
            if any(k in ht.upper().replace(":", "") for k in SK): is_h = True
        elif len(cfc) >= 3 and not cfc.startswith(("-", "*")):
            al = [x for x in cfc if x.isalpha()]
            if al and sum(1 for x in al if x.isupper()) / len(al) > 0.7 and any(k in u for k in SK):
                is_h = True
        if is_h:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run(ht.upper())
            r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = PC
        elif cfc.startswith(("-", "*")):
            bt = cfc.lstrip("-* ").strip()
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.8)
            p.add_run("• ").font.size = Pt(10)
            r = p.add_run(re.sub(r"\*\*(.+?)\*\*", r"\1", bt))
            r.font.size = Pt(10)
        else:
            p = doc.add_paragraph()
            for part in re.split(r"(\*\*.+?\*\*)", cfc):
                if part.startswith("**") and part.endswith("**"):
                    r = p.add_run(part.strip("*"))
                    r.font.size = Pt(10); r.font.bold = True
                else:
                    cl2 = re.sub(r"\*(.+?)\*", r"\1", part)
                    if cl2.strip():
                        r = p.add_run(cl2); r.font.size = Pt(10)
    if ri.get("code") == "IT":
        doc.add_paragraph()
        r = doc.add_paragraph().add_run(
            "Autorizzo il trattamento dei dati personali ai sensi del GDPR (UE 2016/679).")
        r.font.size = Pt(8); r.italic = True
    if ri.get("code") == "IN":
        doc.add_paragraph()
        r = doc.add_paragraph().add_run(
            "Declaration: I hereby declare that all information provided above is true to the best of my knowledge.")
        r.font.size = Pt(9); r.italic = True
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(20)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("AI CV Builder - " + datetime.now().strftime("%d %B %Y"))
    r.font.size = Pt(7); r.font.color.rgb = RGBColor(150, 150, 150); r.italic = True
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

# ============================================================
# PDF EXPORT
# ============================================================
def make_pdf(text, region):
    if not HAS_RL: return None
    text = strip_emojis(text)
    ri = RD_.get(region, {})
    cp = ri.get("cp", (0, 51, 102))
    has_photo = ri.get("pn", False)
    color_hex = "#%02X%02X%02X" % cp
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=1.5*cm, bottomMargin=1.5*cm,
                            leftMargin=2*cm, rightMargin=2*cm)
    ss = getSampleStyleSheet()
    ss.add(ParagraphStyle("CVName", fontSize=20, leading=24, alignment=TA_CENTER, spaceAfter=6, textColor=color_hex, fontName="Helvetica-Bold"))
    ss.add(ParagraphStyle("CVNameL", fontSize=20, leading=24, alignment=TA_LEFT, spaceAfter=6, textColor=color_hex, fontName="Helvetica-Bold"))
    ss.add(ParagraphStyle("CVHead", fontSize=12, leading=14, spaceBefore=12, spaceAfter=4, textColor=color_hex, fontName="Helvetica-Bold"))
    ss.add(ParagraphStyle("CVBody", fontSize=10, leading=13, spaceAfter=2, fontName="Helvetica"))
    ss.add(ParagraphStyle("CVBullet", fontSize=10, leading=13, leftIndent=20, spaceAfter=1, fontName="Helvetica"))
    ss.add(ParagraphStyle("CVPhoto", fontSize=9, leading=12, alignment=TA_CENTER, fontName="Helvetica"))
    lines = [l for l in text.split("\n") if l.strip()]
    SK = ["SUMMARY", "SKILL", "EXPERIENCE", "EDUCATION", "CERTIF", "PROFILE",
          "OBJECTIVE", "PERSONAL", "CAREER", "LANGUAGES", "ADDITIONAL",
          "DECLARATION", "REFERENCES", "QUALIFICATIONS", "WORK HISTORY", "MOTIVATION"]
    name = ""; cs = 0
    for idx, line in enumerate(lines):
        c = re.sub(r"#+\s*", "", line.strip()).strip("*").strip()
        if not c: continue
        u = c.upper().replace(":", "")
        if not name and not any(k in u for k in SK) and len(c) > 1 and \
           not c.lower().startswith(("here", "below", "i have", "this is", "based on")):
            name = c; cs = idx + 1; break
    story = []
    if has_photo and name:
        pt = Table([[Paragraph(name.replace("&", "&amp;"), ss["CVNameL"]),
                     Paragraph("[ PHOTO ]<br/>3.5 x 4.5 cm", ss["CVPhoto"])]],
                   colWidths=[12*cm, 4*cm])
        pt.setStyle(TableStyle([
            ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#E8E8E8")),
            ("BOX", (1, 0), (1, 0), 1, colors.HexColor("#999999")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (1, 0), (1, 0), 25),
            ("BOTTOMPADDING", (1, 0), (1, 0), 25),
        ]))
        story.append(pt); story.append(Spacer(1, 10))
    elif name:
        story.append(Paragraph(name.replace("&", "&amp;"), ss["CVName"]))
    for i in range(cs, len(lines)):
        s = lines[i].strip()
        if not s: continue
        if s.lower().startswith(("here is", "below is", "i have", "this is", "based on")):
            continue
        c = re.sub(r"#+\s*", "", s).strip("*").strip()
        if not c: continue
        u = c.upper().replace(":", "")
        safe = c.replace("&", "&amp;").replace("<", "&lt;")
        is_h = s.startswith("#") or (s.startswith("**") and s.endswith("**") and any(k in u for k in SK))
        if not is_h and len(c) >= 3:
            al = [x for x in c if x.isalpha()]
            if al and sum(1 for x in al if x.isupper()) / len(al) > 0.7 and any(k in u for k in SK):
                is_h = True
        if is_h:
            story.append(Paragraph(safe.upper(), ss["CVHead"]))
        elif c.startswith(("-", "*")):
            stripped = safe.lstrip("-* ")
            story.append(Paragraph("• " + stripped, ss["CVBullet"]))
        else:
            story.append(Paragraph(re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", safe), ss["CVBody"]))
    if ri.get("code") == "IT":
        story.append(Spacer(1, 12))
        story.append(Paragraph(
            "<i>Autorizzo il trattamento dei dati personali ai sensi del GDPR (UE 2016/679).</i>", ss["CVBody"]))
    if ri.get("code") == "IN":
        story.append(Spacer(1, 12))
        story.append(Paragraph(
            "<i>Declaration: I hereby declare that all information provided above is true to the best of my knowledge.</i>", ss["CVBody"]))
    story.append(Spacer(1, 20))
    story.append(Paragraph(
        "<i><font size=7 color='#999999'>AI CV Builder - " +
        datetime.now().strftime("%d %B %Y") + "</font></i>", ss["CVBody"]))
    doc.build(story); buf.seek(0)
    return buf.getvalue()

# ============================================================
# AI PROMPTS
# ============================================================
NO_PRE = "RULES: 1) NO FABRICATION 2) Missing=[Not provided] 3) EXPERIENCE=sum jobs(exclude gaps) 4) Only stated numbers 5) Only stated skills 6) Name first, no preamble 7) ACTUAL years 8) Enhance wording only 9) Strictly follow region format"

def _rb(rn, ri):
    return ("TARGET REGION: " + rn + "\n" +
            "DOCUMENT TYPE: " + ri["term"] + "\n" +
            "LENGTH: " + ri["pages"] + "\n" +
            "PHOTO REQUIREMENT: " + ri["photo"] + "\n" +
            "PERSONAL INFO TO INCLUDE: " + ri["pi"] + "\n" +
            "REQUIRED SECTIONS: " + ri["sec"] + "\n" +
            "STYLE: " + ri["sty"] + "\n" +
            "SPECIAL REQUIREMENTS: " + ri["sp"] + "\n" +
            "ATS: " + ri["ats"] + "\n" +
            "AVOID: " + ri["av"] + "\n\n" +
            "IMPORTANT: Format CV strictly for " + rn + ". If region requires photo, include [PHOTO PLACEHOLDER] at top. Use exact sections listed above. Do NOT use emojis, country flags, or special box-drawing characters in the output.")

def p_gen(jd, info, rn, ri):
    return ("You are an expert CV writer specifically for " + rn + ".\n" + _rb(rn, ri) + "\n" + NO_PRE +
            "\n\nGenerate a " + ri["term"] + " strictly following " + rn + " conventions. Match JD keywords where actual experience aligns." +
            "\n\n---JOB DESCRIPTION---\n" + jd + "\n\n---CANDIDATE INFO---\n" + info)
def p_cmp(cv, jd, rn, ri):
    return ("Expert analyst for " + rn + ".\n" + _rb(rn, ri) +
            "\nProvide: JD MATCH X/100 | REGION COMPLIANCE X/100 | KEY MATCHES | GAPS | SUGGESTIONS" +
            "\n---JD---\n" + jd + "\n---CV---\n" + cv)
def p_imp_cmp(cv, an, jd, rn, ri, kw="", kws="experience"):
    r = ("\nKEYWORD MODE (ATS): MUST place EVERY keyword. Weave naturally." if kws == "force_all"
         else "\nKEYWORD MODE: Only where experience supports. Weave naturally.")
    return ("Expert CV writer for " + rn + ".\n" + _rb(rn, ri) + "\n" + NO_PRE +
            "\nKeep ALL info. Fix issues from analysis. Reformat to strict " + rn + " standard." + r +
            "\n---CV---\n" + cv + "\n---JD---\n" + jd + "\n---ANALYSIS---\n" + an + "\n---KEYWORDS---\n" + kw)
def p_cover(cv, jd, rn, ri):
    return ("Write a cover letter for " + rn + ".\n" + NO_PRE +
            "\nFrom CV+JD. **Bold** key matches. ~350 words. End with KEY MATCHES section. Follow " + rn + " cover letter conventions." +
            "\n---CV---\n" + cv + "\n---JD---\n" + jd + "\n---REGION---\n" + _rb(rn, ri))
def p_ana(cv, rn, ri):
    return ("Expert reviewer for " + rn + ".\n" + _rb(rn, ri) +
            "\nProvide: SCORE X/100 | REGION COMPLIANCE X/100 | STRENGTHS | IMPROVEMENTS | RECOMMENDATIONS" +
            "\n---CV---\n" + cv)
def p_imp(cv, an, rn, ri):
    return ("Expert CV writer for " + rn + ".\n" + _rb(rn, ri) + "\n" + NO_PRE +
            "\nFix ALL issues from analysis. Reformat to strict " + rn + " standards." +
            "\n---CV---\n" + cv + "\n---ANALYSIS---\n" + an)
def p_int(jd):
    return ("Interview coach.\nGenerate 15-20 questions from JD.\nTECHNICAL(10) + BEHAVIOURAL(5) + SITUATIONAL(5)" +
            "\nFor EACH: question | why asked | answer framework (STAR) | key points\nBasic to advanced.\n---JD---\n" + jd)

# ============================================================
# ANTI-REPEAT MOCK PROMPTS — MUCH STRONGER (Issue #1)
# Uses persistent per-domain history + multiple sub-topic seeds +
# explicit topic-rotation instructions
# ============================================================
SUBTOPIC_SEEDS = [
    "technical fundamentals & theory",
    "real-world troubleshooting & debugging scenarios",
    "system design & architecture trade-offs",
    "recent industry trends, tools and best practices",
    "team collaboration & cross-functional communication",
    "edge cases, failures, and lessons learned",
    "performance optimisation and scalability",
    "security, compliance and ethics",
    "mentoring, leadership and conflict resolution",
    "career growth, learning curves and pivots",
]

def _build_avoid_block(history_questions):
    """Build a strong, explicit avoid-list block from previous questions."""
    if not history_questions:
        return ""
    # Last 30 questions, trimmed for token safety
    recent = history_questions[-30:]
    listing = "\n".join(f"- {q[:200]}" for q in recent)
    return ("\n\n=== CRITICAL ANTI-REPEAT RULE ===\n"
            "The following questions have ALREADY been asked in previous rounds for this same domain.\n"
            "You MUST NOT repeat them, paraphrase them, or ask any question covering the same concept.\n"
            "Generate a COMPLETELY DIFFERENT set covering NEW sub-topics, NEW angles, NEW scenarios.\n"
            "Previously asked (DO NOT REPEAT):\n" + listing + "\n=== END ANTI-REPEAT RULE ===\n")

def p_mock_jd(jd, n=15, exp_level="Mid-level (2-5 yrs)", history_questions=None):
    nonce = uuid.uuid4().hex[:12]
    seed = random.randint(10000, 99999)
    focus = random.sample(SUBTOPIC_SEEDS, k=min(4, len(SUBTOPIC_SEEDS)))
    avoid_block = _build_avoid_block(history_questions or [])
    return ("You are a CREATIVE and VARIED interview coach.\n"
            "EXPERIENCE LEVEL OF CANDIDATE: " + exp_level + "\n"
            "Tailor difficulty + scenarios for this level. For Students/Freshers focus on fundamentals, "
            "academic projects, internships. For Senior/Lead focus on architecture, leadership, trade-offs.\n\n"
            "ROTATE FOCUS — this round must emphasise these specific sub-topics (mix them across questions):\n"
            "- " + "\n- ".join(focus) + "\n\n"
            "Generate " + str(n) + " FRESH, UNIQUE, NON-OVERLAPPING questions from JD.\n"
            "60% Technical + 25% Behavioural + 15% Situational.\n"
            "Vary phrasing, scenario, depth. Use modern industry context.\n"
            "Session randomness token (internal — ignore but use to ensure variety): " + nonce + "-" + str(seed) + "\n\n"
            "OUTPUT FORMAT (strict):\nQ1: [question]\nQ2: [question]\n...Only the questions. No preamble, no commentary."
            + avoid_block + "\n---JD---\n" + jd)

def p_mock_dom(dom, n=15, exp_level="Mid-level (2-5 yrs)", history_questions=None):
    nonce = uuid.uuid4().hex[:12]
    seed = random.randint(10000, 99999)
    focus = random.sample(SUBTOPIC_SEEDS, k=min(4, len(SUBTOPIC_SEEDS)))
    avoid_block = _build_avoid_block(history_questions or [])
    return ("You are a CREATIVE and VARIED interview coach for **" + dom + "**.\n"
            "EXPERIENCE LEVEL OF CANDIDATE: " + exp_level + "\n"
            "Tailor difficulty for this level. For Students/Freshers focus on fundamentals, theory, college projects. "
            "For Senior/Lead focus on architecture, leadership, trade-offs, mentoring.\n\n"
            "ROTATE FOCUS — this round must emphasise these specific sub-topics of " + dom + " (mix them across questions):\n"
            "- " + "\n- ".join(focus) + "\n\n"
            "Generate " + str(n) + " FRESH, UNIQUE, NON-OVERLAPPING questions SPECIFICALLY for " + dom + ".\n"
            "Stay tightly focused on " + dom + " — do NOT drift into adjacent fields.\n"
            "60% Technical + 25% Behavioural + 15% Situational.\n"
            "Vary phrasing, scenario, depth. Use modern industry context, recent tools, real-world situations.\n"
            "Session randomness token (internal — ignore but use to ensure variety): " + nonce + "-" + str(seed) + "\n\n"
            "OUTPUT FORMAT (strict):\nQ1: [question]\nQ2: [question]\n...Only the questions. No preamble, no commentary."
            + avoid_block)

def p_mock_eval(q, a, exp_level="Mid-level (2-5 yrs)"):
    return ("Interview coach. Evaluate this answer for a " + exp_level + " candidate.\n"
            "Adjust expectations to that level (be encouraging for Students, demanding for Senior/Lead).\n"
            "Question: " + q + "\nAnswer: " + a +
            "\n\nProvide:\nSCORE X/10 | STRENGTHS | WEAKNESSES | IDEAL ANSWER (STAR) | 3 TIPS | CONFIDENCE: Low/Med/High")

def p_coach(t, c):
    return ("Career coach: " + t + "\nContext: " + c + "\nInsights, Action plan, Pitfalls, Resources, Timeline")
def p_multi(cv, jds, rn, ri):
    return ("Analyst for " + rn + ".\n" + _rb(rn, ri) +
            "\nCompare CV vs MULTIPLE JDs.\nFor EACH: Title | Match X/100 | Matches | Gaps | Keywords\nOVERALL: best fit + why." +
            "\n---CV---\n" + cv + "\n\n" + jds)

INFO_T = """Only provide REAL information.
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

# ============================================================
# HELPER — extract Q-only lines from AI output for history tracking
# ============================================================
def _parse_questions(text):
    if not text: return []
    out = []
    for line in text.split("\n"):
        line = line.strip()
        m = re.match(r"^Q\d+[:.\)]\s*(.+)$", line, flags=re.I)
        if m:
            out.append(m.group(1).strip())
    return out

# ============================================================
# STREAMLIT CONFIG + CSS
# ============================================================
st.set_page_config(page_title="AI CV Builder", page_icon="📄",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    /* Cards now have visual cue they are clickable */
    .home-card {
        padding: 24px; border-radius: 10px; margin-bottom: 12px;
        min-height: 110px; cursor: pointer; transition: transform 0.15s, box-shadow 0.15s;
        border: 2px solid transparent;
    }
    .home-card:hover { transform: translateY(-2px); box-shadow: 0 6px 16px rgba(0,0,0,0.15); }
    .home-card h3 { margin-top: 0; color: #1a1a2e; }
    .home-card p  { color: #444; margin: 0; }
    .home-card .tap-hint { font-size: 12px; color: #1976d2; margin-top: 8px; font-weight: 600; }
    .bg-green  { background:#d4edda; border-color:#c3e6cb; }
    .bg-yellow { background:#fff3cd; border-color:#ffeaa7; }
    .bg-blue   { background:#d1ecf1; border-color:#bee5eb; }
    .bg-purple { background:#e8daef; border-color:#d7bce8; }

    .hero { background:linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); color:white; padding:30px; border-radius:8px; margin-bottom:20px; }
    .hero h1 { color:white; margin:0 0 10px 0; }
    .hero p { color:#b0c4de; margin:0; }
    .stats-bar { background:#e8eaf6; padding:12px; border-radius:8px; margin-top:12px; text-align:center; }
    .region-warning { background:#fff3cd; padding:10px; border-radius:6px; border-left:4px solid #ff9800; margin-bottom:10px; font-size:14px; }
    div[data-testid="stDownloadButton"] button { font-weight: 600; }

    .onboard {
        background: linear-gradient(135deg, #4CAF50 0%, #2E7D32 100%);
        color: white; padding: 24px; border-radius: 12px;
        margin-bottom: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    .onboard h2 { color: white; margin: 0 0 12px 0; font-size: 22px; }
    .onboard p  { color: #E8F5E9; font-size: 16px; line-height: 1.9; margin: 0; }
    .onboard a  { color: #FFEB3B; font-weight: bold; text-decoration: underline; }

    .nokey-banner {
        background: linear-gradient(135deg, #FF9800 0%, #F57C00 100%);
        color: white; padding: 18px; border-radius: 10px; margin-bottom: 16px;
    }
    .nokey-banner h3 { color: white; margin: 0 0 8px 0; }
    .nokey-banner p  { color: #FFF3E0; margin: 0; line-height: 1.7; }

    /* Sidebar arrow visible on scroll */
    section[data-testid="stSidebar"] button[kind="header"] {
        position: fixed !important; top: 10px !important; z-index: 999999 !important;
    }
    [data-testid="collapsedControl"] {
        position: fixed !important; top: 10px !important; left: 10px !important;
        z-index: 999999 !important; background: rgba(255,255,255,0.95) !important;
        border-radius: 6px !important; padding: 4px !important;
    }
</style>
""", unsafe_allow_html=True)



# ============================================================
# PWA: Make the app installable on iPhone/Android home screen
# ============================================================
st.markdown("""
<link rel="manifest" href="./app/static/manifest.json">
<meta name="theme-color" content="#4CAF50">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="CV Builder">
<meta name="application-name" content="CV Builder">
<link rel="apple-touch-icon" href="./app/static/icon-192.png">
<link rel="icon" type="image/png" sizes="192x192" href="./app/static/icon-192.png">
<link rel="icon" type="image/png" sizes="512x512" href="./app/static/icon-512.png">
""", unsafe_allow_html=True)


# ============================================================
# MOBILE: aggressive auto-close sidebar after tab pick (Issue #3)
# Polls every 500ms; on radio change, retries close up to 12 times
# ============================================================
components.html("""
<script>
(function() {
    const W = window.parent;
    const D = W.document;
    function isMobile() { return W.innerWidth < 992; }
    function findCloseBtn() {
        return D.querySelector('section[data-testid="stSidebar"] button[kind="header"]')
            || D.querySelector('[data-testid="stSidebarCollapseButton"]')
            || D.querySelector('[data-testid="collapsedControl"]');
    }
    function isSidebarOpen() {
        const sb = D.querySelector('section[data-testid="stSidebar"]');
        if (!sb) return false;
        const rect = sb.getBoundingClientRect();
        return rect.width > 50;
    }
    function tryClose() {
        if (!isMobile()) return;
        let tries = 0;
        const iv = setInterval(() => {
            if (!isSidebarOpen() || tries > 12) { clearInterval(iv); return; }
            const btn = findCloseBtn();
            if (btn) btn.click();
            tries++;
        }, 80);
    }
    function attach() {
        const radios = D.querySelectorAll('section[data-testid="stSidebar"] input[type="radio"]');
        radios.forEach(r => {
            if (r._autoCloseAttached) return;
            r._autoCloseAttached = true;
            r.addEventListener('change', () => setTimeout(tryClose, 150));
            r.addEventListener('click', () => setTimeout(tryClose, 200));
        });
        // Also listen to label clicks (mobile taps often hit label, not input)
        const labels = D.querySelectorAll('section[data-testid="stSidebar"] label');
        labels.forEach(l => {
            if (l._autoCloseAttached) return;
            l._autoCloseAttached = true;
            l.addEventListener('click', () => setTimeout(tryClose, 200));
        });
    }
    attach();
    setInterval(attach, 1000);
    const obs = new MutationObserver(attach);
    obs.observe(D.body, {childList: true, subtree: true});
})();
</script>
""", height=0)

# ============================================================
# NAVIGATION STATE — supports query-param navigation (Issue #2)
# ============================================================
PAGES = ["🏠 Home", "📝 Generate CV", "🔍 CV vs JD", "📊 CV Analysis",
         "📑 Multi-JD Compare", "🎤 Interview Prep", "🎙️ Mock Interview",
         "🧑‍💼 Coaching", "📚 My Library", "⚙️ Settings"]

# Slug map for query-param navigation from clickable cards
PAGE_SLUGS = {
    "home": "🏠 Home", "generate": "📝 Generate CV", "compare": "🔍 CV vs JD",
    "analysis": "📊 CV Analysis", "multi": "📑 Multi-JD Compare",
    "prep": "🎤 Interview Prep", "mock": "🎙️ Mock Interview",
    "coach": "🧑‍💼 Coaching", "library": "📚 My Library", "settings": "⚙️ Settings",
}

# Initialise page state
if "page" not in st.session_state:
    st.session_state.page = "🏠 Home"

# Handle direct-link navigation from clickable cards (?nav=mock etc.)
qp = st.query_params
if "nav" in qp:
    target_slug = qp.get("nav")
    if target_slug in PAGE_SLUGS:
        st.session_state.page = PAGE_SLUGS[target_slug]
    st.query_params.clear()
    st.rerun()

# ============================================================
# SIDEBAR — now with Claude provider (Bonus)
# ============================================================
with st.sidebar:
    st.markdown("# 📄 AI CV Builder")
    st.markdown("---")
    idx = PAGES.index(st.session_state.page) if st.session_state.page in PAGES else 0
    selected = st.radio("**Navigation**", PAGES, index=idx, label_visibility="collapsed",
                        key="nav_radio")
    if selected != st.session_state.page:
        st.session_state.page = selected
        st.rerun()
    page = st.session_state.page

    st.markdown("---")
    st.markdown("**🤖 AI Provider**")
    provider = st.radio("Provider",
                        ["Gemini (FREE)", "OpenAI (Paid)", "Claude (Paid)"],
                        label_visibility="collapsed", key="prov_radio")
    if "OpenAI" in provider:
        prov = "openai"; models = OPENAI_MODELS
    elif "Claude" in provider:
        prov = "claude"; models = CLAUDE_MODELS
    else:
        prov = "gemini"; models = GEMINI_MODELS

    model_display = st.selectbox("Model", [m[0] + " (" + m[1] + ")" for m in models],
                                 key="model_sel")
    model_id = model_display.split(" (")[0]
    api_key = st.text_input("🔑 API Key", type="password",
                            placeholder="Paste your API key", key="api_key_in")

    if prov == "gemini":
        st.markdown("🔗 [Get FREE Gemini key](https://aistudio.google.com/apikey)")
    elif prov == "openai":
        st.markdown("🔗 [Get OpenAI key](https://platform.openai.com/api-keys)")
    else:
        st.markdown("🔗 [Get Claude key](https://console.anthropic.com/settings/keys)")

    st.markdown("---")
    st.markdown("**🌍 Target Region** ⭐")
    region = st.selectbox("Region", RL, label_visibility="collapsed", key="region_select",
                          help="Choose your target country - CV format adapts (photo/no photo, sections, language style)")
    region_name = region.split(" ", 1)[1] if " " in region else region
    st.caption("📍 **" + region_name + "** format will be used")
    st.caption("SSL: " + SSL_M)

# ============================================================
# HELPERS
# ============================================================
def call_ai(prompt, temperature=0.7):
    """Run AI call with helpful no-API-key reminder (Issue #4)."""
    if not api_key:
        st.markdown("""
        <div class="nokey-banner">
            <h3>🔑 API key missing</h3>
            <p>You haven't entered your API key in the left sidebar yet.<br>
            <b>👉 What to do:</b><br>
            1. Open the sidebar (tap the <b>»</b> arrow top-left on phone).<br>
            2. Choose a provider — <b>Gemini (FREE)</b> is recommended for first-timers.<br>
            3. Paste your API key in the 🔑 box.<br>
            4. Don't have one? <a href="?nav=home" target="_self" style="color:#FFEB3B">
            Go to the Home page</a> for step-by-step setup instructions, or use this direct link:
            <a href="https://aistudio.google.com/apikey" target="_blank" style="color:#FFEB3B">
            aistudio.google.com/apikey</a> (free, 30 seconds).</p>
        </div>
        """, unsafe_allow_html=True)
        return None
    with st.spinner("🤖 AI is thinking..."):
        return ai_call(prompt, api_key, prov, model_id, temperature=temperature)

def download_buttons(text, region, suffix="Document", name_override=None):
    if not text: return
    name = name_override or extract_name(text)
    base_name = f"{name}_{suffix}"
    col1, col2, col3 = st.columns(3)
    with col1:
        st.download_button("💾 Save (.txt)", strip_emojis(text), base_name + ".txt",
                           mime="text/plain", use_container_width=True,
                           key="dl_txt_" + base_name)
    with col2:
        if HAS_DOCX:
            docx_data = make_docx(text, region)
            if docx_data:
                st.download_button("📄 Word (.docx)", docx_data, base_name + ".docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True, key="dl_docx_" + base_name)
        else:
            st.button("📄 Word (need python-docx)", disabled=True,
                      use_container_width=True, key="dl_docx_dis_" + base_name)
    with col3:
        if HAS_RL:
            pdf_data = make_pdf(text, region)
            if pdf_data:
                st.download_button("📑 PDF (.pdf)", pdf_data, base_name + ".pdf",
                    mime="application/pdf", use_container_width=True,
                    key="dl_pdf_" + base_name)
        else:
            st.button("📑 PDF (need reportlab)", disabled=True,
                      use_container_width=True, key="dl_pdf_dis_" + base_name)

def region_note(stored_region):
    if stored_region and stored_region != region:
        st.markdown('<div class="region-warning">⚠️ Region changed from <b>' + stored_region +
                    '</b> to <b>' + region + '</b>. Click Generate/Improve again to reformat for the new region.</div>',
                    unsafe_allow_html=True)

# ============================================================
# PAGE: HOME (clickable cards via HTML link, no extra button - Issue #2)
# ============================================================
if page == "🏠 Home":
    st.markdown(
        '<div class="hero"><h1>📄 AI CV Builder</h1>'
        '<p>Gemini (FREE) + OpenAI + Claude • Region-aware • PDF/Word Export • 130+ Mock Interview Domains<br>'
        'Anti-hallucination: AI only uses YOUR data.</p></div>',
        unsafe_allow_html=True)

    if not api_key:
        st.markdown("""
        <div class="onboard">
            <h2>🆓 First time? Get a FREE Gemini API key in 30 seconds:</h2>
            <p>
                <b>Step 1:</b> Click → <a href="https://aistudio.google.com/apikey" target="_blank">https://aistudio.google.com/apikey</a><br>
                <b>Step 2:</b> Sign in with your Google account<br>
                <b>Step 3:</b> Click 'Create API Key' → copy it → paste in sidebar<br>
                <br>
                <i>It's 100% free with generous daily quota.</i>
            </p>
        </div>
        """, unsafe_allow_html=True)

    # Each card is now a clickable LINK (anchor) — tap anywhere on it
    # uses ?nav=<slug> which the router at the top picks up
    cards = [
        ("📝 Generate CV",     "Create a region-formatted CV from any JD.",         "bg-green",  "generate"),
        ("🔍 CV vs JD",        "Compare CV vs JD. Keywords, gaps, cover letter.",   "bg-yellow", "compare"),
        ("📊 CV Analysis",     "AI scores ATS compliance. Improve & export.",       "bg-blue",   "analysis"),
        ("📑 Multi-JD Compare","Compare your CV against multiple JDs at once.",     "bg-purple", "multi"),
        ("🎤 Interview Prep",  "15-20 questions + STAR frameworks from any JD.",    "bg-green",  "prep"),
        ("🎙️ Mock Interview", "Practice with AI across 130+ domains.",              "bg-yellow", "mock"),
        ("🧑‍💼 Coaching",     "Career advice on 10+ life and career topics.",       "bg-blue",   "coach"),
        ("📚 My Library",      "Save question sets & feedback. Export as JSON.",    "bg-green",  "library"),
        ("⚙️ Settings",        "Providers, regions, mobile tips, library status.",  "bg-purple", "settings"),
    ]
    cols = st.columns(2, gap="medium")
    for i, (title, desc, css_class, slug) in enumerate(cards):
        with cols[i % 2]:
            st.markdown(
                f'<a href="?nav={slug}" target="_self" style="text-decoration:none;color:inherit;">'
                f'<div class="home-card {css_class}">'
                f'<h3>{title}</h3><p>{desc}</p>'
                f'<div class="tap-hint">👉 Tap to open</div>'
                f'</div></a>',
                unsafe_allow_html=True)

    st.markdown(
        '<div class="stats-bar">🌍 <b>14 Regions</b> &nbsp;•&nbsp; 🤖 <b>Gemini + OpenAI + Claude</b> '
        '&nbsp;•&nbsp; 📑 <b>PDF / Word Export</b> &nbsp;•&nbsp; 🎙️ <b>130+ Mock Domains</b> '
        '&nbsp;•&nbsp; 📱 <b>Mobile-friendly</b></div>', unsafe_allow_html=True)
    st.markdown("---")
    st.info("💡 **Currently set to:** " + region + " — your CV will be formatted for this region. Change in the sidebar anytime.")

# ============================================================
# PAGE: GENERATE CV
# ============================================================
elif page == "📝 Generate CV":
    st.title("📝 Generate CV")
    st.caption("Currently generating for: **" + region + "** - " + RD_[region]["photo"])
    region_note(st.session_state.get("gen_region"))
    jd_gen = st.text_area("📋 Job Description", height=200, key="gen_jd")
    info_gen = st.text_area("👤 Your Information", height=400, key="gen_info",
                            value=st.session_state.get("gen_info_val", INFO_T))
    if st.button("✨ Generate CV", type="primary", use_container_width=True):
        if not jd_gen.strip(): st.warning("Please paste a Job Description.")
        elif not info_gen.strip(): st.warning("Please fill in your information.")
        else:
            result = call_ai(p_gen(jd_gen, info_gen, region, RD_[region]))
            if result:
                st.session_state["gen_result"] = result
                st.session_state["gen_region"] = region
                st.session_state["gen_info_val"] = info_gen
    if "gen_result" in st.session_state:
        st.markdown("### ✨ Generated CV (formatted for " +
                    st.session_state.get("gen_region", region) + ")")
        st.text_area("Result", st.session_state["gen_result"], height=500,
                     key="gen_out", label_visibility="collapsed")
        st.markdown("**Download:**")
        download_buttons(st.session_state["gen_result"],
                         st.session_state.get("gen_region", region), "CV")

# ============================================================
# PAGE: CV vs JD
# ============================================================
elif page == "🔍 CV vs JD":
    st.title("🔍 CV vs Job Description")
    st.caption("Currently using region: **" + region + "**")
    region_note(st.session_state.get("cmp_region"))
    col1, col2 = st.columns(2, gap="medium")
    with col1:
        st.markdown("### 📄 Your CV")
        cv_file = st.file_uploader("Upload CV", type=["pdf", "docx", "txt"], key="cmp_cv_file")
        cv_text_input = st.text_area("Or paste CV", height=300, key="cmp_cv_input")
        cv_text = read_file(cv_file) if cv_file else cv_text_input
    with col2:
        st.markdown("### 📋 Job Description")
        jd_cmp = st.text_area("Paste JD", height=300, key="cmp_jd", label_visibility="collapsed")
    if st.button("🔍 Compare CV vs JD", type="primary", use_container_width=True):
        if not cv_text.strip() or not jd_cmp.strip():
            st.warning("Please provide both CV and JD.")
        else:
            result = call_ai(p_cmp(cv_text, jd_cmp, region, RD_[region]))
            if result:
                st.session_state["cmp_result"] = result
                st.session_state["cmp_cv"]     = cv_text
                st.session_state["cmp_jd"]     = jd_cmp
                st.session_state["cmp_region"] = region
    if "cmp_result" in st.session_state:
        st.markdown("### 📊 Comparison Result")
        st.text_area("Result", st.session_state["cmp_result"], height=400,
                     key="cmp_out", label_visibility="collapsed")
        cv_name = extract_name(st.session_state.get("cmp_cv", ""))
        download_buttons(st.session_state["cmp_result"], region,
                         "JD_Comparison", name_override=cv_name)

        st.markdown("---")
        st.markdown("### ✨ Improve CV with Keywords")
        col_k1, col_k2 = st.columns(2)
        with col_k1:
            kw_strength = st.radio("Keyword Strength", ["Experience-based", "Force ALL"],
                                   key="kw_str", horizontal=True)
        with col_k2:
            kw_mode = st.radio("Keywords from", ["Auto from JD", "Manual"],
                               key="kw_mode", horizontal=True)
        manual_kw = ""
        if kw_mode == "Manual":
            manual_kw = st.text_input("Enter keywords (comma-separated)",
                                      placeholder="Docker, Kubernetes, CI/CD")
        if st.button("✨ Improve CV", type="primary", use_container_width=True):
            strength = "force_all" if kw_strength == "Force ALL" else "experience"
            result = call_ai(p_imp_cmp(st.session_state["cmp_cv"], st.session_state["cmp_result"],
                                       st.session_state["cmp_jd"], region, RD_[region],
                                       manual_kw, strength))
            if result:
                st.session_state["cmp_improved"] = result
                st.session_state["cmp_region"]   = region
        if "cmp_improved" in st.session_state:
            st.markdown("#### ✨ Improved CV (formatted for " +
                        st.session_state.get("cmp_region", region) + ")")
            st.text_area("Improved", st.session_state["cmp_improved"], height=500,
                         key="cmp_imp_out", label_visibility="collapsed")
            st.markdown("**Download Improved CV:**")
            download_buttons(st.session_state["cmp_improved"],
                             st.session_state.get("cmp_region", region),
                             "Improved_CV")

        st.markdown("---")
        st.markdown("### 📨 Cover Letter")
        if st.button("📨 Generate Cover Letter", type="primary", use_container_width=True):
            cv_src = st.session_state.get("cmp_improved", st.session_state.get("cmp_cv", ""))
            result = call_ai(p_cover(cv_src, st.session_state["cmp_jd"], region, RD_[region]))
            if result: st.session_state["cmp_cl"] = result
        if "cmp_cl" in st.session_state:
            st.text_area("Cover Letter", st.session_state["cmp_cl"], height=400,
                         key="cmp_cl_out", label_visibility="collapsed")
            st.markdown("**Download Cover Letter:**")
            cv_name = extract_name(st.session_state.get("cmp_improved",
                                   st.session_state.get("cmp_cv", st.session_state["cmp_cl"])))
            download_buttons(st.session_state["cmp_cl"], region,
                             "Cover_Letter", name_override=cv_name)

# ============================================================
# PAGE: CV ANALYSIS
# ============================================================
elif page == "📊 CV Analysis":
    st.title("📊 CV Analysis")
    st.caption("Analysing against **" + region + "** standards.")
    region_note(st.session_state.get("ana_region"))
    cv_ana_file = st.file_uploader("Upload CV", type=["pdf", "docx", "txt"], key="ana_file")
    cv_ana_input = st.text_area("Or paste CV", height=300, key="ana_cv_input")
    cv_ana = read_file(cv_ana_file) if cv_ana_file else cv_ana_input
    st.info("**Score guide:** 90-100 Excellent | 75-89 Good | 60-74 Average | <60 Needs work")
    if st.button("📊 Analyse CV", type="primary", use_container_width=True):
        if not cv_ana.strip(): st.warning("Please provide a CV.")
        else:
            result = call_ai(p_ana(cv_ana, region, RD_[region]))
            if result:
                st.session_state["ana_result"] = result
                st.session_state["ana_cv"]     = cv_ana
                st.session_state["ana_region"] = region
    if "ana_result" in st.session_state:
        st.markdown("### 📊 Analysis")
        st.text_area("Analysis", st.session_state["ana_result"], height=400,
                     key="ana_out", label_visibility="collapsed")
        cv_name = extract_name(st.session_state.get("ana_cv", ""))
        download_buttons(st.session_state["ana_result"], region,
                         "CV_Analysis", name_override=cv_name)
        st.markdown("---")
        if st.button("✨ Improve CV", type="primary", use_container_width=True):
            result = call_ai(p_imp(st.session_state["ana_cv"], st.session_state["ana_result"],
                                   region, RD_[region]))
            if result:
                st.session_state["ana_improved"] = result
                st.session_state["ana_region"]   = region
        if "ana_improved" in st.session_state:
            st.markdown("### ✨ Improved CV (formatted for " +
                        st.session_state.get("ana_region", region) + ")")
            st.text_area("Improved", st.session_state["ana_improved"], height=500,
                         key="ana_imp_out", label_visibility="collapsed")
            st.markdown("**Download:**")
            download_buttons(st.session_state["ana_improved"],
                             st.session_state.get("ana_region", region),
                             "Improved_CV")

# ============================================================
# PAGE: MULTI-JD COMPARE
# ============================================================
elif page == "📑 Multi-JD Compare":
    st.title("📑 Multi-JD Compare")
    st.caption("Using region: **" + region + "**")
    cv_multi_file = st.file_uploader("Upload CV", type=["pdf", "docx", "txt"], key="multi_file")
    cv_multi_input = st.text_area("Or paste CV", height=200, key="multi_cv_input")
    cv_multi = read_file(cv_multi_file) if cv_multi_file else cv_multi_input
    n_jds = st.number_input("Number of JDs to compare", min_value=2, max_value=10,
                            value=2, key="multi_n")
    jd_texts = []
    for i in range(int(n_jds)):
        jd_texts.append(st.text_area("📋 JD " + str(i + 1), height=150,
                                     key="multi_jd_" + str(i)))
    if st.button("📑 Compare All JDs", type="primary", use_container_width=True):
        if not cv_multi.strip():
            st.warning("Please provide a CV.")
        else:
            combined = "\n\n".join(
                ["--- JD " + str(i + 1) + " ---\n" + t
                 for i, t in enumerate(jd_texts) if t.strip()])
            if not combined.strip():
                st.warning("Please fill in at least one JD.")
            else:
                result = call_ai(p_multi(cv_multi, combined, region, RD_[region]))
                if result:
                    st.session_state["multi_result"] = result
                    st.session_state["multi_cv"]     = cv_multi
    if "multi_result" in st.session_state:
        st.markdown("### 📊 Comparison Results")
        st.text_area("Results", st.session_state["multi_result"], height=500,
                     key="multi_out", label_visibility="collapsed")
        st.markdown("**Download:**")
        cv_name = extract_name(st.session_state.get("multi_cv", ""))
        download_buttons(st.session_state["multi_result"], region,
                         "Multi_JD_Results", name_override=cv_name)

# ============================================================
# PAGE: INTERVIEW PREP
# ============================================================
elif page == "🎤 Interview Prep":
    st.title("🎤 Interview Preparation")
    st.info("💡 Questions come from AI knowledge of interview patterns (no web search).")
    jd_intv = st.text_area("📋 Job Description", height=250, key="intv_jd")
    if st.button("🎯 Generate Questions", type="primary", use_container_width=True):
        if not jd_intv.strip():
            st.warning("Please paste a JD.")
        else:
            result = call_ai(p_int(jd_intv))
            if result:
                st.session_state["intv_result"] = result
    if "intv_result" in st.session_state:
        st.markdown("### 🎯 Questions + Answer Frameworks")
        st.text_area("Questions", st.session_state["intv_result"], height=600,
                     key="intv_out", label_visibility="collapsed")
        st.markdown("**Download:**")
        download_buttons(st.session_state["intv_result"], region,
                         "Interview_Questions", name_override="Candidate")

# ============================================================
# PAGE: MOCK INTERVIEW — persistent per-domain history (Issue #1)
# ============================================================
elif page == "🎙️ Mock Interview":
    st.title("🎙️ Mock Interview")
    st.caption("AI generates fresh questions every time. Switch domain, level, or click Regenerate "
               "and the AI is given an explicit anti-repeat list.")

    # ---- Per-domain history of previously generated questions ----
    if "mock_history" not in st.session_state:
        st.session_state["mock_history"] = {}   # key -> list[str questions]

    with st.container(border=True):
        st.markdown("### 📋 Question Source")
        source = st.radio("Source", ["📋 From Job Description", "🎯 By Domain / Field"],
                          horizontal=True, key="mock_src", label_visibility="collapsed")

        mock_jd = ""
        category = None
        mock_domain = None

        if "Job Description" in source:
            mock_jd = st.text_area("Paste JD", height=200, key="mock_jd")
        else:
            cat_col, dom_col = st.columns([1, 2])
            with cat_col:
                category = st.selectbox("Category", list(DOMAIN_GROUPS.keys()), key="mock_cat")
            with dom_col:
                mock_domain = st.selectbox("Domain / Field", DOMAIN_GROUPS[category], key="mock_dom")
            st.caption(f"📚 **{len(DOMAIN_GROUPS[category])}** option(s) under **{category}** "
                       f"• Total: **{sum(len(v) for v in DOMAIN_GROUPS.values())}**")

        exp_level = st.selectbox("🎯 Experience Level", EXP_LEVELS, index=2, key="mock_exp",
                                 help="AI tailors difficulty to your level. Students get fundamentals, "
                                      "Seniors get architecture & leadership scenarios.")

        n_q = st.radio("Number of questions", ["10", "15", "20"], index=1,
                       horizontal=True, key="mock_n")

    # Build per-domain key for history tracking
    history_key = (
        f"JD::{exp_level}::{(mock_jd[:80] or 'na')}" if "Job Description" in source
        else f"DOM::{category}::{mock_domain}::{exp_level}"
    )

    # Hide stale on settings change
    sig = f"{source}|{category}|{mock_domain}|{exp_level}|{n_q}|{mock_jd[:100]}"
    if st.session_state.get("mock_sig") and st.session_state["mock_sig"] != sig and \
       "mock_questions" in st.session_state:
        st.warning("⚠️ Settings changed — old questions hidden. Click **Generate Questions** "
                   "to refresh for the new selection.")
        st.session_state.pop("mock_questions", None)

    col_g1, col_g2 = st.columns([3, 1])
    with col_g1:
        gen_clicked = st.button("🎯 Generate Questions", type="primary", use_container_width=True)
    with col_g2:
        regen_clicked = st.button("🔄 Regenerate (different)", use_container_width=True,
                                  help="Force a brand new set; anti-repeat list grows with each click")

    if gen_clicked or regen_clicked:
        # Always pass history — even on first Generate — to push AI away from any cached patterns
        history = st.session_state["mock_history"].get(history_key, [])

        if "Job Description" in source:
            if not mock_jd.strip():
                st.warning("Please paste a JD.")
            else:
                result = call_ai(
                    p_mock_jd(mock_jd, int(n_q), exp_level, history_questions=history),
                    temperature=1.0)
                if result:
                    st.session_state["mock_questions"] = result
                    st.session_state["mock_sig"] = sig
                    st.session_state["mock_meta"] = {
                        "source": "JD", "domain": "From JD",
                        "level": exp_level, "n": n_q,
                        "ts": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    }
                    # Append parsed questions to history for next round
                    new_qs = _parse_questions(result)
                    if new_qs:
                        st.session_state["mock_history"].setdefault(history_key, []).extend(new_qs)
        else:
            result = call_ai(
                p_mock_dom(mock_domain, int(n_q), exp_level, history_questions=history),
                temperature=1.0)
            if result:
                st.session_state["mock_questions"] = result
                st.session_state["mock_sig"] = sig
                st.session_state["mock_meta"] = {
                    "source": "Domain", "domain": f"{category} → {mock_domain}",
                    "level": exp_level, "n": n_q,
                    "ts": datetime.now().strftime("%Y-%m-%d %H:%M"),
                }
                new_qs = _parse_questions(result)
                if new_qs:
                    st.session_state["mock_history"].setdefault(history_key, []).extend(new_qs)

    if "mock_questions" in st.session_state:
        meta = st.session_state.get("mock_meta", {})
        hist_count = len(st.session_state["mock_history"].get(history_key, []))
        st.markdown(f"### 📋 Generated Questions "
                    f"<span style='font-size:14px;color:#888'>"
                    f"({meta.get('domain','')} • {meta.get('level','')} • {meta.get('ts','')} "
                    f"• history: {hist_count} Qs tracked)</span>",
                    unsafe_allow_html=True)
        st.text_area("Questions", st.session_state["mock_questions"], height=350,
                     key="mock_q_out", label_visibility="collapsed")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("💾 Save this set to My Library", use_container_width=True,
                         key="lib_save_q"):
                lib = st.session_state.setdefault("library", [])
                lib.append({
                    "type": "questions",
                    "meta": meta,
                    "content": st.session_state["mock_questions"],
                })
                st.success(f"✅ Saved! Library has {len(lib)} item(s).")
        with c2:
            if st.button("🧹 Clear history for this domain", use_container_width=True,
                         key="clear_hist", help="Reset anti-repeat tracking for this domain/level"):
                st.session_state["mock_history"].pop(history_key, None)
                st.success("✅ Anti-repeat history cleared for this selection.")

        st.markdown("---")
        st.markdown("### 📝 Practice & Evaluate")
        mock_q = st.text_area("Question", height=80, key="mock_q_in")
        mock_a = st.text_area("Your Answer", height=200, key="mock_a_in")
        if st.button("📊 Evaluate My Answer", type="primary", use_container_width=True):
            if not mock_q.strip() or not mock_a.strip():
                st.warning("Please provide both question and answer.")
            else:
                result = call_ai(p_mock_eval(mock_q, mock_a, exp_level))
                if result:
                    st.session_state["mock_feedback"] = result
                    st.session_state["mock_feedback_q"] = mock_q

        if "mock_feedback" in st.session_state:
            st.markdown("### 📊 AI Feedback")
            st.text_area("Feedback", st.session_state["mock_feedback"], height=400,
                         key="mock_fb_out", label_visibility="collapsed")
            st.markdown("**Download:**")
            download_buttons(st.session_state["mock_feedback"], region,
                             "Mock_Feedback", name_override="Candidate")

            if st.button("💾 Save this feedback to My Library", use_container_width=True,
                         key="lib_save_fb"):
                lib = st.session_state.setdefault("library", [])
                lib.append({
                    "type": "feedback",
                    "meta": {**(st.session_state.get("mock_meta") or {}),
                             "question": st.session_state.get("mock_feedback_q", "")},
                    "content": st.session_state["mock_feedback"],
                })
                st.success(f"✅ Saved! Library has {len(lib)} item(s).")

# ============================================================
# PAGE: COACHING
# ============================================================
elif page == "🧑‍💼 Coaching":
    st.title("🧑‍💼 Career Coaching")
    topic = st.selectbox("Topic",
        ["Career Change", "Salary Negotiation", "Skill Development", "Job Search",
         "Interview Confidence", "LinkedIn", "Networking", "Promotion",
         "Work-Life Balance", "Remote Work"], key="coach_topic")
    context = st.text_area("Your situation/context", height=200, key="coach_ctx",
                           placeholder="Describe where you are, what you want, and challenges...")
    if st.button("💡 Get Advice", type="primary", use_container_width=True):
        if not context.strip():
            st.warning("Please describe your situation.")
        else:
            result = call_ai(p_coach(topic, context))
            if result:
                st.session_state["coach_result"] = result
    if "coach_result" in st.session_state:
        st.markdown("### 💡 Personalised Advice")
        st.text_area("Advice", st.session_state["coach_result"], height=500,
                     key="coach_out", label_visibility="collapsed")
        st.markdown("**Download:**")
        download_buttons(st.session_state["coach_result"], region,
                         "Career_Advice", name_override="Candidate")

# ============================================================
# PAGE: MY LIBRARY
# ============================================================
elif page == "📚 My Library":
    st.title("📚 My Library")
    st.caption("Save your mock interview questions & feedback. Export as JSON to keep permanently.")

    st.info("💡 **How it works:** Items are saved during your current session. "
            "Before closing the app, click **⬇️ Export Library** to download a backup file. "
            "Next time, click **⬆️ Import** to restore your saved items.")

    lib = st.session_state.setdefault("library", [])
    st.success(f"📚 Your library currently has **{len(lib)}** saved item(s).")

    col1, col2 = st.columns(2)
    with col1:
        if lib:
            st.download_button("⬇️ Export Library (JSON backup)",
                               data=json.dumps(lib, indent=2, ensure_ascii=False),
                               file_name=f"my_library_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                               mime="application/json",
                               use_container_width=True)
        else:
            st.button("⬇️ Export (nothing to export yet)", disabled=True, use_container_width=True)
    with col2:
        if lib and st.button("🗑️ Clear ALL items", use_container_width=True):
            st.session_state["library"] = []
            st.rerun()

    st.markdown("**⬆️ Import previously exported library (restore from backup):**")
    up = st.file_uploader("Upload library JSON", type=["json"], key="lib_import")
    if up is not None:
        try:
            data = json.loads(up.read().decode("utf-8"))
            if isinstance(data, list):
                st.session_state["library"] = data + lib
                st.success(f"✅ Imported {len(data)} item(s). Total now: {len(st.session_state['library'])}")
                st.rerun()
            else:
                st.error("Bad file: expected a JSON list.")
        except Exception as e:
            st.error(f"Bad file: {e}")

    st.markdown("---")

    if not lib:
        st.info("📭 Library is empty. Go to **🎙️ Mock Interview**, generate questions or get feedback, "
                "then click **💾 Save to My Library**.")
    else:
        st.markdown("### 📑 Your saved items")
        for i, item in enumerate(reversed(lib)):
            real_idx = len(lib) - 1 - i
            meta = item.get("meta", {})
            with st.expander(f"#{real_idx+1} • {item.get('type','?').title()} • "
                             f"{meta.get('domain','?')} • {meta.get('level','?')} • "
                             f"{meta.get('ts','?')}"):
                st.text_area("Content", item.get("content", ""), height=300,
                             key=f"lib_view_{real_idx}", label_visibility="collapsed")
                cdl1, cdl2 = st.columns(2)
                with cdl1:
                    st.download_button("💾 Download .txt", item.get("content", ""),
                                       file_name=f"library_item_{real_idx+1}.txt",
                                       mime="text/plain", use_container_width=True,
                                       key=f"lib_dl_{real_idx}")
                with cdl2:
                    if st.button("🗑️ Delete", use_container_width=True, key=f"lib_del_{real_idx}"):
                        lib.pop(real_idx)
                        st.rerun()

# ============================================================
# PAGE: SETTINGS
# ============================================================
elif page == "⚙️ Settings":
    st.title("⚙️ Settings & Info")
    with st.container(border=True):
        st.markdown("### 🤖 AI Providers")
        st.markdown("- **Gemini (FREE):** https://aistudio.google.com/apikey")
        st.markdown("- **OpenAI (Paid):** https://platform.openai.com/api-keys")
        st.markdown("- **Claude / Anthropic (Paid):** https://console.anthropic.com/settings/keys")
    with st.container(border=True):
        st.markdown("### 🌍 Region-Aware CV Generation")
        st.markdown("CV format adapts based on selected region:")
        st.markdown("- **UK / US / Canada / Australia:** NO photo, achievement-focused")
        st.markdown("- **Germany / Japan / China / Spain / Italy / UAE:** Photo placeholder included")
        st.markdown("- **India:** Declaration statement at bottom")
        st.markdown("- **Italy:** GDPR privacy clause")
        st.markdown("- **France / EU:** CEFR language levels")
        st.markdown("- **UAE / Gulf:** Nationality + visa status mandatory")
    with st.container(border=True):
        st.markdown("### 💾 Three Download Formats")
        st.markdown("- **💾 Save (.txt):** Plain text")
        st.markdown("- **📄 Word (.docx):** Region-specific styling, photo placeholder if needed")
        st.markdown("- **📑 PDF (.pdf):** Professional PDF, photo box for relevant regions")
        st.markdown("Files are auto-named using the candidate's name (e.g. *Vitthal_Kallappa_CV.pdf*).")
    with st.container(border=True):
        st.markdown("### 🎙️ Mock Interview Coverage")
        st.markdown(f"**{sum(len(v) for v in DOMAIN_GROUPS.values())}** "
                    f"domains across **{len(DOMAIN_GROUPS)}** categories.")
        st.markdown("Engineering now includes **CSE, ISE, IT, ECE, EEE, Biomedical, Robotics, Mining, "
                    "Petroleum, Marine, Textile, Metallurgy, Biotech, Architecture** and more.")
        st.markdown("Medical includes **AYUSH (BAMS/BHMS/BUMS)**, Yoga, Nutrition, OT.")
        st.markdown("Indian Education: **PUC, Diploma, ITI, BCA, BBA, B.A/B.Com/B.Sc, B.Tech, MBA, "
                    "M.Sc/M.A/M.Com, M.Tech, Ph.D**.")
        st.markdown("AI tailors questions by **experience level**: Student/Fresher, Junior, Mid, "
                    "Senior, Lead/Principal.")
        st.markdown("**Anti-repeat tracking:** Every generated question is remembered per "
                    "domain+level and explicitly excluded from future rounds. "
                    "Clear it any time with the 🧹 button on Mock Interview page.")
    with st.container(border=True):
        st.markdown("### 📚 My Library")
        st.markdown("Save mock interview questions & feedback during your session. "
                    "**Always export to JSON before closing** the app — items reset when you leave.")
    with st.container(border=True):
        st.markdown("### 📱 Use on Phone")
        st.markdown("**iPhone (Safari):** Share → Add to Home Screen")
        st.markdown("**Android (Chrome):** Menu → Add to Home Screen")
        st.markdown("Tip: tap any tab in the sidebar — it auto-closes on mobile.")
    docx_status = "OK" if HAS_DOCX else "Missing"
    rl_status   = "OK" if HAS_RL   else "Missing"
    pdf_status  = "OK" if HAS_PDF  else "Missing"
    req_status  = "OK" if HAS_REQ  else "Missing"
    st.caption(f"SSL: {SSL_M} | python-docx: {docx_status} | reportlab: {rl_status} "
               f"| pypdf: {pdf_status} | requests: {req_status}")
