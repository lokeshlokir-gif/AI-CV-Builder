#!/usr/bin/env python3
"""AI CV Builder - Streamlit Web App with Word/PDF export + region-aware CV"""
import streamlit as st
import json, re, time, io, ssl
import urllib.request, urllib.error
from datetime import datetime

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
FB_ = [m[0] for m in GEMINI_MODELS]
MR = 3
RDL = 8

DOMAINS = [
    "General", "AI / Machine Learning", "Data Science / Analytics", "Cyber Security",
    "Cloud / DevOps", "Software Engineering", "Web Development", "Mobile Development",
    "Database / SQL", "Networking", "Project Management", "Agile / Scrum",
    "Finance / Banking", "Law / Legal", "Healthcare / Medical", "Marketing / Digital",
    "Sales / Business Dev", "HR / Recruitment", "Education / Teaching",
    "Agriculture / Environment", "Supply Chain / Logistics", "Manufacturing / Engineering"
]
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

def _openai(pr, key, model, to=120):
    url = "https://api.openai.com/v1/chat/completions"
    pl = {"model": model, "messages": [{"role": "user", "content": pr}], "temperature": 0.7, "max_tokens": 4096}
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
        return "❌ All models failed."
    except Exception as e:
        return f"[Error] {e}"

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

def make_docx(text, region):
    if not HAS_DOCX: return None
    ri = RD_.get(region, {})
    cp = ri.get("cp", (0, 51, 102))
    has_photo = ri.get("pn", False)
    PC = RGBColor(*cp)
    
    doc = Document()
    for s in doc.sections:
        s.top_margin = Cm(1.5)
        s.bottom_margin = Cm(1.5)
        s.left_margin = Cm(2)
        s.right_margin = Cm(2)
    
    lines = [l for l in text.split("\n") if l.strip()]
    SK = ["SUMMARY", "SKILL", "EXPERIENCE", "EDUCATION", "CERTIF", "PROFILE", "OBJECTIVE", "PERSONAL", "CAREER", "LANGUAGES", "ADDITIONAL", "DECLARATION", "REFERENCES", "QUALIFICATIONS", "WORK HISTORY", "MOTIVATION"]
    
    name = ""
    cs = 0
    for idx, line in enumerate(lines):
        c = re.sub(r"#+\s*", "", line.strip()).strip("*").strip()
        if not c: continue
        u = c.upper().replace(":", "")
        if not name and not any(k in u for k in SK) and len(c) > 1 and not c.lower().startswith(("here", "below", "i have", "this is", "based on")):
            name = c
            cs = idx + 1
            break
    
    if has_photo and name:
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        lc = table.rows[0].cells[0]
        lc.width = Cm(12)
        p = lc.paragraphs[0]
        r = p.add_run(name)
        r.font.size = Pt(22)
        r.font.bold = True
        r.font.color.rgb = PC
        rc = table.rows[0].cells[1]
        rc.width = Cm(4)
        pp = rc.paragraphs[0]
        pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pp.paragraph_format.space_before = Pt(20)
        pp.paragraph_format.space_after = Pt(20)
        pp.add_run("[ PHOTO ]\n3.5 x 4.5cm").font.size = Pt(10)
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(name)
        r.font.size = Pt(24)
        r.font.bold = True
        r.font.color.rgb = PC
    
    for i in range(cs, len(lines)):
        line = lines[i].strip()
        if not line: continue
        if line.lower().startswith(("here is", "below is", "i have", "this is", "based on")):
            continue
        c = re.sub(r"#+\s*", "", line).strip()
        cfc = c.strip("*").strip()
        if not cfc: continue
        u = cfc.upper().replace(":", "")
        is_h = False
        ht = cfc
        if line.startswith("#"):
            is_h = True
        elif line.startswith("**") and line.endswith("**") and len(line) > 4:
            ht = line.strip("*").strip()
            u2 = ht.upper().replace(":", "")
            if any(k in u2 for k in SK):
                is_h = True
        elif len(cfc) >= 3 and not cfc.startswith(("-", "•", "*")):
            al = [x for x in cfc if x.isalpha()]
            if al and sum(1 for x in al if x.isupper()) / len(al) > 0.7 and any(k in u for k in SK):
                is_h = True
        
        if is_h:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run(ht.upper())
            r.font.size = Pt(12)
            r.font.bold = True
            r.font.color.rgb = PC
        elif cfc.startswith(("-", "•", "▪", "*")):
            bt = cfc.lstrip("-*•▪ ").strip()
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.8)
            p.add_run("▸ ").font.size = Pt(9)
            r = p.add_run(re.sub(r"\*\*(.+?)\*\*", r"\1", bt))
            r.font.size = Pt(10)
        else:
            p = doc.add_paragraph()
            for part in re.split(r"(\*\*.+?\*\*)", cfc):
                if part.startswith("**") and part.endswith("**"):
                    r = p.add_run(part.strip("*"))
                    r.font.size = Pt(10)
                    r.font.bold = True
                else:
                    cl2 = re.sub(r"\*(.+?)\*", r"\1", part)
                    if cl2.strip():
                        r = p.add_run(cl2)
                        r.font.size = Pt(10)
    
    if ri.get("code") == "IT":
        doc.add_paragraph()
        r = doc.add_paragraph().add_run("Autorizzo il trattamento dei dati personali ai sensi del GDPR (UE 2016/679).")
        r.font.size = Pt(8)
        r.italic = True
    if ri.get("code") == "IN":
        doc.add_paragraph()
        r = doc.add_paragraph().add_run("Declaration: I hereby declare that all information provided above is true to the best of my knowledge.")
        r.font.size = Pt(9)
        r.italic = True
    
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(20)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("AI CV Builder - " + datetime.now().strftime("%d %B %Y"))
    r.font.size = Pt(7)
    r.font.color.rgb = RGBColor(150, 150, 150)
    r.italic = True
    
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

def make_pdf(text, region):
    if not HAS_RL: return None
    ri = RD_.get(region, {})
    cp = ri.get("cp", (0, 51, 102))
    has_photo = ri.get("pn", False)
    color_hex = "#%02X%02X%02X" % cp
    
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=1.5*cm, bottomMargin=1.5*cm, leftMargin=2*cm, rightMargin=2*cm)
    ss = getSampleStyleSheet()
    ss.add(ParagraphStyle("CVName", fontSize=20, leading=24, alignment=TA_CENTER, spaceAfter=6, textColor=color_hex, fontName="Helvetica-Bold"))
    ss.add(ParagraphStyle("CVNameL", fontSize=20, leading=24, alignment=TA_LEFT, spaceAfter=6, textColor=color_hex, fontName="Helvetica-Bold"))
    ss.add(ParagraphStyle("CVHead", fontSize=12, leading=14, spaceBefore=12, spaceAfter=4, textColor=color_hex, fontName="Helvetica-Bold"))
    ss.add(ParagraphStyle("CVBody", fontSize=10, leading=13, spaceAfter=2, fontName="Helvetica"))
    ss.add(ParagraphStyle("CVBullet", fontSize=10, leading=13, leftIndent=20, spaceAfter=1, fontName="Helvetica"))
    ss.add(ParagraphStyle("CVPhoto", fontSize=9, leading=12, alignment=TA_CENTER, fontName="Helvetica"))
    
    lines = [l for l in text.split("\n") if l.strip()]
    SK = ["SUMMARY", "SKILL", "EXPERIENCE", "EDUCATION", "CERTIF", "PROFILE", "OBJECTIVE", "PERSONAL", "CAREER", "LANGUAGES", "ADDITIONAL", "DECLARATION", "REFERENCES", "QUALIFICATIONS", "WORK HISTORY", "MOTIVATION"]
    
    name = ""
    cs = 0
    for idx, line in enumerate(lines):
        c = re.sub(r"#+\s*", "", line.strip()).strip("*").strip()
        if not c: continue
        u = c.upper().replace(":", "")
        if not name and not any(k in u for k in SK) and len(c) > 1 and not c.lower().startswith(("here", "below", "i have", "this is", "based on")):
            name = c
            cs = idx + 1
            break
    
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
        story.append(pt)
        story.append(Spacer(1, 10))
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
        elif c.startswith(("-", "•", "▪", "*")):
            stripped = safe.lstrip("-*•▪ ")
            story.append(Paragraph("▸ " + stripped, ss["CVBullet"]))
        else:
            story.append(Paragraph(re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", safe), ss["CVBody"]))
    
    if ri.get("code") == "IT":
        story.append(Spacer(1, 12))
        story.append(Paragraph("<i>Autorizzo il trattamento dei dati personali ai sensi del GDPR (UE 2016/679).</i>", ss["CVBody"]))
    if ri.get("code") == "IN":
        story.append(Spacer(1, 12))
        story.append(Paragraph("<i>Declaration: I hereby declare that all information provided above is true to the best of my knowledge.</i>", ss["CVBody"]))
    
    story.append(Spacer(1, 20))
    story.append(Paragraph("<i><font size=7 color='#999999'>AI CV Builder - " + datetime.now().strftime("%d %B %Y") + "</font></i>", ss["CVBody"]))
    doc.build(story)
    buf.seek(0)
    return buf.getvalue()

NO_PRE = "RULES: 1) NO FABRICATION 2) Missing=[Not provided] 3) EXPERIENCE=sum jobs(exclude gaps) 4) Only stated numbers 5) Only stated skills 6) Name first, no preamble 7) ACTUAL years 8) Enhance wording only 9) Strictly follow region format"

def _rb(rn, ri):
    return (
        "TARGET REGION: " + rn + "\n" +
        "DOCUMENT TYPE: " + ri["term"] + "\n" +
        "LENGTH: " + ri["pages"] + "\n" +
        "PHOTO REQUIREMENT: " + ri["photo"] + "\n" +
        "PERSONAL INFO TO INCLUDE: " + ri["pi"] + "\n" +
        "REQUIRED SECTIONS: " + ri["sec"] + "\n" +
        "STYLE: " + ri["sty"] + "\n" +
        "SPECIAL REQUIREMENTS: " + ri["sp"] + "\n" +
        "ATS: " + ri["ats"] + "\n" +
        "AVOID: " + ri["av"] + "\n\n" +
        "IMPORTANT: This CV MUST be formatted specifically for " + rn + ". If region requires photo, include [PHOTO PLACEHOLDER] at top. Use exact sections listed above."
    )

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

def p_mock_jd(jd, n=15):
    return ("Interview coach.\nGenerate " + str(n) + " questions from JD.\n60% Technical + 25% Behavioural + 15% Situational" +
            "\nFormat:\nQ1: [question]\nQ2: [question]\n...Only questions.\n---JD---\n" + jd)

def p_mock_dom(dom, n=15):
    return ("Interview coach for " + dom + ".\nGenerate " + str(n) + " questions.\n60% Technical + 25% Behavioural + 15% Situational" +
            "\nBasic to advanced.\nFormat:\nQ1: [question]\nQ2: [question]\n...Only questions.")

def p_mock_eval(q, a):
    return ("Interview coach. Evaluate:\nQuestion: " + q + "\nAnswer: " + a +
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

st.set_page_config(page_title="AI CV Builder", page_icon="📄", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    .card-green { background:#d4edda; padding:20px; border-radius:8px; border:1px solid #c3e6cb; margin-bottom:12px; min-height:100px; }
    .card-yellow { background:#fff3cd; padding:20px; border-radius:8px; border:1px solid #ffeaa7; margin-bottom:12px; min-height:100px; }
    .card-blue { background:#d1ecf1; padding:20px; border-radius:8px; border:1px solid #bee5eb; margin-bottom:12px; min-height:100px; }
    .card-purple { background:#e8daef; padding:20px; border-radius:8px; border:1px solid #d7bce8; margin-bottom:12px; min-height:100px; }
    .card-green h3, .card-yellow h3, .card-blue h3, .card-purple h3 { margin-top:0; color:#1a1a2e; }
    .card-green p, .card-yellow p, .card-blue p, .card-purple p { color:#555; margin:0; }
    .hero { background:linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); color:white; padding:30px; border-radius:8px; margin-bottom:20px; }
    .hero h1 { color:white; margin:0 0 10px 0; }
    .hero p { color:#b0c4de; margin:0; }
    .stats-bar { background:#e8eaf6; padding:12px; border-radius:8px; margin-top:12px; text-align:center; }
    .region-warning { background:#fff3cd; padding:10px; border-radius:6px; border-left:4px solid #ff9800; margin-bottom:10px; font-size:14px; }
    div[data-testid="stDownloadButton"] button { font-weight: 600; }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("# 📄 AI CV Builder")
    st.markdown("---")
    page = st.radio("**Navigation**",
        ["🏠 Home", "📝 Generate CV", "🔍 CV vs JD",
         "📊 CV Analysis", "📑 Multi-JD Compare",
         "🎤 Interview Prep", "🎙️ Mock Interview",
         "🧑‍💼 Coaching", "⚙️ Settings"],
        label_visibility="collapsed")
    st.markdown("---")
    st.markdown("**🤖 AI Provider**")
    provider = st.radio("Provider", ["Gemini (FREE)", "OpenAI (Paid)"], label_visibility="collapsed")
    prov = "openai" if "OpenAI" in provider else "gemini"
    models = OPENAI_MODELS if prov == "openai" else GEMINI_MODELS
    model_display = st.selectbox("Model", [m[0] + " (" + m[1] + ")" for m in models])
    model_id = model_display.split(" (")[0]
    api_key = st.text_input("🔑 API Key", type="password", placeholder="Paste your API key")
    if prov == "gemini":
        st.markdown("🔗 [Get FREE Gemini key](https://aistudio.google.com/apikey)")
    else:
        st.markdown("🔗 [Get OpenAI key](https://platform.openai.com/api-keys)")
    st.markdown("---")
    st.markdown("**🌍 Target Region** ⭐")
    region = st.selectbox("Region", RL, label_visibility="collapsed", key="region_select",
                          help="Choose your target country - CV format adapts (photo/no photo, sections, language style)")
    region_name = region.split(" ", 1)[1] if " " in region else region
    st.caption("📍 **" + region_name + "** format will be used")
    st.caption("SSL: " + SSL_M)

def call_ai(prompt):
    if not api_key:
        st.error("⚠️ Please enter your API key in the sidebar.")
        return None
    with st.spinner("🤖 AI is thinking..."):
        return ai_call(prompt, api_key, prov, model_id)

def download_buttons(text, region, base_name="document"):
    if not text: return
    col1, col2, col3 = st.columns(3)
    with col1:
        st.download_button("💾 Save (.txt)", text, base_name + ".txt", mime="text/plain",
                           use_container_width=True, key="dl_txt_" + base_name)
    with col2:
        if HAS_DOCX:
            docx_data = make_docx(text, region)
            if docx_data:
                st.download_button("📄 Word (.docx)", docx_data, base_name + ".docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True, key="dl_docx_" + base_name)
        else:
            st.button("📄 Word (need python-docx)", disabled=True, use_container_width=True, key="dl_docx_dis_" + base_name)
    with col3:
        if HAS_RL:
            pdf_data = make_pdf(text, region)
            if pdf_data:
                st.download_button("📑 PDF (.pdf)", pdf_data, base_name + ".pdf",
                    mime="application/pdf", use_container_width=True, key="dl_pdf_" + base_name)
        else:
            st.button("📑 PDF (need reportlab)", disabled=True, use_container_width=True, key="dl_pdf_dis_" + base_name)

def region_note(stored_region):
    if stored_region and stored_region != region:
        st.markdown('<div class="region-warning">⚠️ Region changed from <b>' + stored_region +
                    '</b> to <b>' + region + '</b>. Click Generate/Improve again to reformat for the new region.</div>',
                    unsafe_allow_html=True)

if page == "🏠 Home":
    st.markdown('<div class="hero"><h1>📄 AI CV Builder</h1><p>Gemini (FREE) + OpenAI • Dark Mode • PDF/Word Export • Multi-JD Compare<br>Anti-hallucination: AI only uses YOUR data.</p></div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2, gap="medium")
    with col1:
        st.markdown('<div class="card-green"><h3>📝 Generate CV</h3><p>Create region-formatted CV from JD.<br>Word & PDF export.</p></div>', unsafe_allow_html=True)
        st.markdown('<div class="card-blue"><h3>📊 CV Analysis</h3><p>AI scores ATS compliance.<br>Improve & export.</p></div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="card-yellow"><h3>🔍 CV vs JD</h3><p>Compare CV vs JD. Keywords,<br>gaps & cover letter.</p></div>', unsafe_allow_html=True)
        st.markdown('<div class="card-purple"><h3>📑 Multi-JD</h3><p>Compare CV vs multiple JDs.<br>AI ranks best fit.</p></div>', unsafe_allow_html=True)
    st.markdown('<div class="stats-bar">🌍 <b>14 Regions</b> &nbsp;•&nbsp; 🤖 <b>Gemini+OpenAI</b> &nbsp;•&nbsp; 📑 <b>PDF/Word Export</b> &nbsp;•&nbsp; 📱 <b>Mobile-friendly</b></div>', unsafe_allow_html=True)
    st.markdown("---")
    st.info("💡 **Currently set to:** " + region + " - your CV will be formatted for this region. Change in sidebar anytime.")

elif page == "📝 Generate CV":
    st.title("📝 Generate CV")
    st.caption("Currently generating for: **" + region + "** - " + RD_[region]["photo"])
    region_note(st.session_state.get("gen_region"))
    jd_gen = st.text_area("📋 Job Description", height=200, key="gen_jd")
    info_gen = st.text_area("👤 Your Information", height=400, key="gen_info", value=st.session_state.get("gen_info_val", INFO_T))
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
        st.markdown("### ✨ Generated CV (formatted for " + st.session_state.get("gen_region", region) + ")")
        st.text_area("Result", st.session_state["gen_result"], height=500, key="gen_out", label_visibility="collapsed")
        st.markdown("**Download:**")
        download_buttons(st.session_state["gen_result"], st.session_state.get("gen_region", region), "generated_cv")

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
                st.session_state["cmp_cv"] = cv_text
                st.session_state["cmp_jd"] = jd_cmp
                st.session_state["cmp_region"] = region
    if "cmp_result" in st.session_state:
        st.markdown("### 📊 Comparison Result")
        st.text_area("Result", st.session_state["cmp_result"], height=400, key="cmp_out", label_visibility="collapsed")
        st.markdown("---")
        st.markdown("### ✨ Improve CV with Keywords")
        col_k1, col_k2 = st.columns(2)
        with col_k1:
            kw_strength = st.radio("Keyword Strength", ["Experience-based", "Force ALL"], key="kw_str", horizontal=True)
        with col_k2:
            kw_mode = st.radio("Keywords from", ["Auto from JD", "Manual"], key="kw_mode", horizontal=True)
        manual_kw = ""
        if kw_mode == "Manual":
            manual_kw = st.text_input("Enter keywords (comma-separated)", placeholder="Docker, Kubernetes, CI/CD")
        if st.button("✨ Improve CV", type="primary", use_container_width=True):
            strength = "force_all" if kw_strength == "Force ALL" else "experience"
            result = call_ai(p_imp_cmp(st.session_state["cmp_cv"], st.session_state["cmp_result"],
                                       st.session_state["cmp_jd"], region, RD_[region], manual_kw, strength))
            if result:
                st.session_state["cmp_improved"] = result
                st.session_state["cmp_region"] = region
        if "cmp_improved" in st.session_state:
            st.markdown("#### ✨ Improved CV (formatted for " + st.session_state.get("cmp_region", region) + ")")
            st.text_area("Improved", st.session_state["cmp_improved"], height=500, key="cmp_imp_out", label_visibility="collapsed")
            st.markdown("**Download Improved CV:**")
            download_buttons(st.session_state["cmp_improved"], st.session_state.get("cmp_region", region), "improved_cv")
        st.markdown("---")
        st.markdown("### 📨 Cover Letter")
        if st.button("📨 Generate Cover Letter", type="primary", use_container_width=True):
            cv_src = st.session_state.get("cmp_improved", st.session_state.get("cmp_cv", ""))
            result = call_ai(p_cover(cv_src, st.session_state["cmp_jd"], region, RD_[region]))
            if result:
                st.session_state["cmp_cl"] = result
        if "cmp_cl" in st.session_state:
            st.text_area("Cover Letter", st.session_state["cmp_cl"], height=400, key="cmp_cl_out", label_visibility="collapsed")
            st.markdown("**Download Cover Letter:**")
            download_buttons(st.session_state["cmp_cl"], region, "cover_letter")

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
                st.session_state["ana_cv"] = cv_ana
                st.session_state["ana_region"] = region
    if "ana_result" in st.session_state:
        st.markdown("### 📊 Analysis")
        st.text_area("Analysis", st.session_state["ana_result"], height=400, key="ana_out", label_visibility="collapsed")
        st.markdown("---")
        if st.button("✨ Improve CV", type="primary", use_container_width=True):
            result = call_ai(p_imp(st.session_state["ana_cv"], st.session_state["ana_result"], region, RD_[region]))
            if result:
                st.session_state["ana_improved"] = result
                st.session_state["ana_region"] = region
        if "ana_improved" in st.session_state:
            st.markdown("### ✨ Improved CV (formatted for " + st.session_state.get("ana_region", region) + ")")
            st.text_area("Improved", st.session_state["ana_improved"], height=500, key="ana_imp_out", label_visibility="collapsed")
            st.markdown("**Download:**")
            download_buttons(st.session_state["ana_improved"], st.session_state.get("ana_region", region), "improved_cv_ana")

elif page == "📑 Multi-JD Compare":
    st.title("📑 Multi-JD Compare")
    st.caption("Using region: **" + region + "**")
    cv_multi_file = st.file_uploader("Upload CV", type=["pdf", "docx", "txt"], key="multi_file")
    cv_multi_input = st.text_area("Or paste CV", height=200, key="multi_cv_input")
    cv_multi = read_file(cv_multi_file) if cv_multi_file else cv_multi_input
    n_jds = st.number_input("Number of JDs to compare", min_value=2, max_value=10, value=2, key="multi_n")
    jd_texts = []
    for i in range(int(n_jds)):
        jd_texts.append(st.text_area("📋 JD " + str(i+1), height=150, key="multi_jd_" + str(i)))
    if st.button("📑 Compare All JDs", type="primary", use_container_width=True):
        if not cv_multi.strip(): st.warning("Please provide a CV.")
        else:
            combined = "\n\n".join(["--- JD " + str(i+1) + " ---\n" + t for i, t in enumerate(jd_texts) if t.strip()])
            if not combined.strip(): st.warning("Please fill in at least one JD.")
            else:
                result = call_ai(p_multi(cv_multi, combined, region, RD_[region]))
                if result: st.session_state["multi_result"] = result
    if "multi_result" in st.session_state:
        st.markdown("### 📊 Comparison Results")
        st.text_area("Results", st.session_state["multi_result"], height=500, key="multi_out", label_visibility="collapsed")
        st.markdown("**Download:**")
        download_buttons(st.session_state["multi_result"], region, "multi_jd_results")

elif page == "🎤 Interview Prep":
    st.title("🎤 Interview Preparation")
    st.info("💡 Questions come from AI knowledge of interview patterns (no web search).")
    jd_intv = st.text_area("📋 Job Description", height=250, key="intv_jd")
    if st.button("🎯 Generate Questions", type="primary", use_container_width=True):
        if not jd_intv.strip(): st.warning("Please paste a JD.")
        else:
            result = call_ai(p_int(jd_intv))
            if result: st.session_state["intv_result"] = result
    if "intv_result" in st.session_state:
        st.markdown("### 🎯 Questions + Answer Frameworks")
        st.text_area("Questions", st.session_state["intv_result"], height=600, key="intv_out", label_visibility="collapsed")
        st.markdown("**Download:**")
        download_buttons(st.session_state["intv_result"], region, "interview_questions")

elif page == "🎙️ Mock Interview":
    st.title("🎙️ Mock Interview")
    st.caption("AI generates questions, you answer, AI evaluates with STAR + score.")
    with st.container(border=True):
        st.markdown("### 📋 Question Source")
        source = st.radio("Source", ["📋 From Job Description", "🎯 By Domain / Field"],
                          horizontal=True, key="mock_src", label_visibility="collapsed")
        if "JD" in source:
            mock_jd = st.text_area("Paste JD", height=200, key="mock_jd")
        else:
            mock_domain = st.selectbox("Select Domain", DOMAINS, key="mock_dom")
        n_q = st.radio("Number of questions", ["10", "15", "20"], index=1, horizontal=True, key="mock_n")
    if st.button("🎯 Generate Questions", type="primary", use_container_width=True):
        if "JD" in source:
            if not mock_jd.strip(): st.warning("Please paste a JD.")
            else:
                result = call_ai(p_mock_jd(mock_jd, int(n_q)))
                if result: st.session_state["mock_questions"] = result
        else:
            result = call_ai(p_mock_dom(mock_domain, int(n_q)))
            if result: st.session_state["mock_questions"] = result
    if "mock_questions" in st.session_state:
        st.markdown("### 📋 Generated Questions")
        st.text_area("Questions", st.session_state["mock_questions"], height=350, key="mock_q_out", label_visibility="collapsed")
        st.markdown("---")
        st.markdown("### 📝 Practice & Evaluate")
        mock_q = st.text_area("Question", height=80, key="mock_q_in")
        mock_a = st.text_area("Your Answer", height=200, key="mock_a_in")
        if st.button("📊 Evaluate My Answer", type="primary", use_container_width=True):
            if not mock_q.strip() or not mock_a.strip():
                st.warning("Please provide both question and answer.")
            else:
                result = call_ai(p_mock_eval(mock_q, mock_a))
                if result: st.session_state["mock_feedback"] = result
        if "mock_feedback" in st.session_state:
            st.markdown("### 📊 AI Feedback")
            st.text_area("Feedback", st.session_state["mock_feedback"], height=400, key="mock_fb_out", label_visibility="collapsed")
            st.markdown("**Download:**")
            download_buttons(st.session_state["mock_feedback"], region, "mock_feedback")

elif page == "🧑‍💼 Coaching":
    st.title("🧑‍💼 Career Coaching")
    topic = st.selectbox("Topic", ["Career Change", "Salary Negotiation", "Skill Development", "Job Search",
                                    "Interview Confidence", "LinkedIn", "Networking", "Promotion",
                                    "Work-Life Balance", "Remote Work"], key="coach_topic")
    context = st.text_area("Your situation/context", height=200, key="coach_ctx",
                           placeholder="Describe where you are, what you want, and challenges...")
    if st.button("💡 Get Advice", type="primary", use_container_width=True):
        if not context.strip(): st.warning("Please describe your situation.")
        else:
            result = call_ai(p_coach(topic, context))
            if result: st.session_state["coach_result"] = result
    if "coach_result" in st.session_state:
        st.markdown("### 💡 Personalised Advice")
        st.text_area("Advice", st.session_state["coach_result"], height=500, key="coach_out", label_visibility="collapsed")
        st.markdown("**Download:**")
        download_buttons(st.session_state["coach_result"], region, "career_advice")

elif page == "⚙️ Settings":
    st.title("⚙️ Settings & Info")
    with st.container(border=True):
        st.markdown("### 🤖 AI Providers")
        st.markdown("- **Gemini (FREE):** https://aistudio.google.com/apikey")
        st.markdown("- **OpenAI (Paid):** https://platform.openai.com/api-keys")
    with st.container(border=True):
        st.markdown("### 🌍 Region-Aware CV Generation")
        st.markdown("CV format adapts based on selected region:")
        st.markdown("- **UK/US/Canada/Australia**: NO photo, achievement-focused")
        st.markdown("- **Germany/Japan/China**: Photo placeholder included")
        st.markdown("- **India**: Declaration statement at bottom")
        st.markdown("- **Italy**: GDPR privacy clause")
        st.markdown("- **France/EU**: CEFR language levels")
        st.markdown("- **UAE/Gulf**: Nationality + visa status mandatory")
        st.markdown("Change region in sidebar anytime and regenerate to update format.")
    with st.container(border=True):
        st.markdown("### 💾 Three Download Formats")
        st.markdown("- **💾 Save (.txt)**: Plain text")
        st.markdown("- **📄 Word (.docx)**: Formatted with region-specific styling, photo placeholder if needed")
        st.markdown("- **📑 PDF (.pdf)**: Professional PDF, photo box for relevant regions")
    with st.container(border=True):
        st.markdown("### 📱 Use on Phone")
        st.markdown("**iPhone (Safari):** Share → Add to Home Screen")
        st.markdown("**Android (Chrome):** Menu → Add to Home Screen")
    docx_status = "OK" if HAS_DOCX else "Missing"
    rl_status = "OK" if HAS_RL else "Missing"
    st.caption("SSL: " + SSL_M + " | python-docx: " + docx_status + " | reportlab: " + rl_status)
