import os
import io
import json
import re
import random
import hashlib
import secrets
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from pypdf import PdfReader, PdfWriter
from docx import Document
from pptx import Presentation
from dotenv import load_dotenv
from google import genai
from google.genai import types

load_dotenv()
api_key = os.getenv("GEMINI_API_KEY")

st.set_page_config(
    page_title="Study Buddy AI",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------- OBSIDIAN & EMERALD AI UI ----------------
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; }

    .stApp {
        background: radial-gradient(circle at 15% 12%, rgba(16, 185, 129, 0.11) 0%, transparent 40%),
                    radial-gradient(circle at 85% 85%, rgba(6, 182, 212, 0.10) 0%, transparent 42%),
                    linear-gradient(165deg, #090d16 0%, #0f172a 55%, #070a12 100%);
        color: #f1f5f9;
    }
    [data-testid="stSidebar"] {
        background: rgba(10, 15, 28, 0.96);
        border-right: 1px solid rgba(52, 211, 153, 0.18);
    }
    [data-testid="stMetric"] {
        background: rgba(15, 23, 42, 0.78);
        border: 1px solid rgba(52, 211, 153, 0.22);
        padding: 15px 18px;
        border-radius: 14px;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.28);
    }
    .brand-panel {
        background: linear-gradient(145deg, rgba(15, 23, 42, 0.9), rgba(6, 78, 59, 0.35));
        border: 1px solid rgba(52, 211, 153, 0.28);
        border-radius: 20px;
        padding: 30px;
    }
    .feature-item {
        background: rgba(9, 13, 22, 0.65);
        border: 1px solid rgba(148, 163, 184, 0.16);
        border-radius: 12px;
        padding: 13px 16px;
        margin-bottom: 10px;
    }
    .auth-box {
        background: rgba(15, 23, 42, 0.85);
        border: 1px solid rgba(148, 163, 184, 0.2);
        border-radius: 20px;
        padding: 26px;
    }
    .hero-top {
        background: linear-gradient(95deg, #064e3b 0%, #0f766e 50%, #155e75 100%);
        padding: 20px 26px;
        border-radius: 16px;
        color: #ffffff;
        margin-bottom: 18px;
        border: 1px solid rgba(52, 211, 153, 0.3);
    }
    .flashcard-front {
        background: linear-gradient(135deg, #0f172a 0%, #134e4a 100%);
        border: 2px solid #14b8a6;
        border-radius: 20px;
        padding: 46px 28px;
        text-align: center;
        min-height: 230px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        margin: 12px 0;
    }
    .flashcard-back {
        background: linear-gradient(135deg, #064e3b 0%, #090d16 100%);
        border: 2px solid #10b981;
        border-radius: 20px;
        padding: 46px 28px;
        text-align: center;
        min-height: 230px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        margin: 12px 0;
    }
    .clue-box {
        background: rgba(245, 158, 11, 0.14);
        border: 1px solid rgba(251, 191, 36, 0.5);
        border-radius: 12px;
        padding: 12px 16px;
        margin: 8px 0;
        color: #fef3c7;
    }
    .hint-msg {
        background: rgba(245, 158, 11, 0.11);
        border-left: 4px solid #f59e0b;
        border-radius: 10px;
        padding: 12px 15px;
        line-height: 1.55;
    }
    .rev-card {
        background: rgba(15, 23, 42, 0.72);
        border: 1px solid rgba(148, 163, 184, 0.2);
        border-radius: 12px;
        padding: 14px 16px;
        margin-bottom: 10px;
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(90deg, #059669, #0d9488);
        border: none;
        border-radius: 10px;
        font-weight: 600;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

if not api_key or not api_key.strip():
    st.error("GEMINI_API_KEY missing in .env file.")
    st.stop()

@st.cache_resource
def get_client():
    return genai.Client(api_key=api_key.strip())

client = get_client()

# ---------------- DATABASE & SESSION PERSISTENCE ----------------
DB_FILE = "students_db.json"

def load_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {"users": {}}
    return {"users": {}}

def save_db(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

def hash_pwd(pwd: str) -> str:
    return hashlib.sha256(pwd.encode()).hexdigest()

WORKSPACE_KEYS = [
    "study_text", "source_name", "chat_history",
    "flashcards", "fc_index", "fc_flipped", "fc_show_clue",
    "quiz", "quiz_submitted", "last_quiz_result", "quiz_open_hints"
]

def default_workspace():
    return {
        "study_text": "",
        "source_name": "",
        "chat_history": [],
        "flashcards": [],
        "fc_index": 0,
        "fc_flipped": False,
        "fc_show_clue": False,
        "quiz": [],
        "quiz_submitted": False,
        "last_quiz_result": None,
        "quiz_open_hints": []
    }

def sync_workspace_to_db(uid: str):
    if not uid:
        return
    db = load_db()
    if uid in db["users"]:
        db["users"][uid]["saved_workspace"] = {k: st.session_state[k] for k in WORKSPACE_KEYS}
        save_db(db)

def load_workspace_into_state(ws: dict):
    base = default_workspace()
    if isinstance(ws, dict):
        base.update(ws)
    for k, v in base.items():
        st.session_state[k] = v

def reset_study_session(uid: str, text: str = "", source: str = ""):
    load_workspace_into_state(default_workspace())
    st.session_state.study_text = text
    st.session_state.source_name = source
    sync_workspace_to_db(uid)

# Initialize session defaults
if "logged_in_user" not in st.session_state:
    st.session_state.logged_in_user = None
for k, v in default_workspace().items():
    if k not in st.session_state:
        st.session_state[k] = v

# Restore session ONLY if URL query params (?u=...&t=...) exist (browser refresh F5)
if not st.session_state.logged_in_user:
    qp_user = st.query_params.get("u", "")
    qp_token = st.query_params.get("t", "")
    if qp_user and qp_token:
        db = load_db()
        u_rec = db["users"].get(qp_user)
        if u_rec and u_rec.get("session_token") == qp_token:
            st.session_state.logged_in_user = qp_user
            load_workspace_into_state(u_rec.get("saved_workspace", {}))

# ---------------- HELPER TO CONVERT **TEXT** TO HTML <b>TEXT</b> ----------------
def md_bold_to_html(text: str) -> str:
    html_txt = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text.strip())
    html_txt = html_txt.replace("*", "")
    return html_txt.replace("\n", "<br>")

# ---------------- GEMINI & DOCUMENT HELPERS ----------------
def extract_json(raw_text: str):
    cleaned = re.sub(r"```(?:json)?", "", raw_text).strip().rstrip("`").strip()
    s, e = cleaned.find("["), cleaned.rfind("]")
    if s != -1 and e != -1:
        cleaned = cleaned[s:e+1]
    return json.loads(cleaned)

def call_gemini(contents, max_tokens: int = 2500) -> str:
    models = ["gemini-2.5-flash", "gemini-2.5-flash-lite", "gemini-3.8-flash-lite", "gemini-3.5-flash-lite", "gemini-3.8-flash"]
    last_err = None
    for m in models:
        try:
            res = client.models.generate_content(
                model=m,
                contents=contents,
                config=types.GenerateContentConfig(temperature=0.35, max_output_tokens=max_tokens)
            )
            if res and res.text:
                return res.text
        except Exception as exc:
            last_err = exc
    raise last_err

def extract_text_from_file(uploaded_file) -> str:
    fname = uploaded_file.name.lower()
    raw = uploaded_file.getvalue()

    if fname.endswith(".pdf"):
        uploaded_file.seek(0)
        reader = PdfReader(uploaded_file)
        text = "\n".join(p.extract_text() or "" for p in reader.pages)
        if len(text.strip()) >= 50:
            return text
        writer = PdfWriter()
        for i in range(min(len(reader.pages), 2)):
            writer.add_page(reader.pages[i])
        buf = io.BytesIO()
        writer.write(buf)
        return call_gemini([
            types.Part.from_bytes(data=buf.getvalue(), mime_type="application/pdf"),
            "Extract all readable study text and bullet points concisely."
        ], max_tokens=1500)

    if fname.endswith(".docx"):
        uploaded_file.seek(0)
        doc = Document(uploaded_file)
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for tbl in doc.tables:
            for row in tbl.rows:
                parts.append(" | ".join(c.text.strip() for c in row.cells if c.text.strip()))
        return "\n".join(parts)

    if fname.endswith(".pptx"):
        uploaded_file.seek(0)
        prs = Presentation(uploaded_file)
        slides = []
        for idx, slide in enumerate(prs.slides, 1):
            items = [f"--- Slide {idx} ---"]
            for sh in slide.shapes:
                if hasattr(sh, "text") and sh.text.strip():
                    items.append(sh.text.strip())
            slides.append("\n".join(items))
        return "\n\n".join(slides)

    if fname.endswith(".txt"):
        return raw.decode("utf-8", errors="ignore")
    return ""

# ---------------- AUTH SCREEN ----------------
def render_auth_screen():
    st.write("")
    left, right = st.columns([1.2, 1.05], gap="large")

    with left:
        st.markdown("""
        <div class="brand-panel">
            <div style="color:#34d399; font-size:0.78rem; font-weight:700; letter-spacing:1px;">✦ INTELLIGENT ACADEMIC WORKSPACE</div>
            <h1 style="margin:6px 0 10px 0; font-size:2.35rem; font-weight:800; color:#ffffff;">Study Buddy AI</h1>
            <p style="color:#cbd5e1; font-size:1rem; line-height:1.6;">
                Upload lecture slides, PDFs, or Word documents to unlock context-grounded AI tutoring, non-spoiler guided hints, active recall flashcards, and spaced revision analytics.
            </p>
            <div style="margin-top:20px;">
                <div class="feature-item"><b>◈ Context-Grounded Tutor + Hint Mode</b><br><span style="color:#94a3b8; font-size:0.85rem;">Remembers follow-up questions and guides you step-by-step with clues.</span></div>
                <div class="feature-item"><b>◈ Active Recall Flashcard Decks</b><br><span style="color:#94a3b8; font-size:0.85rem;">Generate up to 30 flip cards with progressive clues and shuffle controls.</span></div>
                <div class="feature-item"><b>◈ Self-Assessment & Spaced Revision</b><br><span style="color:#94a3b8; font-size:0.85rem;">Take custom MCQs and track mastery analytics with automated review dates.</span></div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    with right:
        st.markdown('<div class="auth-box">', unsafe_allow_html=True)
        st.markdown("### Access Workspace")
        st.caption("Sign in or create a new student account.")
        t_login, t_reg = st.tabs(["Student Login", "Create Account"])
        db = load_db()

        with t_login:
            with st.form("login_form"):
                l_id = st.text_input("Student ID / Username", placeholder="e.g. alex_101")
                l_pwd = st.text_input("Password", type="password", placeholder="Minimum 5 characters")
                submit_login = st.form_submit_button("Login →", type="primary", use_container_width=True)

            if submit_login:
                uid = l_id.strip().lower()
                if len(l_pwd) < 5:
                    st.error("Password must be at least 5 characters long.")
                elif uid in db["users"] and db["users"][uid]["password"] == hash_pwd(l_pwd):
                    token = secrets.token_hex(16)
                    db["users"][uid]["session_token"] = token
                    db["users"][uid]["saved_workspace"] = default_workspace()
                    save_db(db)
                    st.session_state.logged_in_user = uid
                    load_workspace_into_state(default_workspace())
                    st.query_params["u"] = uid
                    st.query_params["t"] = token
                    st.rerun()
                else:
                    st.error("Invalid Student ID or Password.")

        with t_reg:
            with st.form("register_form"):
                r_name = st.text_input("Full Name", placeholder="e.g. Alex Rivera")
                r_id = st.text_input("Choose Student ID / Username", placeholder="e.g. alex_101")
                r_pwd = st.text_input("Create Password", type="password", placeholder="Minimum 5 characters")
                submit_reg = st.form_submit_button("Create Account ✦", type="primary", use_container_width=True)

            if submit_reg:
                uid = r_id.strip().lower()
                if not r_name.strip() or not uid:
                    st.warning("Please enter your Full Name and Student ID.")
                elif len(r_pwd) < 5:
                    st.error("Password must be at least 5 characters long.")
                elif uid in db["users"]:
                    st.error("This Student ID already exists.")
                else:
                    db["users"][uid] = {
                        "name": r_name.strip(),
                        "password": hash_pwd(r_pwd),
                        "session_token": "",
                        "saved_workspace": default_workspace(),
                        "revision_log": []
                    }
                    save_db(db)
                    st.success("Account created! Switch to 'Student Login' to sign in.")
        st.markdown('</div>', unsafe_allow_html=True)

    components.html("""
    <script>
    const doc = window.parent.document;
    function bindSmartEnter() {
        doc.querySelectorAll('form').forEach(form => {
            const inputs = Array.from(form.querySelectorAll('input[type="text"], input[type="password"]'));
            inputs.forEach((inp, idx) => {
                if (inp.dataset.bound) return;
                inp.dataset.bound = "1";
                inp.addEventListener('keydown', function(e) {
                    if (e.key === 'Enter' && idx < inputs.length - 1) {
                        e.preventDefault();
                        e.stopPropagation();
                        inputs[idx + 1].focus();
                    }
                }, true);
            });
        });
    }
    bindSmartEnter();
    new MutationObserver(bindSmartEnter).observe(doc.body, { childList: true, subtree: true });
    </script>
    """, height=0)

if not st.session_state.logged_in_user:
    render_auth_screen()
    st.stop()

# ---------------- LOGGED-IN WORKSPACE ----------------
db = load_db()
current_uid = st.session_state.logged_in_user
user_profile = db["users"].get(current_uid, {"name": current_uid, "revision_log": []})

with st.sidebar:
    st.markdown(f"""
    <div style="background: rgba(16,185,129,0.12); padding: 14px; border-radius: 14px; border: 1px solid rgba(52,211,153,0.3); margin-bottom: 12px;">
        <div style="font-size: 0.7rem; letter-spacing: 1px; color: #34d399; font-weight: 700;">STUDENT PROFILE</div>
        <div style="font-size: 1.15rem; font-weight: 800; color: #ffffff;">{user_profile['name']}</div>
        <div style="font-size: 0.78rem; color: #94a3b8;">@{current_uid}</div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("Sign Out", use_container_width=True):
        db = load_db()
        if current_uid in db["users"]:
            db["users"][current_uid]["session_token"] = ""
            db["users"][current_uid]["saved_workspace"] = default_workspace()
            save_db(db)
        st.query_params.clear()
        st.session_state.logged_in_user = None
        load_workspace_into_state(default_workspace())
        st.rerun()

    st.divider()
    st.subheader("1. Study Material")
    src_mode = st.radio("Source:", ["Upload File", "Paste Text"], horizontal=True)

    if src_mode == "Upload File":
        up_file = st.file_uploader("PDF, Word (.docx), PPT (.pptx), TXT", type=["pdf", "docx", "pptx", "txt"])
        if up_file and st.button("Load Document", type="primary", use_container_width=True):
            with st.spinner("Reading document..."):
                try:
                    extracted = extract_text_from_file(up_file)
                    if len(extracted.strip()) < 20:
                        st.error("File appears empty or unreadable.")
                    else:
                        reset_study_session(current_uid, extracted.strip()[:18000], up_file.name)
                        st.rerun()
                except Exception as e:
                    st.error(f"Read Error: {e}")
    else:
        p_text = st.text_area("Paste study notes:", height=130, placeholder="Paste lecture notes here...")
        if st.button("Load Notes", type="primary", use_container_width=True):
            if len(p_text.strip()) < 30:
                st.warning("Please paste at least 2–3 sentences.")
            else:
                reset_study_session(current_uid, p_text.strip()[:18000], "Pasted Study Notes")
                st.rerun()

    if st.session_state.study_text:
        st.success(f"**Loaded:** {st.session_state.source_name}\n\n**Chars:** {len(st.session_state.study_text):,}")

    st.divider()
    st.subheader("2. Tutor Mode")
    guided_mode = st.toggle("Guided Hint Mode", value=False)
    if guided_mode:
        st.warning("**Hint Mode Active**")
    else:
        st.info("**Direct Mode Active**")

# Header & KPIs
st.markdown(f"""
<div class="hero-top">
    <h2 style="margin:0; font-weight:800;">Study Buddy AI</h2>
    <p style="margin:4px 0 0 0; opacity:0.9;">Welcome, <b>{user_profile['name']}</b> — Auto-saved learning session active.</p>
</div>
""", unsafe_allow_html=True)

rev_logs = user_profile.get("revision_log", [])
k1, k2, k3, k4 = st.columns(4)
k1.metric("Active Material", "Loaded" if st.session_state.study_text else "None")
k2.metric("Flashcards", f"{len(st.session_state.flashcards)} Cards")
k3.metric("Quizzes Taken", f"{len(rev_logs)} Tests")
avg_sc = round(sum(int(x["Score"].replace("%", "")) for x in rev_logs) / len(rev_logs)) if rev_logs else 0
k4.metric("Mastery Score", f"{avg_sc}%")

st.write("")

if not st.session_state.study_text:
    st.markdown("""
    <div class="auth-box" style="text-align:center; padding: 44px 20px;">
        <h3>Upload Your Study Material to Begin</h3>
        <p style="color:#94a3b8;">Upload a PDF, Word (.docx), PowerPoint (.pptx), or paste notes in the left sidebar.</p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

tab1, tab2, tab3, tab4 = st.tabs([
    "Ask & Guided Tutor",
    "AI Flashcards",
    "Practice Quiz",
    "Spaced Revision & Analytics"
])

# ---------------- TAB 1: CHAT WITH SHORT ANALOGY CLUES (NO ASTERISKS, NO SPOILERS) ----------------
with tab1:
    c_head, c_clr = st.columns([4.5, 1])
    with c_head:
        st.caption(f"Source: **{st.session_state.source_name}** • Mode: **{'Guided Hint Mode' if guided_mode else 'Direct Tutor Mode'}**")
    with c_clr:
        if st.button("Clear Chat", use_container_width=True):
            st.session_state.chat_history = []
            sync_workspace_to_db(current_uid)
            st.rerun()

    chat_box = st.container(height=400, border=True)
    with chat_box:
        if not st.session_state.chat_history:
            st.markdown("<div style='text-align:center; padding:75px 15px; color:#64748b;'>Ask any question from your uploaded study material.</div>", unsafe_allow_html=True)
        for m in st.session_state.chat_history:
            with st.chat_message(m["role"]):
                if m["role"] == "assistant" and m.get("is_hint"):
                    clean_html = md_bold_to_html(m["content"])
                    st.markdown(f'<div class="hint-msg">{clean_html}</div>', unsafe_allow_html=True)
                else:
                    st.markdown(m["content"])

    user_q = st.chat_input("Type something...")
    if user_q and user_q.strip():
        clean_q = user_q.strip()
        if not st.session_state.chat_history or st.session_state.chat_history[-1]["content"] != clean_q:
            st.session_state.chat_history.append({"role": "user", "content": clean_q})
            recent = st.session_state.chat_history[-7:-1]
            mem = "\n".join(f"{'Student' if t['role']=='user' else 'Tutor'}: {t['content']}" for t in recent)

            if guided_mode:
                instr = (
                    "STRICT HINT MODE (NO EXPLANATION, NO SPOILERS):\n"
                    "- Do NOT explain the topic and do NOT write the definition or answer.\n"
                    "- Give ONLY 1 short everyday analogy or thinking angle (MAX 15 words) + 1 short guiding question (MAX 12 words).\n"
                    "- Format strictly like this (2 short lines only):\n"
                    "**Clue:** [10-15 word non-spoiler thinking angle]\n"
                    "**Think:** [1 short question for the student]"
                )
            else:
                instr = "DIRECT TUTOR MODE: Give a clear, concise, structured explanation grounded strictly in the Study Material."

            prompt = f"You are Study Buddy AI.\n{instr}\n\nSTUDY MATERIAL:\n{st.session_state.study_text[:4500]}\n\nCHAT MEMORY:\n{mem or 'None'}\n\nQUESTION:\n{clean_q}"
            try:
                reply = call_gemini(prompt, max_tokens=220 if guided_mode else 1000)
                st.session_state.chat_history.append({"role": "assistant", "content": reply, "is_hint": guided_mode})
                sync_workspace_to_db(current_uid)
                st.rerun()
            except Exception as e:
                st.error(f"AI Error: {e}")

# ---------------- TAB 2: AI FLASHCARDS ----------------
with tab2:
    fc1, fc2, fc3 = st.columns([1.4, 1.4, 1])
    with fc1:
        num_cards = st.slider("Number of Flashcards", 3, 30, 10)
    with fc2:
        st.write("")
        st.write("")
        if st.button("Generate Flashcards", type="primary", use_container_width=True):
            with st.spinner(f"Generating {num_cards} flashcards..."):
                fc_prompt = f"""Create {num_cards} distinct study flashcards from the material below.
Return ONLY a raw JSON array of objects with keys:
- "front": question or term
- "clue": an 8-12 word indirect memory trigger (category or real-life analogy) WITHOUT using answer words
- "back": concise answer

STUDY MATERIAL:
{st.session_state.study_text[:6500]}"""
                try:
                    cards = extract_json(call_gemini(fc_prompt, max_tokens=3500))
                    if isinstance(cards, list) and cards:
                        st.session_state.flashcards = cards
                        st.session_state.fc_index = 0
                        st.session_state.fc_flipped = False
                        st.session_state.fc_show_clue = False
                        sync_workspace_to_db(current_uid)
                        st.rerun()
                except Exception as e:
                    st.error(f"Flashcard Error: {e}")
    with fc3:
        st.write("")
        st.write("")
        if st.session_state.flashcards and st.button("Shuffle Deck", use_container_width=True):
            random.shuffle(st.session_state.flashcards)
            st.session_state.fc_index = 0
            st.session_state.fc_flipped = False
            st.session_state.fc_show_clue = False
            sync_workspace_to_db(current_uid)
            st.rerun()

    if st.session_state.flashcards:
        idx = st.session_state.fc_index
        total = len(st.session_state.flashcards)
        card = st.session_state.flashcards[idx]
        st.progress((idx + 1) / total, text=f"Card {idx + 1} of {total}")

        if not st.session_state.fc_flipped:
            st.markdown(f"""
            <div class="flashcard-front">
                <div style="font-size:0.78rem; letter-spacing:2px; color:#5eead4; font-weight:700; margin-bottom:10px;">CONCEPT / QUESTION ({idx+1}/{total})</div>
                <div style="font-size:1.45rem; font-weight:700; color:#ffffff;">{card.get('front', '')}</div>
            </div>
            """, unsafe_allow_html=True)
            if guided_mode and st.session_state.fc_show_clue:
                clue_txt = card.get("clue") or "Think about the broader category or purpose of this concept."
                st.markdown(f'<div class="clue-box"><b>Hint:</b> {clue_txt}</div>', unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="flashcard-back">
                <div style="font-size:0.78rem; letter-spacing:2px; color:#6ee7b7; font-weight:700; margin-bottom:10px;">ANSWER ({idx+1}/{total})</div>
                <div style="font-size:1.25rem; font-weight:600; color:#ffffff;">{card.get('back', '')}</div>
            </div>
            """, unsafe_allow_html=True)

        cols = st.columns(4 if guided_mode else 3)
        with cols[0]:
            if st.button("← Previous", disabled=(idx == 0), use_container_width=True):
                st.session_state.fc_index -= 1
                st.session_state.fc_flipped = False
                st.session_state.fc_show_clue = False
                sync_workspace_to_db(current_uid)
                st.rerun()
        if guided_mode:
            with cols[1]:
                if st.button("Show Clue", use_container_width=True):
                    st.session_state.fc_show_clue = not st.session_state.fc_show_clue
                    sync_workspace_to_db(current_uid)
                    st.rerun()
            flip_col, next_col = cols[2], cols[3]
        else:
            flip_col, next_col = cols[1], cols[2]

        with flip_col:
            if st.button("Flip Card", type="primary", use_container_width=True):
                st.session_state.fc_flipped = not st.session_state.fc_flipped
                sync_workspace_to_db(current_uid)
                st.rerun()
        with next_col:
            if st.button("Next →", disabled=(idx == total - 1), use_container_width=True):
                st.session_state.fc_index += 1
                st.session_state.fc_flipped = False
                st.session_state.fc_show_clue = False
                sync_workspace_to_db(current_uid)
                st.rerun()

# ---------------- TAB 3: PRACTICE QUIZ (ON-DEMAND "SHOW HINT" BUTTON) ----------------
with tab3:
    q1, q2 = st.columns([1.4, 2])
    with q1:
        num_q = st.slider("Number of MCQs", 3, 30, 10)
    with q2:
        st.write("")
        st.write("")
        if st.button("Generate Practice Quiz", type="primary"):
            with st.spinner(f"Building {num_q} MCQs..."):
                q_prompt = f"""Create {num_q} distinct MCQs from the study material below.
Return ONLY a raw JSON array of objects with exact keys:
- "question": string
- "options": array of 4 strings
- "answer": exact match from options
- "hint": a short 10-word thinking clue that helps eliminate wrong options WITHOUT naming the right option (NEVER mention slides or notes)
- "explanation": 1 sentence

STUDY MATERIAL:
{st.session_state.study_text[:6500]}"""
                try:
                    qs = extract_json(call_gemini(q_prompt, max_tokens=4500))
                    if isinstance(qs, list) and qs:
                        st.session_state.quiz = qs
                        st.session_state.quiz_submitted = False
                        st.session_state.last_quiz_result = None
                        st.session_state.quiz_open_hints = []
                        sync_workspace_to_db(current_uid)
                        st.rerun()
                except Exception as e:
                    st.error(f"Quiz Error: {e}")

    if st.session_state.quiz:
        open_hints = set(st.session_state.get("quiz_open_hints", []))

        # Render questions with an optional "Show Hint" toggle button (only when Guided Hint Mode is ON)
        for i, q in enumerate(st.session_state.quiz):
            if guided_mode:
                q_col, h_col = st.columns([4.5, 1])
                with q_col:
                    st.markdown(f"**Q{i+1}. {q['question']}**")
                with h_col:
                    btn_label = "Hide Hint" if i in open_hints else "✦ Show Hint"
                    if st.button(btn_label, key=f"hint_btn_{i}", use_container_width=True):
                        if i in open_hints:
                            open_hints.remove(i)
                        else:
                            open_hints.add(i)
                        st.session_state.quiz_open_hints = list(open_hints)
                        sync_workspace_to_db(current_uid)
                        st.rerun()
                if i in open_hints and q.get("hint"):
                    st.markdown(f'<div class="clue-box"><b>Hint:</b> {q["hint"]}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f"**Q{i+1}. {q['question']}**")

            st.radio(f"Q{i+1}", q["options"], key=f"q_opt_{i}", label_visibility="collapsed")
            st.divider()

        if st.button("Submit Quiz & Evaluate Score", type="primary"):
            if st.session_state.quiz_submitted:
                st.toast("Quiz is already submitted!", icon=":material/info:")
                st.warning("Quiz is already submitted! Generate a new quiz to take another test.")
            else:
                ans_map = {i: st.session_state.get(f"q_opt_{i}") for i in range(len(st.session_state.quiz))}
                score = sum(1 for i, q in enumerate(st.session_state.quiz) if ans_map[i] == q["answer"])
                total_q = len(st.session_state.quiz)
                pct = round((score / total_q) * 100)
                st.session_state.quiz_submitted = True
                st.session_state.last_quiz_result = {"score": score, "total": total_q, "pct": pct, "chosen": ans_map}

                gap = 1 if pct < 60 else (3 if pct < 85 else 7)
                next_rev = (datetime.now() + timedelta(days=gap)).strftime("%Y-%m-%d")
                db = load_db()
                if current_uid in db["users"]:
                    db["users"][current_uid]["revision_log"].append({
                        "Attempted On": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "Topic": st.session_state.source_name,
                        "Score": f"{pct}%",
                        "Mastery Level": "Needs Review" if pct < 70 else "Mastered",
                        "Next Revision Date": next_rev
                    })
                    save_db(db)
                sync_workspace_to_db(current_uid)
                st.toast("Quiz submitted successfully!", icon=":material/check_circle:")
                st.rerun()

        if st.session_state.quiz_submitted and st.session_state.last_quiz_result:
            res = st.session_state.last_quiz_result
            st.metric("Final Quiz Result", f"{res['score']} / {res['total']} ({res['pct']}%)")
            for i, q in enumerate(st.session_state.quiz):
                ch = res["chosen"].get(str(i), res["chosen"].get(i))
                if ch == q["answer"]:
                    st.success(f"**Q{i+1}: Correct** — {q['answer']}")
                else:
                    st.error(f"**Q{i+1}: Incorrect** — Selected: *{ch}* | Correct: **{q['answer']}**")
                st.caption(q.get("explanation", ""))

# ---------------- TAB 4: SPACED REVISION & ANALYTICS ----------------
with tab4:
    db = load_db()
    logs = db["users"].get(current_uid, {}).get("revision_log", [])

    if not logs:
        st.info("No quizzes completed yet. Complete a Practice Quiz to view your mastery analytics.")
    else:
        scores = [int(x["Score"].replace("%", "")) for x in logs]
        a1, a2, a3 = st.columns(3)
        a1.metric("Best Score", f"{max(scores)}%")
        a2.metric("Mastered (≥70%)", f"{sum(1 for s in scores if s >= 70)} Topics")
        a3.metric("Needs Review (<70%)", f"{sum(1 for s in scores if s < 70)} Topics")

        c_chart, c_plan = st.columns([1.3, 1], gap="large")
        with c_chart:
            st.markdown("#### Score Progress")
            df = pd.DataFrame({
                "Attempt": [f"Test #{i+1}" for i in range(len(logs))],
                "Score (%)": scores
            }).set_index("Attempt")
            st.bar_chart(df, color="#10b981", height=250)

        with c_plan:
            st.markdown("#### Revision Action Plan")
            for item in reversed(logs[-4:]):
                sc = int(item["Score"].replace("%", ""))
                clr = "#10b981" if sc >= 70 else "#f59e0b"
                st.markdown(f"""
                <div class="rev-card" style="border-left: 4px solid {clr};">
                    <b>{item['Topic']}</b><br>
                    <span style="font-size:0.84rem; color:#cbd5e1;">Score: <b>{item['Score']}</b> • Next Review: <b>{item.get('Next Revision Date', '')}</b></span>
                </div>
                """, unsafe_allow_html=True)

        st.markdown("#### Complete Spaced Revision Log")
        st.dataframe(logs, use_container_width=True)