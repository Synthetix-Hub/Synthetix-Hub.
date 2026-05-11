import streamlit as st
import pytesseract
from PIL import Image
import requests
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io
from docx import Document # WICHTIG: Hier oben laden!

# ======================================================
# SYNTHETIX HUB - KONFIGURATION
# ======================================================
GROQ_API_KEY = "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"
PEXELS_API_KEY = "7FGKkAuPmxhJIQ9ew2S5HRSAnrkMkjDctylDNrNtzbsdxLhJ6fKOq5fy"
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

st.set_page_config(page_title="Synthetix Hub | AI OS", page_icon="🌌", layout="wide")

# Custom CSS
st.markdown("""
    <style>
    .main-title { font-size: 45px; font-weight: 800; color: #00ffa6; margin-bottom: 10px; }
    .stButton>button { background-color: #00ffa6 !important; color: black !important; font-weight: bold !important; border-radius: 10px !important; }
    </style>
    """, unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("## 🌌 Synthetix Hub")
    st.divider()
    menu = st.radio("Tool auswählen:", ["📽️ Präsentations-Generator", "📝 Text-Optimierer", "🎓 Homework Engine", "⚙️ Einstellungen"])
    st.divider()
    st.caption("Version 1.2.1 | Admin Mode")

# --- TOOL 1: PRÄSENTATIONEN ---
if menu == "📽️ Präsentations-Generator":
    st.markdown('<p class="main-title">📽️ Präsentations-Generator</p>', unsafe_allow_html=True)
    t_in = st.text_input("Thema:")
    if st.button("🚀 Engine starten") and t_in:
        st.info("Generierung läuft...")

# --- TOOL 2: TEXT-OPTIMIERER ---
elif menu == "📝 Text-Optimierer":
    st.markdown('<p class="main-title">📝 Text-Optimierer</p>', unsafe_allow_html=True)
    raw = st.text_area("Dein Text:", height=200, placeholder="Tippe hier...")
    if st.button("✨ Veredeln") and raw:
        st.info("Poliere Text...")

# --- TOOL 3: HOMEWORK ENGINE ---
elif menu == "🎓 Homework Engine":
    st.markdown('<p class="main-title">🎓 Homework Engine</p>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        aufgabe = st.text_area("Aufgabe/Zusatzinfos:", height=150, placeholder="Was soll ich tun?")
        name_input = st.text_input("Dein Name:", "Schüler")
    with col2:
        img_file = st.file_uploader("Arbeitsblatt scannen", type=["jpg", "png", "jpeg"])

    if st.button("🔍 Lösung generieren & Word erstellen"):
        if aufgabe or img_file:
            with st.spinner("Synthetix arbeitet..."):
                input_data = aufgabe
                if img_file:
                    img = Image.open(img_file)
                    input_data = pytesseract.image_to_string(img, lang='deu') + "\n" + aufgabe
                
                # API Call
                headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}
                payload = {"model": "llama-3.3-70b-versatile", "messages": [{"role": "user", "content": f"Löse ausführlich: {input_data}"}]}
                res = requests.post("https://api.groq.com/openai/v1/chat/completions", json=payload, headers=headers).json()
                loesung = res['choices'][0]['message']['content']
                
                # WORD BAUEN
                doc = Document()
                doc.add_heading('Hausaufgaben-Lösung', 0)
                doc.add_paragraph(f"Name: {name_input}")
                doc.add_paragraph(loesung)
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                
                st.markdown("### ✅ Lösung gefunden!")
                st.write(loesung)
                st.download_button("📥 Als Word (.docx) herunterladen", doc_io, f"Hausaufgabe_{name_input}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# --- SETTINGS ---
elif menu == "⚙️ Einstellungen":
    st.markdown('<p class="main-title">⚙️ Einstellungen</p>', unsafe_allow_html=True)
    st.write("System: Online")