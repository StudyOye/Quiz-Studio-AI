import os
import json
import io
import time
from typing import List, Dict, Any
from PIL import Image
import pypdf
import streamlit as st
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Quiz Show Creator AI", page_icon="⚡", layout="wide")

def get_gemini_client():
    """
    Universal API Key retrieval:
    1. Checks Streamlit Secrets (for Streamlit Community Cloud)
    2. Falls back to OS Environment Variables (for Google Colab / Local / Docker)
    """
    api_key = None
    try:
        if "GEMINI_API_KEY" in st.secrets:
            api_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass

    if not api_key:
        api_key = os.environ.get("GEMINI_API_KEY")

    if not api_key:
        st.error("🔑 Gemini API Key not found! Add GEMINI_API_KEY in Streamlit Secrets or Environment Variables.")
        st.stop()

    return genai.Client(api_key=api_key)

def extract_text_from_pdfs(pdf_files) -> str:
    combined_text = ""
    for pdf_file in pdf_files:
        reader = pypdf.PdfReader(pdf_file)
        for page in reader.pages:
            combined_text += page.extract_text() or ""
    return combined_text

def generate_questions_with_ai(rounds_config: List[Dict], extracted_text: str, images: List[Any]) -> Dict:
    client = get_gemini_client()

    prompt = f"""
    You are an expert quiz master. Generate a structured set of quiz questions based on the provided source text and images.

    Round Configuration Requirements:
    {json.dumps(rounds_config, indent=2)}

    Source Material Text:
    {extracted_text[:8000]}

    Return STRICT JSON adhering to this exact schema:
    {{
      "rounds": [
        {{
          "round_name": "Round Name",
          "questions": [
            {{
              "id": 1,
              "question": "Question text?",
              "options": ["Option A", "Option B", "Option C", "Option D"],
              "answer": "Correct Answer",
              "explanation": "Explanation for Quiz Master"
            }}
          ]
        }}
      ]
    }}
    """

    contents = [prompt]
    for img in images:
        contents.append(Image.open(img))

    # Active Gemini Models Target Chain
    candidate_models = [
        "gemini-3.8-flash",
        "models/gemini-3.8-flash",
        "gemini-3.1-pro-preview",
        "models/gemini-3.1-pro-preview"
    ]

    last_err = None
    for model_name in candidate_models:
        for attempt in range(3):  # Exponential backoff retry loop for 429 rate limits
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json"
                    )
                )
                return json.loads(response.text)
            except Exception as e:
                last_err = e
                err_str = str(e)
                if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                    time.sleep(6 * (attempt + 1))  # Backoff delay: 6s, 12s, 18s
                    continue
                elif "404" in err_str or "NOT_FOUND" in err_str:
                    break  # Jump immediately to next model in list
                else:
                    break

    raise last_err

def create_pptx_deck(quiz_data: Dict, rounds_config: List[Dict]) -> io.BytesIO:
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    config_map = {r['name']: r for r in rounds_config}

    # Cover Slide
    slide = prs.slides.add_slide(blank_layout)
    tb = slide.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(10.33), Inches(2))
    p = tb.text_frame.paragraphs[0]
    p.text = "GRAND QUIZ SHOW"
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.color.rgb = RGBColor(79, 70, 229)
    p.alignment = PP_ALIGN.CENTER

    for round_item in quiz_data.get("rounds", []):
        r_name = round_item["round_name"]
        r_config = config_map.get(r_name, {"timer": 30, "points": 10, "negative": 0})

        # Round Intro Slide
        slide = prs.slides.add_slide(blank_layout)
        tb = slide.shapes.add_textbox(Inches(1.5), Inches(3), Inches(10.33), Inches(2))
        p = tb.text_frame.paragraphs[0]
        p.text = f"ROUND: {r_name.upper()}"
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = RGBColor(15, 23, 42)
        p.alignment = PP_ALIGN.CENTER

        # Question Slides
        for q_idx, q in enumerate(round_item["questions"]):
            slide = prs.slides.add_slide(blank_layout)

            tb_timer = slide.shapes.add_textbox(Inches(10.2), Inches(0.4), Inches(2.6), Inches(0.8))
            p_t = tb_timer.text_frame.paragraphs[0]
            p_t.text = f"⏱️ {r_config['timer']} Secs"
            p_t.font.size = Pt(22)
            p_t.font.bold = True
            p_t.font.color.rgb = RGBColor(225, 29, 72)

            tb_info = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(9.0), Inches(0.8))
            p_i = tb_info.text_frame.paragraphs[0]
            p_i.text = f"{r_name} | Question {q_idx + 1}"
            p_i.font.size = Pt(18)
            p_i.font.color.rgb = RGBColor(100, 116, 139)

            tb_q = slide.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.33), Inches(2.3))
            tb_q.text_frame.word_wrap = True
            p_q = tb_q.text_frame.paragraphs[0]
            p_q.text = q["question"]
            p_q.font.size = Pt(28)
            p_q.font.bold = True
            p_q.font.color.rgb = RGBColor(15, 23, 42)

            if q.get("options") and len(q["options"]) > 0:
                top_pos = 4.2
                for opt_idx, opt in enumerate(q["options"]):
                    tb_opt = slide.shapes.add_textbox(Inches(1.2), Inches(top_pos), Inches(10.5), Inches(0.6))
                    p_opt = tb_opt.text_frame.paragraphs[0]
                    p_opt.text = f"{chr(65 + opt_idx)}. {opt}"
                    p_opt.font.size = Pt(22)
                    p_opt.font.color.rgb = RGBColor(51, 65, 85)
                    top_pos += 0.7

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

def create_rulebook_pdf(rounds_config: List[Dict]) -> io.BytesIO:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []

    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=22, leading=26, textColor=colors.HexColor("#4F46E5"), spaceAfter=15)
    story.append(Paragraph("Quiz Master Official Rulebook", title_style))
    story.append(Spacer(1, 10))

    table_data = [["Round Name", "Timer", "Points (+)", "Negative (-)", "Buzzer Mode"]]
    for r in rounds_config:
        table_data.append([
            r['name'],
            f"{r['timer']}s",
            f"+{r['points']}",
            f"-{r['negative']}",
            "Enabled" if r['buzzer'] else "Disabled"
        ])

    table = Table(table_data, colWidths=[150, 80, 80, 80, 100])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#4F46E5")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 11),
        ('BOTTOMPADDING', (0,0), (-1,0), 8),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#F8FAFC")),
        ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#CBD5E1")),
    ]))

    story.append(table)
    doc.build(story)
    buffer.seek(0)
    return buffer

def create_qa_master_pdf(quiz_data: Dict) -> io.BytesIO:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []

    title_style = ParagraphStyle('QATitle', parent=styles['Heading1'], fontSize=22, leading=26, textColor=colors.HexColor("#4F46E5"), spaceAfter=15)
    round_style = ParagraphStyle('QARound', parent=styles['Heading2'], fontSize=16, leading=20, textColor=colors.HexColor("#0F172A"), spaceBefore=12, spaceAfter=8)
    q_style = ParagraphStyle('QAQuestion', parent=styles['Normal'], fontSize=11, leading=15, fontName='Helvetica-Bold', textColor=colors.HexColor("#1E293B"))
    ans_style = ParagraphStyle('QAAnswer', parent=styles['Normal'], fontSize=10, leading=14, textColor=colors.HexColor("#166534"))
    exp_style = ParagraphStyle('QAExplanation', parent=styles['Normal'], fontSize=9, leading=13, textColor=colors.HexColor("#475569"))

    story.append(Paragraph("Quiz Master Official Answer Key & Guide", title_style))
    story.append(Spacer(1, 10))

    for r in quiz_data.get("rounds", []):
        story.append(Paragraph(f"📖 {r['round_name']}", round_style))
        story.append(Spacer(1, 4))

        for idx, q in enumerate(r["questions"]):
            opts = f"<br/><i>Options:</i> {', '.join(q['options'])}" if q.get('options') else ""
            q_text = f"<b>Q{idx+1}: {q['question']}</b>{opts}"
            story.append(Paragraph(q_text, q_style))
            story.append(Spacer(1, 3))

            ans_text = f"<b>Correct Answer:</b> {q['answer']}"
            story.append(Paragraph(ans_text, ans_style))
            story.append(Spacer(1, 2))

            exp_text = f"<i>Explanation:</i> {q['explanation']}"
            story.append(Paragraph(exp_text, exp_style))
            story.append(Spacer(1, 10))

    doc.build(story)
    buffer.seek(0)
    return buffer

# Streamlit UI Construction
st.title("⚡ Quiz Show Creator AI")
st.caption("Printable Deliverables & Automated Quiz Master Guide")

st.sidebar.header("📁 Upload Source Files")
uploaded_pdfs = st.sidebar.file_uploader("Upload PDFs", type=["pdf"], accept_multiple_files=True)
uploaded_imgs = st.sidebar.file_uploader("Upload JPEGs", type=["jpeg", "jpg"], accept_multiple_files=True)

st.header("1. Configure Quiz Rounds")
num_rounds = st.number_input("Number of Rounds", min_value=1, max_value=6, value=2)

rounds_config = []
cols = st.columns(num_rounds)

for i in range(num_rounds):
    with cols[i]:
        st.subheader(f"Round {i+1}")
        r_name = st.text_input("Name", value=f"Round {i+1}", key=f"rname_{i}")
        q_count = st.number_input("Questions", min_value=1, max_value=10, value=3, key=f"qcnt_{i}")
        timer = st.number_input("Timer (sec)", min_value=5, max_value=120, value=30, key=f"timer_{i}")
        points = st.number_input("Points (+)", min_value=1, max_value=50, value=10, key=f"pts_{i}")
        negative = st.number_input("Negative (-)", min_value=0, max_value=50, value=2, key=f"neg_{i}")
        buzzer = st.checkbox("Buzzer Mode", value=False, key=f"buzz_{i}")

        rounds_config.append({
            "name": r_name,
            "count": q_count,
            "timer": timer,
            "points": points,
            "negative": negative,
            "buzzer": buzzer
        })

st.markdown("---")
st.header("2. Build Quiz Bundle")

if st.button("🚀 Generate Presentation & Printable PDFs", type="primary"):
    if not uploaded_pdfs and not uploaded_imgs:
        st.warning("Please upload at least one PDF or JPEG file in the sidebar.")
    else:
        with st.spinner("Extracting content and querying Gemini AI..."):
            extracted_text = extract_text_from_pdfs(uploaded_pdfs) if uploaded_pdfs else ""
            try:
                quiz_data = generate_questions_with_ai(rounds_config, extracted_text, uploaded_imgs or [])
                st.session_state['quiz_data'] = quiz_data
                st.success("Quiz bundle generated successfully!")
            except Exception as e:
                st.error(f"Error during generation: {str(e)}")

if 'quiz_data' in st.session_state:
    quiz_data = st.session_state['quiz_data']
    st.markdown("---")
    st.header("3. Download Deliverables")

    col1, col2, col3 = st.columns(3)

    with col1:
        pptx_buffer = create_pptx_deck(quiz_data, rounds_config)
        st.download_button(
            label="📊 Slide Deck (.pptx)",
            data=pptx_buffer,
            file_name="Quiz_Show_Deck.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )

    with col2:
        rulebook_pdf = create_rulebook_pdf(rounds_config)
        st.download_button(
            label="📑 Rulebook PDF (.pdf)",
            data=rulebook_pdf,
            file_name="Quiz_Master_Rulebook.pdf",
            mime="application/pdf",
            use_container_width=True
        )

    with col3:
        qa_pdf = create_qa_master_pdf(quiz_data)
        st.download_button(
            label="🖨️ Quiz Master QA PDF (.pdf)",
            data=qa_pdf,
            file_name="Quiz_Master_QA_Key.pdf",
            mime="application/pdf",
            use_container_width=True
        )

    st.subheader("Quiz Master Question & Answer Sheet Preview")
    for r in quiz_data.get("rounds", []):
        with st.expander(f"📖 {r['round_name']}"):
            for idx, q in enumerate(r["questions"]):
                st.write(f"**Q{idx+1}: {q['question']}**")
                if q.get("options"):
                    st.write(f"*Options:* {', '.join(q['options'])}")
                st.write(f"**Answer:** {q['answer']}")
                st.write(f"*Explanation:* {q['explanation']}")
                st.markdown("---")
