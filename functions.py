import re
from pypdf import PdfReader, PdfWriter
import os
import subprocess
from docxtpl import DocxTemplate
import pandas as pd
from fuzzywuzzy import fuzz

# Fuzzy matching helper function
def is_name_match(name1, name2, threshold=80):
    """
    Compare two names using fuzzy matching.
    Returns True if the similarity score is above the threshold.
    """
    return fuzz.ratio(name1, name2) >= threshold

SCORE_MAP = {
    "Rarely": 1,
    "Sometimes": 2,
    "Often": 3,
    "Always": 4
}

QUESTION_CATEGORIES = {
    "I discuss issues with others to try to find solutions that meet everyone's needs.": "Collaborating",
    "I try to negotiate and use a give-and-take approach to problem situations.": "Compromising",
    "I try to meet the expectations of others.": "Accommodating",
    "I would argue my case and insist on the advantages of my point of view.": "Competing",
    "When there is a disagreement, I gather as much information as I can and keep the lines of communication open.": "Collaborating",
    "When I find myself in an argument, I usually say very little and try to leave as soon as possible.": "Avoiding",
    "I try to see conflicts from both sides. What do I need? What does the other person need? What are the issues involved?": "Collaborating",
    "I prefer to compromise when solving problems and just move on.": "Compromising",
    "I find conflicts exhilarating; I enjoy the battle of wits that usually follows.": "Competing",
    "Being in a disagreement with other people makes me feel uncomfortable and anxious.": "Avoiding",
    "I try to meet the wishes of my friends and family.": "Accommodating",
    "I can figure out what needs to be done and I am usually right.": "Competing",
    "To break deadlocks, I would meet people halfway.": "Compromising",
    "I may not get what I want but its a small price to pay for keeping the peace.": "Accommodating",
    "I avoid hard feelings by keeping my disagreements with others to myself.": "Avoiding",
}

# Strengths dictionary
STRENGTH_DATA = {
    "Spirituality": {
        "underuse": "lack of purpose; disconnected from sacred",
        "optimal": "finding purpose; pursuing life meaning/connecting with sacred",
        "overuse": "fanatical; preachy/rigid beliefs"
    },
    "Gratitude": {
        "underuse": "entitled; unappreciative",
        "optimal": "positive expectations; optimistic",
        "overuse": "dependent; blind acceptance/loss of individuality"
    },
    "Hope": {
        "underuse": "apathy; pessimistic despair",
        "optimal": "positive expectations; optimistic",
        "overuse": "delusional positivity; ignoring problems"
    },
    "Humor": {
        "underuse": "grim; unapproachable",
        "optimal": "healthy levity; group-oriented",
        "overuse": "excessive teasing; belittling"
    },
    "Kindness": {
        "underuse": "aloof; selfish",
        "optimal": "compassion; empathy in action",
        "overuse": "martyrdom; compassion fatigue"
    },
    "Love": {
        "underuse": "disconnected; lonely",
        "optimal": "warmth and closeness with others",
        "overuse": "clinging; ignoring personal boundaries"
    },
    "Bravery": {
        "underuse": "fear-driven; easily intimidated",
        "optimal": "standing up for beliefs; persevering through adversity",
        "overuse": "reckless risk-taking"
    },
    "Curiosity": {
        "underuse": "uninterested; apathetic",
        "optimal": "information seeking; exploration",
        "overuse": "scattered focus; superficial dabbling"
    },
    "Love Of Learning": {
        "underuse": "disengaged with knowledge",
        "optimal": "intentional learning; open minded",
        "overuse": "analysis paralysis; ignoring practicality"
    },
    "Perspective": {
        "underuse": "unaware; limited viewpoint",
        "optimal": "wisdom-based insight; broad perspective",
        "overuse": "overthinking; constant re-evaluation"
    },
    "Creativity": {
        "underuse": "uninspired; stuck thinking",
        "optimal": "imaginative solutions; innovative",
        "overuse": "unrealistic; ignoring constraints"
    },
    "Judgment": {
        "underuse": "uncritical acceptance; naive",
        "optimal": "thoughtful consideration; balanced reasoning",
        "overuse": "hypercritical; indecisive"
    },
    "Zest": {
        "underuse": "low energy; indifferent",
        "optimal": "enthusiasm; active engagement",
        "overuse": "impulsivity; burnout from overcommitment"
    },
    "Perseverance": {
        "underuse": "easily give up; no follow-through",
        "optimal": "steadfast pursuit of goals; resilience",
        "overuse": "stubbornness; ignoring diminishing returns"
    },
    "Honesty": {
        "underuse": "deception; lack of authenticity",
        "optimal": "authentic self-expression; responsibility",
        "overuse": "bluntness; ignoring tact or empathy"
    },
    "Leadership": {
        "underuse": "lack of direction; passive group involvement",
        "optimal": "guiding vision; collaborative organization",
        "overuse": "domineering; micromanagement"
    },
    "Teamwork": {
        "underuse": "isolated; lacking group synergy",
        "optimal": "cooperative effort; shared goals",
        "overuse": "groupthink; conformity"
    },
    "Fairness": {
        "underuse": "bias; partial treatment",
        "optimal": "equitable decisions; impartial justice",
        "overuse": "inflexible adherence to rules over context"
    },
    "Forgiveness": {
        "underuse": "resentful; vengeful",
        "optimal": "letting go of grudges; understanding",
        "overuse": "enabling repeated harm; ignoring boundaries"
    },
    "Humility": {
        "underuse": "arrogance; self-centeredness",
        "optimal": "accurate self-view; respectful of others",
        "overuse": "self-effacing; inability to accept credit"
    },
    "Prudence": {
        "underuse": "impulsive; risky decisions",
        "optimal": "thoughtful planning; caution",
        "overuse": "overly cautious; fear of risk"
    },
    "Self-Regulation": {
        "underuse": "indulgent; lacking discipline",
        "optimal": "balanced habits; emotional control",
        "overuse": "rigidity; denying basic needs"
    },
    "Appreciation Of Beauty & Excellence": {
        "underuse": "oblivious; uninterested in excellence",
        "optimal": "valuing artistry, skill, or beauty",
        "overuse": "hyperfocus on perfection; aesthetic snobbery"
    },
    "Social Intelligence": {
        "underuse": "clueless about social cues; insensitive",
        "optimal": "aware of social dynamics; empathetic communication",
        "overuse": "manipulative; overthinking interactions"
    }
}

# -------------------------------
# PDF Extraction using pdfminer and PyMuPDF
# -------------------------------
from pdfminer.high_level import extract_text

import fitz  # PyMuPDF

def parse_via_pdf(pdf_path):
    print(f"Reading PDF using PyMuPDF from: {pdf_path}")
    doc = fitz.open(pdf_path)
    full_text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        print(f"--- Page {page_num + 1} ---")
        print(text)
        full_text += text + "\n"
    doc.close()

    print("\n=== Full Extracted Text ===")
    print(full_text)
    print("===========================\n")

    # Extract participant name
    name_match = re.search(r"^(.*?)\nVIA Character Strengths Profile", full_text, re.MULTILINE)
    if name_match:
        person_name = name_match.group(1).strip()
        person_name = re.sub(r'\s+', ' ', person_name)
    else:
        person_name = "Unknown"

    # Extract strengths e.g. "1. Humor"
    pattern = re.compile(r"(\d+)\.\s+(.+)")
    matches = pattern.findall(full_text)
    results = [(int(rank), strength.strip()) for rank, strength in matches]

    print(f"Extracted Name: {person_name}")
    print("Extracted Strengths:")
    for rank, strength in results:
        print(f"{rank}: {strength}")

    return person_name, results

# -------------------------------
# DOCX to PDF Conversion using LibreOffice
# -------------------------------
def convert_to_pdf_via_libreoffice(docx_path, output_dir=None):
    if output_dir is None:
        output_dir = os.path.dirname(docx_path) or "."
    
    # Use the 'libreoffice' command which should be available in Render's environment via shell.nix
    command = [
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        docx_path,
        "--outdir", output_dir
    ]
    subprocess.run(command, check=True)
    pdf_path = os.path.join(output_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    return pdf_path

# -------------------------------
# Fill Sweet Spot Template DOCX and Convert to PDF
# -------------------------------
def fill_template(parsed_strengths, strength_data, person_name, template_path, output_docx_path):
    """
    Fills the Sweet Spot Template with the parsed strengths and definitions,
    then converts the filled DOCX to a PDF.
    """
    context = {}
    context["name"] = person_name
    for i in range(24):
        placeholder_index = i + 1
        if i < len(parsed_strengths):
            _, strength = parsed_strengths[i]
            strength_title = strength.strip().title()
            context[f"strength{placeholder_index}"] = strength_title
            if strength_title in strength_data:
                context[f"underuse{placeholder_index}"] = strength_data[strength_title]["underuse"]
                context[f"optimal{placeholder_index}"] = strength_data[strength_title]["optimal"]
                context[f"overuse{placeholder_index}"] = strength_data[strength_title]["overuse"]
            else:
                context[f"underuse{placeholder_index}"] = ""
                context[f"optimal{placeholder_index}"] = ""
                context[f"overuse{placeholder_index}"] = ""
        else:
            context[f"strength{placeholder_index}"] = ""
            context[f"underuse{placeholder_index}"] = ""
            context[f"optimal{placeholder_index}"] = ""
            context[f"overuse{placeholder_index}"] = ""
    
    doc = DocxTemplate(template_path)
    doc.render(context)
    doc.save(output_docx_path)
    print(f"Template has been filled and saved as: {output_docx_path}")
    
    pdf_output_path = convert_to_pdf_via_libreoffice(output_docx_path)
    print(f"Converted to PDF: {pdf_output_path}")
    return pdf_output_path

# -------------------------------
# Conflict Resolution DOCX Generation for Batch
# -------------------------------
def fill_conflict_docs(csv_path, template_path, output_dir="."):
    """
    Processes the conflict CSV for each respondent, fills a DOCX template, converts it to PDF,
    and returns a list of participant names.
    """
    df = pd.read_csv(csv_path)
    participant_names = []
    for idx, row in df.iterrows():
        full_name = str(row["First and Last Name"]).strip()
        if pd.isna(full_name) or full_name == "":
            continue
        participant_names.append(full_name)
        category_scores = {category: 0 for category in QUESTION_CATEGORIES.values()}
        for question_col, category in QUESTION_CATEGORIES.items():
            if question_col in df.columns:
                answer_text = str(row[question_col]).strip()
                numeric_score = SCORE_MAP.get(answer_text, 0)
                category_scores[category] += numeric_score
        context = {
            "name": full_name,
            "Col": category_scores["Collaborating"],
            "Com": category_scores["Competing"],
            "Avo": category_scores["Avoiding"],
            "Acc": category_scores["Accommodating"],
            "Co2": category_scores["Compromising"],
        }
        safe_name = full_name.replace(" ", "_")
        output_filename = f"{safe_name}_ConflictStyle3.docx"
        output_path = os.path.join(output_dir, output_filename)
        doc = DocxTemplate(template_path)
        doc.render(context)
        doc.save(output_path)
        pdf_output_path = convert_to_pdf_via_libreoffice(output_path)
        os.remove(output_path)
    return participant_names

# -------------------------------
# Conflict Resolution DOCX for a Single Participant
# -------------------------------
def fill_conflict_docs_for_one(csv_path, template_path, output_dir, participant_name):
    import os
    import pandas as pd
    from docxtpl import DocxTemplate

    df = pd.read_csv(csv_path)
    filtered_df = df[df["First and Last Name"] == participant_name]
    if filtered_df.empty:
        print(f"No responses found for {participant_name} in {csv_path}")
        return
    row = filtered_df.iloc[0]
    full_name = str(row["First and Last Name"]).strip()
    category_scores = {
        "Collaborating": 0,
        "Compromising": 0,
        "Avoiding": 0,
        "Accommodating": 0,
        "Competing": 0
    }
    for question_col, category in QUESTION_CATEGORIES.items():
        if question_col in df.columns:
            answer_text = str(row[question_col]).strip()
            numeric_score = SCORE_MAP.get(answer_text, 0)
            category_scores[category] += numeric_score
        else:
            print(f"Warning: '{question_col}' not found in CSV columns.")
    context = {
        "name": full_name,
        "Col": category_scores["Collaborating"],
        "Com": category_scores["Competing"],
        "Avo": category_scores["Avoiding"],
        "Acc": category_scores["Accommodating"],
        "Co2": category_scores["Compromising"],
    }
    doc = DocxTemplate(template_path)
    doc.render(context)
    safe_name = full_name.replace(" ", "_")
    output_filename = f"{safe_name}_ConflictStyle3.docx"
    output_path = os.path.join(output_dir, output_filename)
    doc.save(output_path)
    print(f"Saved DOCX: {output_path}")
    pdf_output_path = convert_to_pdf_via_libreoffice(output_path, output_dir)
    print(f"Converted to PDF: {pdf_output_path}")
    os.remove(output_path)
    return pdf_output_path

# -------------------------------
# PDF Merging Function
# -------------------------------
def merge_custom_pages_by_index(template_pdf, cover_pdf, via_pdf, sweet_pdf, conflict_pdf, output_pdf):
    """
    Merges PDFs by replacing specific pages in template_pdf with custom PDFs.
    - Page 0 -> cover_pdf
    - Page 4 -> via_pdf
    - Page 8 -> sweet_pdf
    - Page 11 -> conflict_pdf
    """
    writer = PdfWriter()
    template_reader = PdfReader(template_pdf)
    cover_reader    = PdfReader(cover_pdf)
    via_reader      = PdfReader(via_pdf)
    sweet_reader    = PdfReader(sweet_pdf)
    conflict_reader = PdfReader(conflict_pdf)
    for i in range(len(template_reader.pages)):
        if i == 0:
            for cp in cover_reader.pages:
                writer.add_page(cp)
        elif i == 4:
            for vp in via_reader.pages:
                writer.add_page(vp)
        elif i == 8:
            for sp in sweet_reader.pages:
                writer.add_page(sp)
        elif i == 11:
            for cr in conflict_reader.pages:
                writer.add_page(cr)
        else:
            writer.add_page(template_reader.pages[i])
    with open(output_pdf, "wb") as out:
        writer.write(out)
    print(f"Merged PDF created: {output_pdf}")

# -------------------------------
# Page Numbering Functions
# -------------------------------
from io import BytesIO
from reportlab.pdfgen import canvas

def create_page_number_overlay(page_width, page_height, page_number, margin=36):
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))
    c.setFont("Times-Roman", 10)
    text = str(page_number)
    text_width = c.stringWidth(text, "Times-Roman", 10)
    x = page_width - margin - text_width
    y = margin
    c.drawString(x, y, text)
    c.save()
    packet.seek(0)
    overlay_reader = PdfReader(packet)
    return overlay_reader.pages[0]

def paginate_pdf(input_pdf, output_pdf, start_page_index=3, start_page_number=3):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    num_pages = len(reader.pages)
    for i in range(num_pages):
        page = reader.pages[i]
        if i >= start_page_index:
            page_number = start_page_number + (i - start_page_index)
            page_width = float(page.mediabox.upper_right[0])
            page_height = float(page.mediabox.upper_right[1])
            overlay = create_page_number_overlay(page_width, page_height, page_number)
            page.merge_page(overlay)
        writer.add_page(page)
    with open(output_pdf, "wb") as f:
        writer.write(f)
    print(f"Paginated PDF saved as: {output_pdf}")

# -------------------------------
# VIA Survey Processing Function for Individual Mode
# -------------------------------
def process_via_survey(pdf_path, strength_data, template_path, output_folder):
    """
    Processes the VIA survey PDF by extracting the participant's name and strengths,
    filling the Sweet Spot template, and converting to PDF.
    """
    person_name, parsed_strengths = parse_via_pdf(pdf_path)
    safe_name = person_name.replace(" ", "_")
    output_docx_path = os.path.join(output_folder, f"{safe_name}_SweetSpot.docx")
    sweet_spot_pdf = fill_template(parsed_strengths, strength_data, person_name, template_path, output_docx_path)
    return sweet_spot_pdf

# -------------------------------
# Cover Page Generation Function
# -------------------------------
def generate_cover_pdf(participant_name=None, date=None, cohort=None, output_folder="."):
    """
    Generates a customized cover page PDF using a DOCX cover template.
    """
    cover_template_path = os.path.join("resources", "coverTemplate.docx")
    safe_name = participant_name.replace(" ", "_")
    output_docx_path = os.path.join(output_folder, f"{safe_name}_Cover.docx")
    context = {
        "name": participant_name,
        "date": date,
        "cohort": cohort
    }
    doc = DocxTemplate(cover_template_path)
    doc.render(context)
    doc.save(output_docx_path)
    print(f"Cover DOCX saved as: {output_docx_path}")
    cover_pdf = convert_to_pdf_via_libreoffice(output_docx_path, output_folder)
    print(f"Cover PDF saved as: {cover_pdf}")
    os.remove(output_docx_path)
    print(f"Intermediate DOCX file {output_docx_path} deleted.")
    return cover_pdf
