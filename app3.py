from flask import Flask, request, render_template, send_file, send_from_directory
import os
import pandas as pd
import zipfile
from io import BytesIO
from urllib.parse import quote
from functions import (
    generate_cover_pdf,
    parse_via_pdf,
    fill_template,
    fill_conflict_docs_for_one,
    merge_custom_pages_by_index,
    paginate_pdf,
    is_name_match,
    STRENGTH_DATA
)

app = Flask(__name__)

# Define resource paths
BIG_TEMPLATE_PDF = os.path.join("resources", "bigTemplate.pdf")
CONFLICT_TEMPLATE_DOCX = os.path.join("resources", "Conflict_Template.docx")
SWEET_SPOT_TEMPLATE_DOCX = os.path.join("resources", "Sweet_Spot_Template.docx")

# Define output folder
OUTPUT_FOLDER = "output"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/", methods=["GET"])
def index():
    return render_template("upload.html")

@app.route("/generate", methods=["POST"])
def generate():
    mode = request.form.get("mode")
    template_version = request.form.get("template")  # Read the selected template version

    # Define the template file based on the selected version
    if template_version == "Open":
        template_pdf = os.path.join("resources", "bigTemplate.pdf")
    elif template_version == "Team":
        template_pdf = os.path.join("resources", "teamTemplate.pdf")
    elif template_version == "Tiny":
        template_pdf = os.path.join("resources", "tinyTemplate.pdf")
    else:
        return "Invalid template selected."

    if mode == "individual":
        # 1. Get form inputs
        participant_name = request.form.get("participantName").strip()
        term = request.form.get("date").strip()
        cohort = request.form.get("cohort").strip()
        
        # 2. Save uploaded files
        via_file = request.files["viaFile"]
        conflict_csv_file = request.files["conflictCSV"]

        via_filepath = os.path.join(OUTPUT_FOLDER, f"{participant_name}_via.pdf")
        conflict_csv_path = os.path.join(OUTPUT_FOLDER, f"{participant_name}_conflict.csv")

        via_file.save(via_filepath)
        conflict_csv_file.save(conflict_csv_path)

        # 3. Generate Cover Page
        cover_pdf = generate_cover_pdf(participant_name, term, cohort, OUTPUT_FOLDER)

        # 4. Parse VIA Survey
        parsed_name, results = parse_via_pdf(via_filepath)
        final_name = participant_name

        # 5. Fill Sweet Spot Template
        sweet_output_docx = os.path.join(OUTPUT_FOLDER, f"{final_name}_SweetSpot.docx")
        sweet_pdf = fill_template(
            results,
            STRENGTH_DATA,
            final_name,
            SWEET_SPOT_TEMPLATE_DOCX,
            sweet_output_docx
        )

        # 6. Process Conflict Resolution
        conflict_pdf = fill_conflict_docs_for_one(
            conflict_csv_path,
            CONFLICT_TEMPLATE_DOCX,
            OUTPUT_FOLDER,
            final_name
        )

        # 7. Merge PDFs
        merged_pdf = os.path.join(OUTPUT_FOLDER, f"{final_name.replace(' ', '_')}_merged.pdf")
        merge_custom_pages_by_index(
            template_pdf=template_pdf,  # Use the selected template
            cover_pdf=cover_pdf,
            via_pdf=via_filepath,
            sweet_pdf=sweet_pdf,
            conflict_pdf=conflict_pdf,
            output_pdf=merged_pdf
        )

        # 8. Paginate the Merged PDF
        final_workbook_pdf = os.path.join(OUTPUT_FOLDER, f"{final_name.replace(' ', '_')}_workbook.pdf")
        paginate_pdf(merged_pdf, final_workbook_pdf, start_page_index=3, start_page_number=3)

        # 9. Generate the report for individual mode
        report_html = generate_individual_report(final_name, final_workbook_pdf)
        
        # Render the report HTML directly in the browser
        return report_html

    elif mode == "batch":
        # Initialize a list to track generated files
        generated_files = []

        # 1. Get form inputs
        term = request.form.get("batchDate").strip()
        cohort = request.form.get("batchCohort").strip()
        
        # 2. Save uploaded files
        via_files = request.files.getlist("viaFiles")
        conflict_csv_file = request.files["conflictCSVBatch"]

        conflict_csv_path = os.path.join(OUTPUT_FOLDER, "batch_conflict.csv")
        conflict_csv_file.save(conflict_csv_path)

        # 3. Parse the CSV to get participant names
        df = pd.read_csv(conflict_csv_path)
        csv_names = set(df["First and Last Name"].str.strip().dropna().unique())

        # 4. Parse the VIA PDFs to get participant names
        pdf_names = {}
        for via_file in via_files:
            via_filepath = os.path.join(OUTPUT_FOLDER, via_file.filename)
            via_file.save(via_filepath)
            participant_name, _ = parse_via_pdf(via_filepath)
            pdf_names[via_file.filename] = participant_name

        # 5. Track mismatches and skipped participants
        matched_pairs = []
        missing_pdf = []
        missing_csv = []
        name_mismatches = []

        # 6. Match names between CSV and PDFs
        for csv_name in csv_names:
            matched = False
            for pdf_filename, pdf_name in pdf_names.items():
                if is_name_match(csv_name, pdf_name):
                    matched_pairs.append((csv_name, pdf_name, pdf_filename))
                    matched = True
                    break
            if not matched:
                missing_pdf.append(csv_name)

        for pdf_filename, pdf_name in pdf_names.items():
            matched = False
            for csv_name in csv_names:
                if is_name_match(csv_name, pdf_name):
                    matched = True
                    break
            if not matched:
                missing_csv.append(pdf_name)

        # 7. Generate workbooks for matched pairs
        for csv_name, pdf_name, pdf_filename in matched_pairs:
            via_filepath = os.path.join(OUTPUT_FOLDER, pdf_filename)
            conflict_pdf = fill_conflict_docs_for_one(
                conflict_csv_path,
                CONFLICT_TEMPLATE_DOCX,
                OUTPUT_FOLDER,
                csv_name
            )
            if not conflict_pdf:
                name_mismatches.append((csv_name, pdf_name))
                continue

            # Generate cover page
            cover_pdf = generate_cover_pdf(csv_name, term, cohort, OUTPUT_FOLDER)

            # Parse VIA PDF
            parsed_name, results = parse_via_pdf(via_filepath)

            # Fill Sweet Spot Template
            sweet_output_docx = os.path.join(OUTPUT_FOLDER, f"{csv_name.replace(' ', '_')}_SweetSpot.docx")
            sweet_pdf = fill_template(
                results,
                STRENGTH_DATA,
                csv_name,
                SWEET_SPOT_TEMPLATE_DOCX,
                sweet_output_docx
            )

            # Merge PDFs
            merged_pdf = os.path.join(OUTPUT_FOLDER, f"{csv_name.replace(' ', '_')}_merged.pdf")
            merge_custom_pages_by_index(
                template_pdf=template_pdf,  # Use the selected template
                cover_pdf=cover_pdf,
                via_pdf=via_filepath,
                sweet_pdf=sweet_pdf,
                conflict_pdf=conflict_pdf,
                output_pdf=merged_pdf
            )

            # Paginate the Merged PDF
            final_workbook_pdf = os.path.join(OUTPUT_FOLDER, f"{csv_name.replace(' ', '_')}_workbook.pdf")
            paginate_pdf(merged_pdf, final_workbook_pdf, start_page_index=3, start_page_number=3)

            # Add the generated workbook to the list
            generated_files.append(final_workbook_pdf)

        # 8. Generate the report for batch mode
        report_html = generate_report(matched_pairs, missing_pdf, missing_csv, name_mismatches, generated_files)
        
        # Render the report HTML directly in the browser
        return report_html
    else:
        return "Invalid mode selected."


def generate_individual_report(participant_name, workbook_path):
    """
    Generates an HTML report for individual mode.
    """
    file_name = os.path.basename(workbook_path)
    download_link = f"/download_file/{quote(file_name)}"

    report = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>Workbook Generation Report</title>
        <style>
            body {{ font-family: Arial, sans-serif; }}
            h1 {{ color: #333; }}
            .success {{ color: green; }}
            .download-link {{ margin-top: 20px; }}
            .back-button {{ margin-top: 20px; }}
        </style>
    </head>
    <body>
        <h1>Workbook Generation Report</h1>
        <p class="success">Workbook for <strong>{participant_name}</strong> generated successfully!</p>
        <div class="download-link">
            <a href='{download_link}'><button>Download Workbook</button></a>
        </div>
        <div class="back-button">
            <a href='/'><button>Generate More / Return</button></a>
        </div>
    </body>
    </html>
    """
    return report

def generate_report(matched_pairs, missing_pdf, missing_csv, name_mismatches, generated_files):
    """
    Generates an HTML report summarizing the batch processing results.
    """
    # Encode file paths for the "Download All" link
    encoded_files = [quote(file_path) for file_path in generated_files]
    download_all_link = f"/download_all?files={'&files='.join(encoded_files)}"

    report = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>Batch Processing Report</title>
        <style>
            body {{ font-family: Arial, sans-serif; }}
            h1 {{ color: #333; }}
            .section {{ margin-bottom: 20px; }}
            .section h2 {{ color: #555; }}
            ul {{ list-style-type: none; padding: 0; }}
            li {{ margin: 5px 0; }}
            .success {{ color: green; }}
            .warning {{ color: orange; }}
            .error {{ color: red; }}
            .download-all {{ margin-top: 20px; }}
            .back-button {{ margin-top: 20px; }}
        </style>
    </head>
    <body>
        <h1>Batch Processing Report</h1>
        
        <div class="section">
            <h2>Successfully Generated Workbooks</h2>
            <ul>
    """
    for csv_name, pdf_name, _ in matched_pairs:
        report += f"<li class='success'>{csv_name} (matched with {pdf_name})</li>"

    report += """
            </ul>
        </div>
        
        <div class="section">
            <h2>Participants Missing PDFs</h2>
            <ul>
    """
    for name in missing_pdf:
        report += f"<li class='warning'>{name}</li>"

    report += """
            </ul>
        </div>
        
        <div class="section">
            <h2>PDFs Missing CSV Entries</h2>
            <ul>
    """
    for name in missing_csv:
        report += f"<li class='warning'>{name}</li>"

    report += """
            </ul>
        </div>
        
        <div class="section">
            <h2>Name Mismatches</h2>
            <ul>
    """
    for csv_name, pdf_name in name_mismatches:
        report += f"<li class='error'>{csv_name} (CSV) vs. {pdf_name} (PDF)</li>"

    report += f"""
            </ul>
        </div>

        <div class="section">
            <h2>Download Generated Workbooks</h2>
            <ul>
    """
    for file_path in generated_files:
        file_name = os.path.basename(file_path)
        report += f"<li><a href='/download_file/{quote(file_name)}'>{file_name}</a></li>"

    report += f"""
            </ul>
            <div class="download-all">
                <a href='{download_all_link}'><button>Download All as ZIP</button></a>
            </div>
        </div>
        <div class="back-button">
            <a href='/'><button>Generate More / Return</button></a>
        </div>
    </body>
    </html>
    """
    return report

@app.route("/download_file/<filename>")
def download_file(filename):
    """
    Allows users to download a specific generated workbook.
    """
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

@app.route("/download_all")
def download_all():
    """
    Allows users to download all generated workbooks as a ZIP file.
    """
    # Get the list of generated files from the request arguments
    encoded_files = request.args.getlist("files")
    generated_files = [os.path.join(OUTPUT_FOLDER, os.path.basename(file)) for file in encoded_files]

    # Create a BytesIO object to hold the ZIP file in memory
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for file_path in generated_files:
            if os.path.exists(file_path):
                file_name = os.path.basename(file_path)
                zip_file.write(file_path, arcname=file_name)
            else:
                print(f"File not found: {file_path}")

    # Move the buffer's pointer to the beginning
    zip_buffer.seek(0)

    # Return the ZIP file as a response
    return send_file(
        zip_buffer,
        mimetype="application/zip",
        as_attachment=True,
        download_name="workbooks.zip"
    )

if __name__ == "__main__":
    app.run(debug=True)