import re
import os
import tempfile
from datetime import datetime
from flask import Flask, render_template, request, send_file, after_this_request
from docx import Document
from docx.oxml import OxmlElement

app = Flask(__name__)

# Path to your Word document template
TEMPLATE_PATH = os.path.join('templates', 'MOM_template.docx')


# ✅ Extract meeting date and agenda points from email content
def extract_meeting_details(text):
    date_match = re.search(r"scheduled on ([A-Za-z]+, [A-Za-z]+ \d{1,2}, \d{4})", text)
    if date_match:
        raw_date = date_match.group(1)
        try:
            clean_date = " ".join(raw_date.split()[1:])  # Remove weekday
            parsed_date = datetime.strptime(clean_date, "%B %d, %Y")
            meeting_date = parsed_date.strftime("%B %d, %Y")
        except Exception:
            meeting_date = raw_date
    else:
        meeting_date = "Unknown Date"

    agenda_section = re.search(r"(1\..*?)(?:\n\n|\Z)", text, re.DOTALL)
    agenda_text = agenda_section.group(1).strip() if agenda_section else ""
    agenda_points = re.findall(r"\d+\..*", agenda_text)

    return meeting_date, agenda_points


# ✅ Replace placeholder in header
def replace_header_text(doc, placeholder, replacement):
    for section in doc.sections:
        header = section.header
        for paragraph in header.paragraphs:
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, replacement)


# ✅ Generate final Word document
def generate_doc(meeting_date, agenda_points):
    doc = Document(TEMPLATE_PATH)

    # Replace MEETING_DATE and MEETING_DATE_u placeholders
    for paragraph in doc.paragraphs:
        if "<<MEETING_DATE>>" in paragraph.text:
            replaced = paragraph.text.replace("<<MEETING_DATE>>", meeting_date)
            paragraph.clear()
            paragraph.add_run(replaced)

        if "<<MEETING_DATE_u>>" in paragraph.text:
            replaced = paragraph.text.replace("<<MEETING_DATE_u>>", meeting_date)
            paragraph.clear()
            run = paragraph.add_run(replaced)
            run.underline = True

    # Replace MEETING_DATE in header if used
    replace_header_text(doc, "<<MEETING_DATE>>", meeting_date)

    # Insert agenda points (bold) and remove <<AGENDA>> placeholder
    for paragraph in doc.paragraphs:
        if "<<AGENDA>>" in paragraph.text:
            for agenda in reversed(agenda_points):  # reverse to keep correct order when inserting before
                new_para = paragraph.insert_paragraph_before("")
                run = new_para.add_run(agenda)
                run.bold = True
            # Remove the <<AGENDA>> placeholder paragraph
            p = paragraph._element
            p.getparent().remove(p)
            break

    # Save to temporary file
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(temp_file.name)
    return temp_file.name, meeting_date


# ✅ Flask Routes
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        email_text = request.form['email_text']
        file_path, meeting_date = generate_doc(*extract_meeting_details(email_text))

        @after_this_request
        def cleanup(response):
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except Exception as e:
                print(f"Error deleting temp file: {e}")
            return response

        filename = f"MoM of the MBA Committee meeting dated {meeting_date}.docx"
        return send_file(file_path, as_attachment=True, download_name=filename)

    return render_template("index.html")


# ✅ Run the Flask App
if __name__ == "__main__":
    app.run(debug=True)
