from flask import Flask, render_template, request, redirect, send_from_directory, send_file, flash
import os
import pandas as pd
from docx import Document
from werkzeug.utils import secure_filename
from docx2pdf import convert
from email.message import EmailMessage
import smtplib
import zipfile
from io import BytesIO
import unicodedata
import re

from docx.shared import Inches

app = Flask(__name__)
app.secret_key = 'supersecret'

app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'
app.config['PDF_FOLDER'] = 'pdfs'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
os.makedirs(app.config['PDF_FOLDER'], exist_ok=True)


def render_docx_template(doc: Document, context: dict) -> Document:
    for para in doc.paragraphs:
        for key, val in context.items():
            if f"{{{{{key}}}}}" in para.text:
                for run in para.runs:
                    run.text = run.text.replace(f"{{{{{key}}}}}", str(val))
    return doc


def send_email_with_attachment(receiver, pdf_path, sender, password, subject, body):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = receiver
    msg.set_content(body)

    with open(pdf_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(pdf_path)
        msg.add_attachment(file_data, maintype="application", subtype="pdf", filename=file_name)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender, password)
        smtp.send_message(msg)


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        excel_file = request.files["excel"]
        word_file = request.files["template"]
        sender_email = request.form.get("sender_email")
        sender_pass = request.form.get("sender_pass")
        email_subject = request.form.get("email_subject")
        email_body = request.form.get("email_body")

        if not excel_file or not word_file:
            flash("Veuillez sélectionner les deux fichiers.")
            return redirect("/")

        excel_filename = secure_filename(excel_file.filename)
        word_filename = secure_filename(word_file.filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
        word_path = os.path.join(app.config['UPLOAD_FOLDER'], word_filename)

        excel_file.save(excel_path)
        word_file.save(word_path)

        
        os.replace(excel_path, os.path.join(app.config['UPLOAD_FOLDER'], "dernier_fichier.xlsx"))

        df = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'], "dernier_fichier.xlsx"))

        output_files = []
        for _, row in df.iterrows():
            context = {col.lower(): row[col] for col in df.columns}
            context["exercice"] = "2024-2025"
            doc = Document(word_path)
            filled_doc = render_docx_template(doc, context)

            fname_base = f"Fiche_{context.get('prenom', 'nom')}_{context.get('nom', '')}".replace(" ", "_")
            docx_path = os.path.join(app.config['OUTPUT_FOLDER'], fname_base + ".docx")
            pdf_path = os.path.join(app.config['PDF_FOLDER'], fname_base + ".pdf")
            filled_doc.save(docx_path)

            try:
                convert(docx_path, pdf_path)
            except Exception as e:
                print("Erreur conversion PDF:", e)

            output_files.append(fname_base + ".pdf")

        return render_template("result.html", files=output_files,
                               sender_email=sender_email,
                               sender_pass=sender_pass,
                               email_subject=email_subject,
                               email_body=email_body)

    return render_template("index.html")


@app.route("/send_all", methods=["POST"])
def send_all():
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], "dernier_fichier.xlsx")
    df = pd.read_excel(excel_path)

    
    sender_email = "joeamikaze@gmail.com"
    sender_pass = "gjox fxqm npej duyl"

    
    email_subject = "Votre fiche PDF"
    email_body = "Bonjour, \n\nVeuillez trouver ci-joint votre document PDF.\n\nCordialement."

    for _, row in df.iterrows():
        context = {col.lower(): row[col] for col in df.columns}
        context["exercice"] = "2024-2025"
        fname_base = f"Fiche_{context.get('prenom', 'nom')}_{context.get('nom', '')}".replace(" ", "_")
        pdf_path = os.path.join(app.config['PDF_FOLDER'], fname_base + ".pdf")

        recipient_email = context.get("email")  # ← doit correspondre exactement au nom de la colonne

        if recipient_email and pd.notna(recipient_email) and os.path.exists(pdf_path):
            try:
                send_email_with_attachment(
                    recipient_email,
                    pdf_path,
                    sender_email,
                    sender_pass,
                    email_subject,
                    email_body
                )
                print(f"✔️ Envoyé à {recipient_email}")
            except Exception as e:
                print(f"❌ Erreur envoi à {recipient_email}:", e)

    flash("📨 Tous les fichiers ont été envoyés par e-mail avec succès.", "success")
    return redirect("/")

@app.route("/download/<filename>")
def download(filename):
    return send_from_directory(app.config['PDF_FOLDER'], filename, as_attachment=True)


@app.route('/download_all.zip')
def download_all():
    files = [f for f in os.listdir(app.config['PDF_FOLDER']) if f.endswith('.pdf')]
    if not files:
        return "Aucun fichier PDF à télécharger.", 404

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        for filename in files:
            file_path = os.path.join(app.config['PDF_FOLDER'], filename)
            safe_name = clean_filename(filename)
            zip_file.write(file_path, arcname=safe_name)

    zip_buffer.seek(0)
    return send_file(
        zip_buffer,
        mimetype='application/zip',
        as_attachment=True,
        download_name='fiches_pdf.zip'
    )


def clean_filename(filename):
    try:
        filename.encode('utf-8')
        return filename
    except UnicodeEncodeError:
        nfkd_form = unicodedata.normalize('NFKD', filename)
        clean_name = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
        clean_name = re.sub(r'[^\w\s.-]', '', clean_name)
        return clean_name


if __name__ == "__main__":
    app.run(debug=True)
