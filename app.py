#!/usr/bin/env python3
"""
PINNACLE LGS — Web Service for PDF & Excel generation
Called by Apps Script via HTTP POST
"""

from flask import Flask, request, send_file, jsonify
import os
import io
import base64
import tempfile
import json
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2 import service_account
from generator import build_pdf, build_excel
def get_drive_service():
    creds_json = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON')
    creds_dict = json.loads(creds_json)
    creds = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=['https://www.googleapis.com/auth/drive.file']
    )
    return build('drive', 'v3', credentials=creds)
app = Flask(__name__)

# Static assets: logo, signature, photos
ASSETS_DIR = os.path.join(os.path.dirname(__file__), "assets")

@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "service": "Pinnacle LGS PDF Generator",
        "status": "running",
        "endpoints": ["/generate-pdf", "/generate-excel", "/generate-both"]
    })

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "healthy"}), 200

@app.route("/generate-pdf", methods=["POST"])
def generate_pdf():
    try:
        data = request.get_json(force=True)
        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        tmp.close()
        build_pdf(data, tmp.name, os.path.basename(tmp.name))
        with open(tmp.name, "rb") as f:
            pdf_bytes = f.read()
        os.unlink(tmp.name)
        # Return base64 so Apps Script can decode easily
        return jsonify({
            "success": True,
            "pdf_base64": base64.b64encode(pdf_bytes).decode("ascii"),
            "filename": data.get("filename_pdf", "proforma.pdf")
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/generate-excel", methods=["POST"])
def generate_excel():
    try:
        data = request.get_json(force=True)
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        build_excel(data, tmp.name)
        with open(tmp.name, "rb") as f:
            xlsx_bytes = f.read()
        os.unlink(tmp.name)
        return jsonify({
            "success": True,
            "xlsx_base64": base64.b64encode(xlsx_bytes).decode("ascii"),
            "filename": data.get("filename_xlsx", "proforma.xlsx")
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/generate-both", methods=["POST"])
def generate_both():
    """Generates both PDF and Excel in one call (saves a round-trip)"""
    try:
        data = request.get_json(force=True)

        # PDF
        pdf_tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        pdf_tmp.close()
        build_pdf(data, pdf_tmp.name, data.get("filename_pdf", "proforma.pdf"))
        with open(pdf_tmp.name, "rb") as f:
            pdf_bytes = f.read()
        os.unlink(pdf_tmp.name)

        # Excel
        xlsx_tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        xlsx_tmp.close()
        build_excel(data, xlsx_tmp.name)
        with open(xlsx_tmp.name, "rb") as f:
            xlsx_bytes = f.read()
        os.unlink(xlsx_tmp.name)

        return jsonify({
            "success": True,
            "pdf_base64": base64.b64encode(pdf_bytes).decode("ascii"),
            "xlsx_base64": base64.b64encode(xlsx_bytes).decode("ascii"),
            "filename_pdf": data.get("filename_pdf", "proforma.pdf"),
            "filename_xlsx": data.get("filename_xlsx", "proforma.xlsx")
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
@app.route('/generate-and-upload', methods=['POST'])
def generate_and_upload():
    try:
        data = request.json
        folder_id = data.get('folder_id')
        filename  = data.get('filename', 'Proforma.pdf')

        # Génération PDF (même logique qu'existant)
        import tempfile
        pdf_tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        pdf_tmp.close()
        build_pdf(data, pdf_tmp.name)
        with open(pdf_tmp.name, "rb") as f:
            pdf_bytes = f.read()
        os.unlink(pdf_tmp.name)

        # Upload direct vers Drive
        service = get_drive_service()
        file_metadata = {'name': filename, 'parents': [folder_id]}
        media = MediaIoBaseUpload(
            io.BytesIO(pdf_bytes),
            mimetype='application/pdf',
            resumable=False
        )
        uploaded = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()

        return jsonify({'file_id': uploaded['id'], 'success': True})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
