from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from pdf2docx import Converter
import tempfile, os, uuid

app = Flask(__name__)
CORS(app, origins=["https://classedobsdecath.ch", "http://localhost"])

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"})

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({"error": "Aucun fichier"}), 400
    
    pdf_file = request.files['file']
    if not pdf_file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Format PDF requis"}), 400
    
    uid = str(uuid.uuid4())[:8]
    pdf_path = f"/tmp/{uid}.pdf"
    docx_path = f"/tmp/{uid}.docx"
    
    try:
        pdf_file.save(pdf_path)
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        return send_file(
            docx_path,
            as_attachment=True,
            download_name=pdf_file.filename.replace('.pdf', '.docx'),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        for p in [pdf_path, docx_path]:
            if os.path.exists(p):
                os.remove(p)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
