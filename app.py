"""Local web UI: python app.py  ->  http://127.0.0.1:5000"""
import shutil
import tempfile
import threading
import webbrowser
from pathlib import Path

from flask import Flask, after_this_request, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from converter import FORMATS, ConversionError, convert, has_text_layer

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB


@app.get("/")
def index():
    return render_template("index.html", formats=FORMATS)


@app.post("/convert")
def convert_route():
    file = request.files.get("file")
    fmt = request.form.get("format", "docx")
    if not file or not file.filename:
        return jsonify(error="No file uploaded."), 400
    if not file.filename.lower().endswith(".pdf"):
        return jsonify(error="Please upload a .pdf file."), 400
    if fmt not in FORMATS:
        return jsonify(error=f"Unknown format: {fmt}"), 400

    stem = Path(file.filename).stem or "document"
    safe_stem = secure_filename(stem) or "document"
    work = Path(tempfile.mkdtemp(prefix="pdfconv_"))

    @after_this_request
    def cleanup(response):
        shutil.rmtree(work, ignore_errors=True)
        return response

    src = work / f"{safe_stem}.pdf"
    file.save(src)
    password = request.form.get("password") or None
    try:
        out = convert(src, fmt, work / f"{safe_stem}{FORMATS[fmt]['ext']}",
                      request.form.get("pages"), password)
        scanned = fmt != "png" and not has_text_layer(src, password)
    except ConversionError as e:
        return jsonify(error=str(e)), 422

    # Read into memory so the temp folder can be removed on Windows right away.
    import io
    data = io.BytesIO(out.read_bytes())
    resp = send_file(data, as_attachment=True, download_name=f"{stem}{FORMATS[fmt]['ext']}")
    if scanned:
        resp.headers["X-Warning"] = "scanned"
    return resp


@app.errorhandler(413)
def too_large(_):
    return jsonify(error="File is too large (limit 200 MB)."), 413


if __name__ == "__main__":
    url = "http://127.0.0.1:5000"
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    print(f"file-converter running at {url}  (Ctrl+C to stop)")
    app.run(host="127.0.0.1", port=5000, debug=False)
