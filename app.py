from flask import Flask, render_template, request, send_file, redirect, url_for, abort
from logic.billtracker import build_excel
from logic.toc_linker import process_pdf as process_toc
import os
import io, zipfile
from flask import jsonify

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = (
    100 * 1024 * 1024
)  # 100 MB total per request; adjust as needed

# Optional very simple shared password
APP_PASSWORD = os.getenv("APP_PASSWORD", "")


def require_auth(fn):
    def wrapper(*args, **kwargs):
        if APP_PASSWORD:
            if request.method == "POST":
                if request.form.get("password") != APP_PASSWORD:
                    abort(401)
            else:
                # GET: show login if not provided (very lightweight)
                pass
        return fn(*args, **kwargs)

    wrapper.__name__ = fn.__name__
    return wrapper


@app.post("/billtracker-json")
def bill_post_json():
    data = request.get_json(silent=True) or {}
    rows = data.get("rows", [])
    if not rows:
        return jsonify({"error": "No rows provided"}), 400

    # Build the Excel from provided rows (no PDFs uploaded)
    out = build_excel_from_rows(rows)
    return send_file(out, as_attachment=True, download_name="pdf_page_counts.xlsx")


@app.get("/")
def index():
    return render_template("index.html", require_pw=bool(APP_PASSWORD))


@app.get("/billtracker")
def bill_get():
    return render_template("billtracker.html", require_pw=bool(APP_PASSWORD))


@app.post("/billtracker")
@require_auth
def bill_post():
    files = request.files.getlist("pdfs")
    if not files:
        return redirect(url_for("bill_get"))
    out = build_excel(files)
    return send_file(out, as_attachment=True, download_name="pdf_page_counts.xlsx")


@app.get("/toc-linker")
def toc_get():
    return render_template("toc_linker.html", require_pw=bool(APP_PASSWORD))


@app.post("/toc-linker")
@require_auth
def toc_post():
    # Gather all uploaded PDFs (now multiple supported)
    pdfs = request.files.getlist("pdfs")
    rng = request.form.get("range", "").strip()

    if not pdfs or not rng:
        return redirect(url_for("toc_get"))

    # If only one file, keep the simple single-PDF response
    if len(pdfs) == 1:
        pdf = pdfs[0]
        out = process_toc(pdf, rng)
        name = (pdf.filename or "output.pdf").rsplit(".", 1)[0] + "_TOCLinked.pdf"
        return send_file(out, as_attachment=True, download_name=name)

    # For multiple files, process each and return a ZIP
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for pdf in pdfs:
            if not pdf or not pdf.filename:
                continue
            try:
                out = process_toc(pdf, rng)
                out_name = pdf.filename.rsplit(".", 1)[0] + "_TOCLinked.pdf"
                zf.writestr(out_name, out.getvalue())
            except Exception as e:
                # If a file fails, include a small .txt error entry so the user knows which one failed
                err_name = pdf.filename.rsplit(".", 1)[0] + "_ERROR.txt"
                zf.writestr(err_name, f"Failed to process: {pdf.filename}\n{e}")

    zip_buf.seek(0)
    return send_file(
        zip_buf,
        as_attachment=True,
        download_name="TOCLinked_Batch.zip",
        mimetype="application/zip",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=True)
