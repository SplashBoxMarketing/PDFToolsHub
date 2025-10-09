from flask import Flask, render_template, request, send_file, redirect, url_for, abort
from logic.billtracker import build_excel
from logic.toc_linker import process_pdf as process_toc
import os

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
    pdf = request.files.get("pdf")
    rng = request.form.get("range", "").strip()
    if not pdf or not rng:
        return redirect(url_for("toc_get"))
    out = process_toc(pdf, rng)
    name = (pdf.filename or "output.pdf").rsplit(".", 1)[0] + "_TOCLinked.pdf"
    return send_file(out, as_attachment=True, download_name=name)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=True)
