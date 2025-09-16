from flask import Flask, request, send_file
import io, os, sys, tempfile, shutil, uuid, subprocess

app = Flask(__name__)

DEFAULT_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "default_template.pptx")

# --- Health: match "/" and "/api" (and optional trailing slash) ---
@app.get("/")
@app.get("/api")
def health():
    return {"status": "ok"}

# --- POST: match "/" and "/api" (works for both ways Vercel mounts the path) ---
@app.route("/", methods=["POST"])
def generate():
    if "excel" not in request.files:
        return ("Missing file: need 'excel'", 400)

    excel = request.files["excel"]
    ppt   = request.files.get("template")

    if not excel.filename.lower().endswith((".xlsx", ".xls")):
        return ("Excel must be .xlsx or .xls", 400)
    if ppt and ppt.filename and (not ppt.filename.lower().endswith(".pptx")):
        return ("Template must be .pptx", 400)

    if (not ppt or not ppt.filename) and not os.path.exists(DEFAULT_TEMPLATE_PATH):
        return ("Server template missing. Please add api/default_template.pptx to the repo.", 500)

    work = os.path.join(tempfile.gettempdir(), f"imarc_{uuid.uuid4().hex}")
    os.makedirs(work, exist_ok=True)
    try:
        excel_path = os.path.join(work, "datasheet_imarc.xlsx")
        ppt_path   = os.path.join(work, "template.pptx")

        excel.save(excel_path)
        if ppt and ppt.filename:
            ppt.save(ppt_path)
        else:
            shutil.copyfile(DEFAULT_TEMPLATE_PATH, ppt_path)

        script_src = os.path.join(os.path.dirname(__file__), "generate_poc.py")
        script_dst = os.path.join(work, "generate_poc.py")
        shutil.copyfile(script_src, script_dst)

        proc = subprocess.run(
            [sys.executable, script_dst],
            cwd=work,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=120,
        )
        if proc.returncode != 0:
            return (f"Script failed\nSTDOUT:\n{proc.stdout}\n\nSTDERR:\n{proc.stderr}", 500)

        out_path = os.path.join(work, "updated_poc.pptx")
        if not os.path.exists(out_path):
            return ("Output PPTX not found (expected 'updated_poc.pptx')", 500)

        with open(out_path, "rb") as f:
            data = f.read()
        return send_file(
            io.BytesIO(data),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name="updated_poc.pptx",
        )
    finally:
        try: shutil.rmtree(work)
        except Exception: pass
