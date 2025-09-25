import io, os, sys, tempfile, shutil, uuid, subprocess, traceback
from flask import Flask, request, send_file, make_response
from flask_cors import CORS
import subprocess

app = Flask(__name__)

# Allow CORS for Vercel + local
CORS(app, resources={r"/*": {"origins": [
    "http://localhost:3000",
    "https://ppt-crafter.vercel.app"
]}}, supports_credentials=True)

DEFAULT_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "api", "default_template.pptx")

# --- Health check ---
@app.get("/")
def health_root():
    return {"status": "ok", "message": "PPT Crafter API is running"}

# --- POST endpoint ---
@app.route("/api", methods=["POST", "OPTIONS"])
def generate():
    try:
        # Handle CORS preflight
        if request.method == "OPTIONS":
            response = make_response()
            response.headers["Access-Control-Allow-Origin"] = request.headers.get("Origin", "*")
            response.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
            response.headers["Access-Control-Allow-Headers"] = "Content-Type"
            return response, 200

        if "excel" not in request.files:
            return ("Missing file: need 'excel'", 400)

        excel = request.files["excel"]
        ppt   = request.files.get("template")

        if not excel.filename.lower().endswith((".xlsx", ".xls")):
            return ("Excel must be .xlsx or .xls", 400)
        if ppt and ppt.filename and not ppt.filename.lower().endswith(".pptx"):
            return ("Template must be .pptx", 400)

        if (not ppt or not ppt.filename) and not os.path.exists(DEFAULT_TEMPLATE_PATH):
            return ("Server template missing. Please add api/default_template.pptx to the repo.", 500)

        work = os.path.join(tempfile.gettempdir(), f"imarc_{uuid.uuid4().hex}")
        os.makedirs(work, exist_ok=True)
        print("=== DEBUG: Work dir created:", work)

        try:
            excel_path = os.path.join(work, "datasheet_imarc.xlsx")
            ppt_path   = os.path.join(work, "template.pptx")

            excel.save(excel_path)
            print("=== DEBUG: Excel saved at", excel_path)

            if ppt and ppt.filename:
                ppt.save(ppt_path)
                print("=== DEBUG: Custom template saved at", ppt_path)
            else:
                shutil.copyfile(DEFAULT_TEMPLATE_PATH, ppt_path)
                print("=== DEBUG: Default template copied at", ppt_path)

            script_src = os.path.join(os.path.dirname(__file__), "generate_poc.py")
            script_dst = os.path.join(work, "generate_poc.py")
            shutil.copyfile(script_src, script_dst)
            print("=== DEBUG: Script copied to work dir:", script_dst)

            try:
                proc = subprocess.run(
                    [
                        sys.executable,
                        script_dst,
                        os.path.basename(excel_path),
                        os.path.basename(ppt_path),
                        "updated_poc.pptx",
                    ],
                    cwd=work,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    timeout=240,  # must be less than gunicorn worker timeout
                )
            except subprocess.TimeoutExpired as te:
                # Child took too long — return a controlled error (with CORS header)
                msg = f"Script timeout (after {te.timeout}s)."
                print("=== DEBUG: Subprocess timeout ===", te)
                response = (f"{msg}\n\nSTDOUT:\n{te.stdout}\n\nSTDERR:\n{te.stderr}", 500)
                # ensure CORS header on error
                resp = make_response(response[0], 500)
                resp.headers["Access-Control-Allow-Origin"] = request.headers.get("Origin", "*")
                return resp
            except Exception as e:
                print("=== DEBUG: Subprocess run failed ===", e)
                resp = make_response(f"Failed to run script: {e}", 500)
                resp.headers["Access-Control-Allow-Origin"] = request.headers.get("Origin", "*")
                return resp

            # regular logging of child output
            print("=== DEBUG: Subprocess finished ===")
            print("Return Code:", proc.returncode)
            print("STDOUT:\n", proc.stdout)
            print("STDERR:\n", proc.stderr)

            if proc.returncode != 0:
                resp_text = f"Script failed\nSTDOUT:\n{proc.stdout}\n\nSTDERR:\n{proc.stderr}"
                resp = make_response(resp_text, 500)
                resp.headers["Access-Control-Allow-Origin"] = request.headers.get("Origin", "*")
                return resp

            out_path = os.path.join(work, "updated_poc.pptx")
            if not os.path.exists(out_path):
                print("=== DEBUG: Output file not found ===")
                return ("Output PPTX not found (expected 'updated_poc.pptx')", 500)

            with open(out_path, "rb") as f:
                data = f.read()

            response = send_file(
                io.BytesIO(data),
                mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                as_attachment=True,
                download_name="updated_poc.pptx",
            )
            response.headers["Access-Control-Allow-Origin"] = request.headers.get("Origin", "*")
            return response

        finally:
            try:
                shutil.rmtree(work)
                print("=== DEBUG: Work dir cleaned ===")
            except Exception as e:
                print("=== DEBUG: Failed to clean work dir:", e)

    except Exception as e:
        print("=== EXCEPTION in /api ===")
        print(traceback.format_exc())
        return (f"Exception: {str(e)}", 500)
