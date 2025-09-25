import io
import os
import sys
import tempfile
import shutil
import uuid
import traceback

from flask import Flask, request, send_file, make_response
from flask_cors import CORS

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": [
    "http://localhost:3000",
    "https://ppt-crafter.vercel.app"
]}}, supports_credentials=True)

# Path to default template in repo (adjust if your layout differs)
DEFAULT_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "api", "default_template.pptx")

# --- Health check ---
@app.get("/")
def health_root():
    return {"status": "ok", "message": "PPT Crafter API is running"}

# --- POST endpoint ---
@app.route("/api", methods=["POST", "OPTIONS"])
def generate():
    try:
        # CORS preflight
        if request.method == "OPTIONS":
            resp = make_response()
            resp.headers["Access-Control-Allow-Origin"] = request.headers.get("Origin", "*")
            resp.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
            resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
            return resp, 200

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
            out_path   = os.path.join(work, "updated_poc.pptx")

            excel.save(excel_path)
            print("=== DEBUG: Excel saved at", excel_path)

            if ppt and ppt.filename:
                ppt.save(ppt_path)
                print("=== DEBUG: Custom template saved at", ppt_path)
            else:
                shutil.copyfile(DEFAULT_TEMPLATE_PATH, ppt_path)
                print("=== DEBUG: Default template copied at", ppt_path)

            # --- Direct call: import and run the generator function ---
            try:
                # Import here so top-level imports in generate_poc don't run before temp files are ready
                from generate_poc import main as generate_main
            except Exception as e:
                print("=== ERROR importing generate_poc ===", e)
                resp = make_response(f"Failed to import generator: {e}", 500)
                resp.headers["Access-Control-Allow-Origin"] = request.headers.get("Origin", "*")
                return resp

            try:
                # Call generator with full absolute paths
                generate_main(excel_path, ppt_path, out_path)
            except Exception as e:
                tb = traceback.format_exc()
                print("=== ERROR running generate_main ===")
                print(tb)
                resp = make_response(f"Generator raised an exception:\n\n{tb}", 500)
                resp.headers["Access-Control-Allow-Origin"] = request.headers.get("Origin", "*")
                return resp

            if not os.path.exists(out_path):
                print("=== DEBUG: Output file not found after generator run ===")
                resp = make_response("Output PPTX not found (expected 'updated_poc.pptx')", 500)
                resp.headers["Access-Control-Allow-Origin"] = request.headers.get("Origin", "*")
                return resp

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