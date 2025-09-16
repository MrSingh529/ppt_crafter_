from flask import Flask, request, send_file
import os
from generate_poc import main  # assume your script has a `main()` function

app = Flask(__name__)

@app.route("/")
def home():
    return {"status": "ok", "message": "PPT Crafter API is running"}

@app.route("/generate", methods=["POST"])
def generate():
    excel = request.files.get("excel")
    ppt = request.files.get("ppt")

    if not excel:
        return {"error": "Excel file is required"}, 400

    excel.save("input.xlsx")
    template_file = "test_ppt.pptx"
    if ppt:
        ppt.save("template.pptx")
        template_file = "template.pptx"

    output_file = "result.pptx"
    main("input.xlsx", template_file, output_file)

    return send_file(output_file, as_attachment=True)