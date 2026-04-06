from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from scheduler import generate_timetable
from utils import export_to_excel

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")

        if not file or file.filename == "":
            return "No file uploaded ❌"

        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        try:
            df = pd.read_excel(filepath)
        except Exception as e:
            return f"Error reading file: {str(e)}"

        try:
            timetable = generate_timetable(df)
            excel_path = os.path.join(OUTPUT_FOLDER, "timetable.xlsx")
            export_to_excel(timetable, excel_path)
        except Exception as e:
            return f"Processing Error: {str(e)}"

        return render_template("result.html", timetable=timetable)

    return render_template("index.html")

@app.route("/download/excel")
def download_excel():
    return send_file(os.path.join(OUTPUT_FOLDER, "timetable.xlsx"), as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
