import io
import shutil
import tempfile
import uuid
import zipfile
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

from generator import find_placeholders, generate_all, read_excel

BASE = Path(__file__).parent
DEFAULT_OUT = BASE / "output"
WORK = Path(tempfile.mkdtemp(prefix="hrdocs_"))

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
sessions = {}  # token -> {"xlsx":..., "docx":..., "out":...}


@app.get("/")
def index():
    return render_template("index.html", default_out=str(DEFAULT_OUT))


@app.post("/upload")
def upload():
    x, d = request.files.get("xlsx"), request.files.get("docx")
    if not x or not d:
        return jsonify(error="Загрузите оба файла: Excel и Word"), 400
    if not x.filename.lower().endswith(".xlsx") or not d.filename.lower().endswith(".docx"):
        return jsonify(error="Нужны файлы .xlsx и .docx"), 400
    token = uuid.uuid4().hex
    folder = WORK / token
    folder.mkdir()
    xp, dp = folder / "data.xlsx", folder / "template.docx"
    x.save(xp)
    d.save(dp)
    try:
        header, records = read_excel(xp)
        placeholders = find_placeholders(dp)
    except Exception as e:
        shutil.rmtree(folder, ignore_errors=True)
        return jsonify(error=f"Не удалось прочитать файлы: {e}"), 400
    sessions[token] = {"xlsx": xp, "docx": dp, "out": None}
    return jsonify(
        token=token,
        columns=header,
        count=len(records),
        preview=records[:5],
        placeholders=placeholders,
        unmatched=[p for p in placeholders if p not in header],
    )


@app.post("/generate")
def generate():
    body = request.get_json(force=True)
    s = sessions.get(body.get("token"))
    if not s:
        return jsonify(error="Сессия устарела, загрузите файлы заново"), 400
    out = Path(body.get("out_dir") or DEFAULT_OUT).expanduser()
    try:
        files, warnings = generate_all(s["xlsx"], s["docx"], out, body.get("name_column"))
    except Exception as e:
        return jsonify(error=str(e)), 400
    s["out"], s["files"] = out, files
    return jsonify(files=files, warnings=warnings, out_dir=str(out))


@app.get("/zip/<token>")
def zip_result(token):
    s = sessions.get(token)
    if not s or not s.get("files"):
        return "Нет результатов", 404
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for f in s["files"]:
            z.write(s["out"] / f, f)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="documents.zip")


if __name__ == "__main__":
    print("Откройте http://127.0.0.1:5050")
    app.run(host="127.0.0.1", port=5050)
