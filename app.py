from flask import Flask, render_template, request, send_file
from pathlib import Path
import process_excel
import tempfile

app = Flask(__name__)

# 一時フォルダを使用
UPLOAD_FOLDER = Path(tempfile.gettempdir())
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/execute/<step>", methods=["POST"])
def execute(step):
    """
    step: "1" または "2"
    POSTデータでファイルを受け取り実行し、生成ファイルを返す
    """
    files = []
    for i in range(5):  # input0～input4
        file = request.files.get(f"input{i}")
        if file:
            temp_path = UPLOAD_FOLDER / file.filename
            file.save(temp_path)
            files.append(temp_path)

    if step == "1":
        if not files:
            return "手順1: ファイルが選択されていません", 400
        output_file = process_excel.run_step1(files[0])
    elif step == "2":
        if len(files) < 2:
            return "手順2: 下4つのファイルが選択されていません", 400
        selected_files = files[1:5]  # 下4つ
        output_file = process_excel.run_step2(selected_files)
    else:
        return "無効なステップ", 400

    return send_file(output_file, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
