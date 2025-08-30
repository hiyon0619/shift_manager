import os

def create_html(original_file, generated_files):
    save_dir = os.path.join(os.path.dirname(__file__), "templates")
    os.makedirs(save_dir, exist_ok=True)
    html_path = os.path.join(save_dir, "index.html")

    html_content = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>シフト管理システム</title>
<style>
  body { font-family: Arial; margin:0; padding:0; background:#333; display:flex; justify-content:center; align-items:center; height:100vh; }
  .file-box-container { display:flex; flex-direction:column; align-items:center; gap:15px; padding:20px; background:#fff; border-radius:8px; width:650px; }
  header { text-align:center; font-size:36px; font-weight:bold; color:white; background:#007bff; width:100%; padding:20px 0; border-radius:6px; margin-bottom:20px; }
  .file-box { background:#f0f0f0; color:#333; padding:10px 15px; border-radius:6px; text-align:center; font-size:14px; width:200px; box-shadow:1px 1px 5px rgba(0,0,0,0.2); cursor:pointer; transition: transform 0.2s; }
  .file-box:hover { transform: translateY(-3px); }
  .action-container { display:flex; gap:20px; justify-content:center; margin-top:30px; width:100%; }
  .action-box { flex:1; max-width:600px; text-align:center; padding:15px; font-size:16px; font-weight:bold; cursor:pointer; border-radius:6px; background:#007bff; color:white; transition: transform 0.2s; }
  .action-box:hover { transform: translateY(-3px); }
</style>
</head>
<body>
<div class="file-box-container">
<header>シフト管理システム</header>
"""

    # 元ファイル + 生成ファイル
    all_files = [original_file] + generated_files
    while len(all_files) < 5:
        all_files.append("-")  # 空ボックス補充

    for f in all_files[:5]:
        html_content += f'<div class="file-box">{f}</div>\n'

    html_content += """
<div class="action-container">
  <div class="action-box" onclick="alert('手順1実行')">手順1実行</div>
  <div class="action-box" onclick="alert('手順2実行')">手順2実行</div>
</div>
</div>
</body>
</html>
"""

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    return html_path
