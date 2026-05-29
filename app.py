"""
抖店上架助手 - Web 应用后端
"""
import json
import csv
import base64
import os
import re
from pathlib import Path
from io import BytesIO

from flask import Flask, request, jsonify, render_template, send_file
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import requests

app = Flask(__name__)
CORS(app)

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "商品图" / "参考图"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# 通义千问配置
DASHSCOPE_API_KEY = "sk-3cc39ec1fde843f48e049c943419d0ef"
DASHSCOPE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
MODEL_NAME = "qwen-vl-max"

# 加载下拉选项配置
OPTIONS_CONFIG = {}
try:
    with open(BASE_DIR / "下拉选项配置.json", "r", encoding="utf-8") as f:
        OPTIONS_CONFIG = json.load(f)
except:
    pass

# 加载字段配置表，获取每个类目的字段列表
def load_field_config():
    """读取字段配置表，返回 {类目: [字段列表]}"""
    config = {}
    try:
        with open(BASE_DIR / "字段配置表.csv", "r", encoding="gbk") as f:
            for row in csv.reader(f):
                if len(row) < 7 or not row[0].strip():
                    continue
                cat = row[0].strip()
                field = row[1].strip()
                if cat not in config:
                    config[cat] = []
                config[cat].append(field)
    except:
        pass
    return config


# ==================== 路由 ====================

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/config", methods=["GET"])
def get_config():
    """返回配置信息"""
    field_config = load_field_config()
    return jsonify({
        "options": OPTIONS_CONFIG,
        "fields": field_config,
        "categories": [k for k in field_config.keys() if k != "所有类目"]
    })


@app.route("/api/upload-excel", methods=["POST"])
def upload_excel():
    """上传并解析 Excel 文件"""
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "未上传文件"}), 400

    # 保存文件
    path = BASE_DIR / "当前表格.xlsx"
    file.save(str(path))

    # 解析所有 Sheet
    wb = load_workbook(path, data_only=True)
    sheets_data = {}

    for sn in wb.sheetnames:
        if sn.startswith("_") or sn == "通用配置":
            continue
        ws = wb[sn]
        rows = []
        headers = []
        # 读取表头（第2行）
        for c in range(1, ws.max_column + 1):
            h = ws.cell(row=2, column=c).value
            if h and not ws.column_dimensions.get(get_column_letter(c), None):
                headers.append(str(h).replace("\n", " "))
            elif h:
                col_dim = ws.column_dimensions[get_column_letter(c)]
                if not col_dim.hidden:
                    headers.append(str(h).replace("\n", " "))
                else:
                    headers.append(None)  # 隐藏列标记
            else:
                headers.append(None)

        # 读取数据行（第5行起）
        for r in range(5, ws.max_row + 1):
            row_data = {}
            has_content = False
            for c in range(1, len(headers) + 1):
                h = headers[c - 1]
                if h is None:
                    continue
                val = ws.cell(row=r, column=c).value
                if val is not None:
                    val = str(val).strip()
                    has_content = True
                else:
                    val = ""
                row_data[h] = val
            if has_content:
                row_data["_row"] = r
                row_data["_category"] = sn
                rows.append(row_data)

        sheets_data[sn] = {
            "headers": [h for h in headers if h is not None],
            "rows": rows
        }

    wb.close()
    return jsonify({"sheets": sheets_data, "sheetNames": list(sheets_data.keys())})


@app.route("/api/upload-image", methods=["POST"])
def upload_image():
    """上传参考图"""
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "未上传文件"}), 400
    filename = file.filename
    file.save(str(UPLOAD_DIR / filename))
    return jsonify({"ok": True, "filename": filename})


@app.route("/api/ai-recognize", methods=["POST"])
def ai_recognize():
    """调用通义千问识图，识别服装属性"""
    data = request.json
    image_filename = data.get("image", "")
    category = data.get("category", "")
    fields_to_recognize = data.get("fields", [])

    image_path = UPLOAD_DIR / image_filename
    if not image_path.exists():
        return jsonify({"error": f"图片不存在: {image_filename}"}), 404

    # 读取图片并转 base64
    with open(image_path, "rb") as f:
        image_data = base64.b64encode(f.read()).decode()

    # 获取该类目的下拉选项
    options = OPTIONS_CONFIG.get(category, {})
    all_category_options = OPTIONS_CONFIG.get("所有类目", {})
    merged_options = {**all_category_options, **options}

    # 构建提示词
    fields_desc = []
    for fn in fields_to_recognize:
        opts = merged_options.get(fn, {}).get("options", "")
        if not opts:
            opts = merged_options.get(fn, {}).get("fixed", "")
        hint = f"- {fn}"
        if opts:
            opts_list = [o.strip() for o in opts.replace("/", ",").split(",") if o.strip()]
            if opts_list:
                hint += f"（从以下选一个最接近的：{'/'.join(opts_list[:30])}）"
        fields_desc.append(hint)

    prompt = f"""请识别这张服装图片的属性。这是一个{category}类目的商品。

请返回以下属性(JSON格式):
{chr(10).join(fields_desc)}

注意：
1. 只返回JSON，不要其他文字
2. 每个字段从给出的选项中选最接近的值
3. 如果无法判断，填空字符串
4. 颜色如果有多个，用/分隔，如"白色/黑色"
5. JSON的key用中文，与上面字段名完全一致

返回格式示例: {{"领型": "圆领", "袖长": "短袖", "颜色": "白色"}}"""

    try:
        resp = requests.post(
            DASHSCOPE_URL,
            headers={"Authorization": f"Bearer {DASHSCOPE_API_KEY}"},
            json={
                "model": MODEL_NAME,
                "messages": [{
                    "role": "user",
                    "content": [
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_data}"}},
                        {"type": "text", "text": prompt}
                    ]
                }],
                "max_tokens": 500
            },
            timeout=30
        )
        result = resp.json()
        content = result["choices"][0]["message"]["content"]

        # 尝试提取JSON
        json_match = re.search(r'\{[^}]+\}', content, re.DOTALL)
        if json_match:
            attrs = json.loads(json_match.group())
            return jsonify({"ok": True, "attributes": attrs, "raw": content})
        return jsonify({"ok": True, "attributes": {}, "raw": content})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/generate-title", methods=["POST"])
def generate_title():
    """生成30字商品标题和12字短标题"""
    data = request.json
    category = data.get("category", "")
    attributes = data.get("attributes", {})
    season = data.get("season", "夏季")

    # 构建属性描述
    attr_desc = " ".join([f"{k}:{v}" for k, v in attributes.items() if v])

    prompt = f"""你是一个抖音服装电商运营专家。请根据以下信息生成商品标题。

类目：{category}
属性：{attr_desc}
季节：{season}

要求：
1. 标题正好30个字，模仿抖音爆款风格
2. 标题必须包含类目关键词（如"连衣裙""T恤"等）
3. 如果是时尚套装，标题必须以"套装"结尾
4. 年份使用2026
5. 同时生成一个12个字的短标题（完整通顺的摘要，不能是截断的）
6. 只返回JSON：{{"title": "30字标题", "short_title": "12字短标题"}}
7. 不要任何其他文字"""

    try:
        resp = requests.post(
            DASHSCOPE_URL,
            headers={"Authorization": f"Bearer {DASHSCOPE_API_KEY}"},
            json={
                "model": "qwen-plus",
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": 200
            },
            timeout=20
        )
        result = resp.json()
        content = result["choices"][0]["message"]["content"]
        json_match = re.search(r'\{[^}]+\}', content, re.DOTALL)
        if json_match:
            title_data = json.loads(json_match.group())
            return jsonify({"ok": True, **title_data})
        return jsonify({"ok": False, "raw": content})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/save-excel", methods=["POST"])
def save_excel():
    """保存编辑后的数据到 Excel"""
    data = request.json
    sheets_data = data.get("sheets", {})

    path = BASE_DIR / "当前表格.xlsx"
    wb = load_workbook(path)

    for sn, sheet_info in sheets_data.items():
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        headers = sheet_info.get("headers", [])
        rows = sheet_info.get("rows", [])

        # 构建表头到列号的映射
        header_to_col = {}
        for c in range(1, ws.max_column + 1):
            h = ws.cell(row=2, column=c).value
            if h:
                header_to_col[str(h).replace("\n", " ").strip()] = c

        for row_data in rows:
            row_num = row_data.get("_row", 0)
            if not row_num:
                continue
            for h, val in row_data.items():
                if h.startswith("_"):
                    continue
                col_num = header_to_col.get(h)
                if col_num and val:
                    cell = ws.cell(row=row_num, column=col_num)
                    cell.value = val

    wb.save(str(path))
    wb.close()
    return jsonify({"ok": True})


@app.route("/api/export-excel", methods=["GET"])
def export_excel():
    """下载当前 Excel"""
    path = BASE_DIR / "当前表格.xlsx"
    if path.exists():
        return send_file(str(path), as_attachment=True, download_name="商品上架模板.xlsx")
    return jsonify({"error": "文件不存在"}), 404


@app.route("/api/start-auto-upload", methods=["POST"])
def start_auto_upload():
    """启动浏览器自动化上架（预留）"""
    return jsonify({"ok": False, "message": "浏览器自动化功能开发中，请先手动核对表格后上架"})


if __name__ == "__main__":
    print("抖店上架助手启动中...")
    print("打开浏览器访问: http://127.0.0.1:5000")
    app.run(debug=True, port=5000)
