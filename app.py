# -*- coding: utf-8 -*-
import os
import io
import tempfile
import zipfile
from typing import Optional, Tuple

from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

# 保持 pdf.py 原样不改，这里仅导入
from pdf import PDFtoDocxConverter


def parse_page_range(page_range: str) -> Tuple[int, Optional[int]]:
    """将类似 "1-5" 或 "3" 的页码范围解析为 (start_page_0_based, end_page_exclusive)
    与 pdf.py 的交互式逻辑一致：
    - 空字符串: (0, None)
    - "a-b": (a-1, b)
    - "a": (a-1, a)
    """
    start_page, end_page = 0, None
    if not page_range:
        return start_page, end_page
    try:
        page_range = page_range.strip()
        if "-" in page_range:
            a, b = page_range.split("-", 1)
            start_page = max(int(a) - 1, 0)
            end_page = int(b)
        else:
            start_page = max(int(page_range) - 1, 0)
            end_page = start_page + 1
    except Exception:
        # 解析失败则转换全部
        start_page, end_page = 0, None
    return start_page, end_page


def create_app() -> Flask:
    app = Flask(__name__, static_folder='static', template_folder='templates')

    # 初始化转换器
    converter = PDFtoDocxConverter()

    @app.get('/')
    def index():
        return render_template('index.html')

    @app.post('/api/pdf-info')
    def pdf_info():
        file = request.files.get('pdf')
        if not file:
            return jsonify({"ok": False, "msg": "请上传PDF文件(pdf)"}), 400
        # 将上传文件写入临时文件
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            file.save(tmp.name)
            tmp_pdf_path = tmp.name
        try:
            info = converter.get_pdf_info(tmp_pdf_path)
            if not info:
                return jsonify({"ok": False, "msg": "获取PDF信息失败"}), 500
            return jsonify({"ok": True, "data": info})
        finally:
            try:
                os.remove(tmp_pdf_path)
            except Exception:
                pass

    @app.post('/api/convert-single')
    def convert_single():
        file = request.files.get('pdf')
        page_range = request.form.get('range', '')
        if not file:
            return jsonify({"ok": False, "msg": "请上传PDF文件(pdf)"}), 400

        start_page, end_page = parse_page_range(page_range)

        # 保存PDF到临时文件
        with tempfile.TemporaryDirectory() as td:
            pdf_path = os.path.join(td, secure_filename(file.filename) or 'input.pdf')
            file.save(pdf_path)

            # 输出DOCX路径
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            docx_path = os.path.join(td, f"{base_name}.docx")

            ok = converter.convert_single_file(pdf_path, docx_path, start_page, end_page)
            if not ok or not os.path.exists(docx_path):
                return jsonify({"ok": False, "msg": "转换失败"}), 500

            # 将DOCX作为下载返回
            return send_file(
                docx_path,
                as_attachment=True,
                download_name=f"{base_name}.docx",
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )

    @app.post('/api/convert-batch')
    def convert_batch():
        files = request.files.getlist('pdfs')
        page_range = request.form.get('range', '')
        if not files:
            return jsonify({"ok": False, "msg": "请上传至少一个PDF文件(pdfs)"}), 400

        start_page, end_page = parse_page_range(page_range)

        with tempfile.TemporaryDirectory() as td:
            out_dir = os.path.join(td, 'out')
            os.makedirs(out_dir, exist_ok=True)

            success = 0
            for f in files:
                if not f or not f.filename.lower().endswith('.pdf'):
                    continue
                safe_name = secure_filename(f.filename)
                pdf_path = os.path.join(td, safe_name or 'input.pdf')
                f.save(pdf_path)
                base = os.path.splitext(os.path.basename(pdf_path))[0]
                docx_path = os.path.join(out_dir, f"{base}.docx")
                if converter.convert_single_file(pdf_path, docx_path, start_page, end_page):
                    success += 1

            if success == 0:
                return jsonify({"ok": False, "msg": "没有文件成功转换"}), 500

            # 打包为zip返回
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for name in os.listdir(out_dir):
                    full_path = os.path.join(out_dir, name)
                    if os.path.isfile(full_path):
                        zf.write(full_path, arcname=name)
            zip_buffer.seek(0)

            return send_file(
                zip_buffer,
                as_attachment=True,
                download_name='converted_docx.zip',
                mimetype='application/zip'
            )

    return app


if __name__ == '__main__':
    # 直接运行：python app.py
    app = create_app()
    app.run(host='0.0.0.0', port=5000, debug=True)
