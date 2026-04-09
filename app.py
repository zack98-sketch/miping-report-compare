# -*- coding: utf-8 -*-
"""
密评报告附录D对比工具 - Web版
上传两份docx报告 → 解析附录D表格 → HTML即时展示差异
用法: python app.py  → 打开浏览器 http://127.0.0.1:5678
"""

import os, sys, json, time
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file

# ── Paths ──────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)

sys.path.insert(0, BASE_DIR)
from core_engine import extract_docx_elements, find_appendix_d, build_d_hierarchy, compare_hierarchies

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB
app.config['JSON_AS_ASCII'] = False


# ==================== Pages ====================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/compare', methods=['POST'])
def compare():
    """Upload two docx files and return comparison as JSON"""
    try:
        # Save uploaded files
        file_a = request.files['file_a']
        file_b = request.files['file_b']
        if not file_a or not file_b:
            return jsonify({'ok': False, 'error': '请选择两份文件'})

        ts = int(time.time() * 1000)
        path_a = os.path.join(UPLOAD_DIR, f'a_{ts}_{safe_fn(file_a.filename)}')
        path_b = os.path.join(UPLOAD_DIR, f'b_{ts}_{safe_fn(file_b.filename)}')
        file_a.save(path_a)
        file_b.save(path_b)

        # Parse both documents
        t0 = time.time()
        print(f"[{datetime.now():%H:%M:%S}] Parsing document A: {file_a.filename}")
        elems_a = extract_docx_elements(path_a)
        
        print(f"[{datetime.now():%H:%M:%S}] Parsing document B: {file_b.filename}")
        elems_b = extract_docx_elements(path_b)

        print(f"[{datetime.now():%H:%M:%S}] Parsed: A={len(elems_a)} elements, B={len(elems_b)} elements")

        # Find Appendix D in each
        idx_a = find_appendix_d(elems_a)
        idx_b = find_appendix_d(elems_b)
        if idx_a < 0 or idx_b < 0:
            return jsonify({
                'ok': False,
                'error': f'未在文档中找到"附录D 单项测评结果记录"。请确认文档格式正确。',
                'details': {'found_a': idx_a >= 0, 'found_b': idx_b >= 0}
            })

        print(f"[{datetime.now():%H:%M:%S}] Appendix D found: A@{idx_a}, B@{idx_b}")

        # Build hierarchies
        hierarchy_a = build_d_hierarchy(elems_a, idx_a)
        hierarchy_b = build_d_hierarchy(elems_b, idx_b)
        elapsed = time.time() - t0
        
        # Compare
        comparison = compare_hierarchies(hierarchy_a, hierarchy_b)
        
        # Build response
        stats = {
            **comparison['stats'],
            'parse_time': round(elapsed, 1),
            'total_tables_a': sum(len(o.get('tables', [])) for s in hierarchy_a for o in s.get('objects', [])),
            'total_tables_b': sum(len(o.get('tables', [])) for s in hierarchy_b for o in s.get('objects', [])),
        }
        
        response = {
            'ok': True,
            'stats': stats,
            'sections': comparison['sections'],
            'file_a': {'name': file_a.filename, 'size_mb': round(os.path.getsize(path_a)/1048576, 1)},
            'file_b': {'name': file_b.filename, 'size_mb': round(os.path.getsize(path_b)/1048576, 1)},
            'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }

        # Clean up temp files
        try:
            os.remove(path_a); os.remove(path_b)
        except:
            pass

        return jsonify(response)

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'ok': False, 'error': str(e)})


@app.route('/result')
def result_page():
    """Show result page (for POST-redirect pattern or direct access with data)"""
    return render_template('index.html')


def safe_fn(name):
    """Sanitize filename"""
    import re
    name = os.path.basename(name)
    name = re.sub(r'[^\w\u4e00-\u9fff\-_.]', '_', name)[:80]
    return name or 'upload.docx'


# ==================== Main ====================

if __name__ == '__main__':
    port = 5678
    host = '127.0.0.1'
    
    # Force UTF-8 for Windows console
    import sys
    if sys.platform == 'win32':
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    
    print('=' * 60)
    print('  密评报告附录D 对比工具 v1.0')
    print('=' * 60)
    print(f'  启动中... 请打开浏览器访问:')
    print(f'  http://{host}:{port}')
    print('=' * 60)
    
    # Auto open browser
    import threading, webbrowser
    def _open():
        time.sleep(1.5)
        webbrowser.open(f'http://{host}:{port}')
    threading.Thread(target=_open, daemon=True).start()
    
    app.run(host=host, port=port, debug=False)
