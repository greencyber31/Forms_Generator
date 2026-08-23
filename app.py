import os
import sys
import json
import queue
import threading
import subprocess
from datetime import datetime
from pathlib import Path
from flask import Flask, render_template, request, jsonify, Response, send_file, send_from_directory
from werkzeug.utils import secure_filename
from pypdf import PdfReader

from template_builder import extract_excel_headers, extract_docx_tags, auto_match_tags, render_sample_pdf_preview, render_docx_to_pdf_preview
from generator_engine import run_batch_generation

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.abspath("uploads")
app.config['OUTPUT_FOLDER'] = os.path.abspath("output")
app.config['TEMP_FOLDER']   = os.path.abspath("temp_docs")
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

Path(app.config['UPLOAD_FOLDER']).mkdir(parents=True, exist_ok=True)
Path(app.config['OUTPUT_FOLDER']).mkdir(parents=True, exist_ok=True)
Path(app.config['TEMP_FOLDER']).mkdir(parents=True, exist_ok=True)

workspace_state = {
    "excel_file": None,
    "template_file": None,
    "transmittal_template": None,
    "headers": [],
    "app_docx_tags": [],
    "trans_docx_tags": [],
    "mappings": {
        "template": {},
        "transmittal": {}
    },
    "transmittal_columns": []
}

def build_default_transmittal_columns(headers):
    raw_headers = [h["original"] if isinstance(h, dict) else str(h) for h in headers]
    defaults = [
        ['last name', 'surname', 'lname'],
        ['first name', 'fname'],
        ['middle name', 'middlename', 'mname'],
        ['birth date', 'birthdate', 'birthday', 'dob', 'bday'],
        ['gender', 'sex'],
        ['reference no', 'reference no.', 'cfitf no', 'farmer id', 'ref no']
    ]
    col_items = []
    used_indices = set()
    order_counter = 1

    for def_candidates in defaults:
        for idx, h in enumerate(raw_headers):
            if idx in used_indices:
                continue
            h_clean = h.strip().lower().replace('.', '').replace('_', ' ')
            if any(cand in h_clean for cand in def_candidates):
                col_items.append({
                    "header": h,
                    "enabled": True,
                    "order": order_counter
                })
                used_indices.add(idx)
                order_counter += 1
                break

    for idx, h in enumerate(raw_headers):
        if idx not in used_indices:
            col_items.append({
                "header": h,
                "enabled": False,
                "order": order_counter
            })
            order_counter += 1

    return col_items

def init_workspace():
    excel_path = "CFITF Farmers.xlsx"
    app_template = "Application for Crop Insurance 2026 with tags.docx"
    if not os.path.exists(app_template):
        app_template = "Application for Crop Insurance 2026 final only.docx"
    trans_template = "Transmittal_Template.docx"

    if os.path.exists(excel_path):
        workspace_state["excel_file"] = os.path.abspath(excel_path)
        try:
            workspace_state["headers"] = extract_excel_headers(excel_path)
            workspace_state["transmittal_columns"] = build_default_transmittal_columns(workspace_state["headers"])
        except Exception:
            pass

    excel_raw_headers = [h["original"] for h in workspace_state["headers"]]

    if os.path.exists(app_template):
        workspace_state["template_file"] = os.path.abspath(app_template)
        try:
            workspace_state["app_docx_tags"] = extract_docx_tags(app_template)
            workspace_state["mappings"]["template"] = auto_match_tags(excel_raw_headers, workspace_state["app_docx_tags"])
        except Exception:
            pass

    if os.path.exists(trans_template):
        workspace_state["transmittal_template"] = os.path.abspath(trans_template)
        try:
            workspace_state["trans_docx_tags"] = extract_docx_tags(trans_template)
            workspace_state["mappings"]["transmittal"] = auto_match_tags(excel_raw_headers, workspace_state["trans_docx_tags"])
        except Exception:
            pass

init_workspace()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/workspace', methods=['GET'])
def get_workspace():
    return jsonify({
        "excel_file": os.path.basename(workspace_state["excel_file"]) if workspace_state["excel_file"] else None,
        "template_file": os.path.basename(workspace_state["template_file"]) if workspace_state["template_file"] else None,
        "transmittal_template": os.path.basename(workspace_state["transmittal_template"]) if workspace_state["transmittal_template"] else None,
        "headers": workspace_state["headers"],
        "app_docx_tags": workspace_state["app_docx_tags"],
        "trans_docx_tags": workspace_state["trans_docx_tags"],
        "mappings": workspace_state["mappings"],
        "transmittal_columns": workspace_state["transmittal_columns"]
    })

@app.route('/api/upload', methods=['POST'])
def upload_file():
    file_type = request.form.get('file_type')
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "Empty filename"}), 400

    filename = secure_filename(file.filename)
    save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(save_path)

    excel_raw_headers = [h["original"] for h in workspace_state["headers"]]

    if file_type == 'excel':
        workspace_state["excel_file"] = save_path
        workspace_state["headers"] = extract_excel_headers(save_path)
        workspace_state["transmittal_columns"] = build_default_transmittal_columns(workspace_state["headers"])
        excel_raw_headers = [h["original"] for h in workspace_state["headers"]]
        # Re-run auto-match with new excel headers
        if workspace_state["template_file"]:
            workspace_state["mappings"]["template"] = auto_match_tags(excel_raw_headers, workspace_state["app_docx_tags"])
        if workspace_state["transmittal_template"]:
            workspace_state["mappings"]["transmittal"] = auto_match_tags(excel_raw_headers, workspace_state["trans_docx_tags"])

    elif file_type == 'template':
        workspace_state["template_file"] = save_path
        workspace_state["app_docx_tags"] = extract_docx_tags(save_path)
        workspace_state["mappings"]["template"] = auto_match_tags(excel_raw_headers, workspace_state["app_docx_tags"])

    elif file_type == 'transmittal':
        workspace_state["transmittal_template"] = save_path
        workspace_state["trans_docx_tags"] = extract_docx_tags(save_path)
        workspace_state["mappings"]["transmittal"] = auto_match_tags(excel_raw_headers, workspace_state["trans_docx_tags"])

    trigger_background_preview_warmup()

    return jsonify({
        "message": f"Successfully uploaded {filename}",
        "file_type": file_type,
        "workspace": {
            "excel_file": os.path.basename(workspace_state["excel_file"]) if workspace_state["excel_file"] else None,
            "template_file": os.path.basename(workspace_state["template_file"]) if workspace_state["template_file"] else None,
            "transmittal_template": os.path.basename(workspace_state["transmittal_template"]) if workspace_state["transmittal_template"] else None,
            "headers": workspace_state["headers"],
            "app_docx_tags": workspace_state["app_docx_tags"],
            "trans_docx_tags": workspace_state["trans_docx_tags"],
            "mappings": workspace_state["mappings"],
            "transmittal_columns": workspace_state["transmittal_columns"]
        }
    })

@app.route('/api/template/mapping', methods=['GET', 'POST'])
def handle_mapping():
    if request.method == 'GET':
        doc_type = request.args.get('type', 'template')
        docx_tags = workspace_state["app_docx_tags"] if doc_type == 'template' else workspace_state["trans_docx_tags"]
        excel_headers = [h["original"] for h in workspace_state["headers"]]
        current_mapping = workspace_state["mappings"].get(doc_type, {})
        auto_mappings = auto_match_tags(excel_headers, docx_tags)

        return jsonify({
            "doc_type": doc_type,
            "docx_tags": docx_tags,
            "excel_headers": excel_headers,
            "mapping": current_mapping,
            "auto_mappings": auto_mappings
        })

    elif request.method == 'POST':
        data = request.json
        doc_type = data.get('doc_type', 'template')
        mapping = data.get('mapping', {})

        workspace_state["mappings"][doc_type] = mapping
        trigger_background_preview_warmup()
        return jsonify({
            "message": f"Saved field mappings for {doc_type}!",
            "mapping": mapping
        })

@app.route('/api/template/auto-match', methods=['POST'])
def auto_match_route():
    data = request.json
    doc_type = data.get('doc_type', 'template')
    docx_tags = workspace_state["app_docx_tags"] if doc_type == 'template' else workspace_state["trans_docx_tags"]
    excel_headers = [h["original"] for h in workspace_state["headers"]]

    auto_map = auto_match_tags(excel_headers, docx_tags)
    workspace_state["mappings"][doc_type] = auto_map
    trigger_background_preview_warmup()

    return jsonify({
        "message": f"Auto-matched {len(auto_map)} field tags!",
        "mapping": auto_map
    })

@app.route('/api/template/transmittal-columns', methods=['GET', 'POST'])
def handle_transmittal_columns():
    if request.method == 'GET':
        if not workspace_state["transmittal_columns"] and workspace_state["headers"]:
            workspace_state["transmittal_columns"] = build_default_transmittal_columns(workspace_state["headers"])
        return jsonify({
            "columns": workspace_state["transmittal_columns"]
        })
    elif request.method == 'POST':
        data = request.json
        cols = data.get('columns', [])
        workspace_state["transmittal_columns"] = cols
        trigger_background_preview_warmup()
        return jsonify({
            "message": "Saved transmittal columns configuration!",
            "columns": cols
        })

def _background_render_preview(doc_type: str):
    try:
        target_path = workspace_state["template_file"] if doc_type == 'template' else workspace_state["transmittal_template"]
        excel_path = workspace_state["excel_file"]
        mapping = workspace_state["mappings"].get(doc_type, {})
        trans_cols = workspace_state.get("transmittal_columns", [])

        if target_path and os.path.exists(target_path) and excel_path and os.path.exists(excel_path):
            render_sample_pdf_preview(
                excel_path,
                target_path,
                mapping,
                app.config['TEMP_FOLDER'],
                f"preview_{doc_type}",
                transmittal_columns=trans_cols
            )
    except Exception:
        pass

def trigger_background_preview_warmup():
    """Spawns non-blocking background threads to pre-render preview PDFs in advance."""
    t1 = threading.Thread(target=_background_render_preview, args=('template',), daemon=True)
    t2 = threading.Thread(target=_background_render_preview, args=('transmittal',), daemon=True)
    t1.start()
    t2.start()

@app.route('/api/template/pdf-preview', methods=['GET'])
def get_pdf_preview():
    doc_type = request.args.get('type', 'template')
    target_path = workspace_state["template_file"] if doc_type == 'template' else workspace_state["transmittal_template"]
    excel_path = workspace_state["excel_file"]
    mapping = workspace_state["mappings"].get(doc_type, {})
    trans_cols = workspace_state.get("transmittal_columns", [])

    if not target_path or not os.path.exists(target_path):
        return jsonify({"error": "Template file not loaded"}), 400

    try:
        pdf_name = f"preview_{doc_type}"
        pdf_path = render_sample_pdf_preview(
            excel_path,
            target_path,
            mapping,
            app.config['TEMP_FOLDER'],
            pdf_name,
            transmittal_columns=trans_cols
        )
        return send_file(pdf_path, mimetype='application/pdf')
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/process/stream')
def stream_process():
    primary_group = request.args.get('primary_group')
    secondary_group = request.args.get('secondary_group')
    bundle_group = request.args.get('bundle_group')
    test_limit = request.args.get('test_limit', type=int)
    max_workers = request.args.get('max_workers', default=4, type=int)

    excel_path = workspace_state["excel_file"]
    template_path = workspace_state["template_file"]
    trans_path = workspace_state["transmittal_template"]
    app_mapping = workspace_state["mappings"].get("template", {})
    trans_mapping = workspace_state["mappings"].get("transmittal", {})
    trans_cols = workspace_state.get("transmittal_columns", [])

    if not excel_path or not template_path or not trans_path:
        def err_generator():
            yield f"data: {json.dumps({'type': 'log', 'status': 'error', 'message': 'Missing required uploaded files!'})}\n\n"
        return Response(err_generator(), mimetype='text/event-stream')

    msg_queue = queue.Queue()

    def progress_callback(data):
        msg_queue.put(data)

    def worker():
        try:
            summary = run_batch_generation(
                excel_file=excel_path,
                template_file=template_path,
                transmittal_template=trans_path,
                output_dir=app.config['OUTPUT_FOLDER'],
                temp_dir=app.config['TEMP_FOLDER'],
                primary_group_col=primary_group,
                secondary_group_col=secondary_group,
                bundle_group_col=bundle_group,
                test_limit=test_limit,
                max_workers=max_workers,
                template_mapping=app_mapping,
                transmittal_mapping=trans_mapping,
                transmittal_columns=trans_cols,
                progress_callback=progress_callback
            )
            msg_queue.put({"type": "complete", "summary": summary})
        except Exception as exc:
            msg_queue.put({"type": "error", "message": str(exc)})
        finally:
            msg_queue.put(None)

    threading.Thread(target=worker, daemon=True).start()

    def generate_events():
        while True:
            item = msg_queue.get()
            if item is None:
                break
            yield f"data: {json.dumps(item)}\n\n"

    return Response(generate_events(), mimetype='text/event-stream')

def format_size(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.2f} MB"

def scan_output_directory(base_dir: str) -> dict:
    base_path = Path(base_dir).resolve()
    if not base_path.exists():
        return {
            "total_files": 0,
            "total_folders": 0,
            "total_size_bytes": 0,
            "total_size_formatted": "0 B",
            "folders": []
        }

    folders_dict = {}
    root_files = []
    total_files = 0
    total_size = 0

    for root, dirs, files in os.walk(base_path):
        current_path = Path(root).resolve()
        pdf_files = [f for f in files if f.lower().endswith('.pdf')]
        if not pdf_files:
            continue

        rel_dir = os.path.relpath(current_path, base_path)
        is_root = (rel_dir == '.')

        file_list = []
        folder_size = 0

        for f_name in sorted(pdf_files):
            f_path = current_path / f_name
            try:
                stat = f_path.stat()
                f_size = stat.st_size
                mtime = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                f_size = 0
                mtime = ""

            total_files += 1
            total_size += f_size
            folder_size += f_size

            rel_file_path = (Path(rel_dir) / f_name).as_posix() if not is_root else f_name

            file_list.append({
                "name": f_name,
                "rel_path": rel_file_path,
                "size_bytes": f_size,
                "size_formatted": format_size(f_size),
                "modified": mtime
            })

        if is_root:
            root_files.extend(file_list)
        else:
            folder_name = os.path.basename(rel_dir)
            folders_dict[rel_dir] = {
                "name": folder_name,
                "rel_path": Path(rel_dir).as_posix(),
                "file_count": len(file_list),
                "folder_size_bytes": folder_size,
                "folder_size_formatted": format_size(folder_size),
                "files": file_list
            }

    folders_list = []
    if root_files:
        folders_list.append({
            "name": "Root Output",
            "rel_path": ".",
            "file_count": len(root_files),
            "folder_size_bytes": sum(f["size_bytes"] for f in root_files),
            "folder_size_formatted": format_size(sum(f["size_bytes"] for f in root_files)),
            "files": root_files
        })

    for folder_info in sorted(folders_dict.values(), key=lambda x: x["name"].lower()):
        folders_list.append(folder_info)

    return {
        "total_files": total_files,
        "total_folders": len(folders_list),
        "total_size_bytes": total_size,
        "total_size_formatted": format_size(total_size),
        "folders": folders_list
    }

@app.route('/api/output/tree', methods=['GET'])
def get_output_tree():
    out_dir = app.config['OUTPUT_FOLDER']
    tree = scan_output_directory(out_dir)
    return jsonify(tree)

@app.route('/api/output/view-file', methods=['GET'])
def view_output_file():
    rel_path = request.args.get('path', '')
    if not rel_path:
        return jsonify({"error": "No file path provided"}), 400

    out_base = Path(app.config['OUTPUT_FOLDER']).resolve()
    target_path = (out_base / rel_path).resolve()

    # Directory traversal prevention check
    if not str(target_path).startswith(str(out_base)):
        return jsonify({"error": "Access denied"}), 403

    if not target_path.exists() or not target_path.is_file():
        return jsonify({"error": "File not found"}), 404

    return send_file(str(target_path), mimetype='application/pdf')

@app.route('/api/output/download-file', methods=['GET'])
def download_output_file():
    rel_path = request.args.get('path', '')
    if not rel_path:
        return jsonify({"error": "No file path provided"}), 400

    out_base = Path(app.config['OUTPUT_FOLDER']).resolve()
    target_path = (out_base / rel_path).resolve()

    if not str(target_path).startswith(str(out_base)):
        return jsonify({"error": "Access denied"}), 403

    if not target_path.exists() or not target_path.is_file():
        return jsonify({"error": "File not found"}), 404

    return send_file(str(target_path), as_attachment=True, download_name=target_path.name)

@app.route('/api/output/save-pdf', methods=['POST'])
def save_output_pdf():
    rel_path = request.form.get('path', '').strip()
    if not rel_path:
        return jsonify({"error": "No file path provided"}), 400

    if 'file' not in request.files:
        return jsonify({"error": "No PDF file payload provided"}), 400

    out_base = Path(app.config['OUTPUT_FOLDER']).resolve()
    target_path = (out_base / rel_path).resolve()

    if not str(target_path).startswith(str(out_base)):
        return jsonify({"error": "Access denied"}), 403

    if not target_path.parent.exists():
        target_path.parent.mkdir(parents=True, exist_ok=True)

    uploaded_file = request.files['file']
    uploaded_file.save(str(target_path))

    # Invalidate cache for this file
    if str(target_path) in pdf_search_cache:
        del pdf_search_cache[str(target_path)]

    return jsonify({
        "message": f"Successfully saved changes directly to {target_path.name}!",
        "file_name": target_path.name,
        "size_formatted": format_size(target_path.stat().st_size)
    })

@app.route('/api/output/open', methods=['POST'])
def open_output_dir():
    data = request.json or {}
    rel_folder = data.get('folder', '')
    out_dir = Path(app.config['OUTPUT_FOLDER']).resolve()

    if rel_folder and rel_folder != '.':
        target_dir = (out_dir / rel_folder).resolve()
        if str(target_dir).startswith(str(out_dir)) and target_dir.exists():
            out_dir = target_dir

    try:
        if os.name == 'nt':
            os.startfile(str(out_dir))
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', str(out_dir)])
        else:
            subprocess.Popen(['xdg-open', str(out_dir)])
        return jsonify({"message": f"Opened folder: {out_dir}"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# =========================================================================
# Deep PDF Content Search & Caching Engine
# =========================================================================
pdf_search_cache = {}

def get_pdf_page_texts(file_path):
    try:
        mtime = os.path.getmtime(file_path)
        cached = pdf_search_cache.get(file_path)
        if cached and cached.get('mtime') == mtime:
            return cached.get('pages', [])
        
        reader = PdfReader(file_path)
        pages_data = []
        for idx, page in enumerate(reader.pages):
            text = page.extract_text() or ''
            clean_text = ' '.join(text.split())
            pages_data.append((idx + 1, clean_text))
        
        pdf_search_cache[file_path] = {'mtime': mtime, 'pages': pages_data}
        return pages_data
    except Exception as e:
        print(f"Error reading PDF {file_path}: {e}")
        return []

@app.route('/api/output/search', methods=['GET'])
def search_output_content():
    query = request.args.get('q', '').strip()
    if not query:
        return jsonify({"query": "", "results": [], "total_matching_files": 0, "total_matches": 0})
    
    q_lower = query.lower()
    out_base = Path(app.config['OUTPUT_FOLDER']).resolve()
    if not out_base.exists():
        return jsonify({"query": query, "results": [], "total_matching_files": 0, "total_matches": 0})
    
    results = []
    total_matches = 0
    
    for root, dirs, files in os.walk(out_base):
        for file_name in files:
            if not file_name.lower().endswith('.pdf'):
                continue
            
            full_path = os.path.join(root, file_name)
            rel_file = os.path.relpath(full_path, out_base).replace('\\', '/')
            rel_folder = os.path.dirname(rel_file) or '.'
            folder_display = os.path.basename(root) if rel_folder != '.' else 'Root Directory'
            
            name_matches = (q_lower in file_name.lower()) or (q_lower in folder_display.lower())
            pages_data = get_pdf_page_texts(full_path)
            matched_snippets = []
            
            for page_num, page_text in pages_data:
                page_text_lower = page_text.lower()
                pos = 0
                while True:
                    idx = page_text_lower.find(q_lower, pos)
                    if idx == -1:
                        break
                    
                    start = max(0, idx - 40)
                    end = min(len(page_text), idx + len(query) + 40)
                    snippet = ("..." if start > 0 else "") + page_text[start:end].strip() + ("..." if end < len(page_text) else "")
                    
                    matched_snippets.append({
                        "page": page_num,
                        "snippet": snippet
                    })
                    
                    pos = idx + len(q_lower)
                    if len(matched_snippets) >= 8:
                        break
                if len(matched_snippets) >= 8:
                    break
            
            if matched_snippets or name_matches:
                size_bytes = os.path.getsize(full_path)
                match_count = len(matched_snippets)
                total_matches += match_count
                results.append({
                    "file_name": file_name,
                    "rel_path": rel_file,
                    "folder_name": folder_display,
                    "folder_rel_path": rel_folder,
                    "size_formatted": format_size(size_bytes),
                    "page_count": len(pages_data),
                    "match_count": match_count,
                    "name_match": name_matches,
                    "snippets": matched_snippets
                })
    
    results.sort(key=lambda x: (x['match_count'] > 0, x['match_count']), reverse=True)
    
    return jsonify({
        "query": query,
        "total_matching_files": len(results),
        "total_matches": total_matches,
        "results": results
    })

@app.route('/api/output/doc-info', methods=['GET'])
def get_output_doc_info():
    rel_path = request.args.get('path', '').strip()
    if not rel_path:
        return jsonify({"error": "No path provided"}), 400
    
    out_base = Path(app.config['OUTPUT_FOLDER']).resolve()
    target_path = (out_base / rel_path).resolve()
    
    if not str(target_path).startswith(str(out_base)):
        return jsonify({"error": "Access denied"}), 403
    
    if not target_path.exists() or not target_path.is_file():
        return jsonify({"error": "File not found"}), 404
    
    pages_data = get_pdf_page_texts(str(target_path))
    size_bytes = os.path.getsize(str(target_path))
    
    return jsonify({
        "file_name": target_path.name,
        "rel_path": rel_path,
        "size_formatted": format_size(size_bytes),
        "page_count": len(pages_data),
        "pages": [{"page": p[0], "text": p[1]} for p in pages_data]
    })

if __name__ == '__main__':
    print("Starting PCIC Form Generator Web Server on http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)
