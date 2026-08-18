import os
import json
import queue
import threading
import subprocess
from pathlib import Path
from flask import Flask, render_template, request, jsonify, Response, send_file, send_from_directory
from werkzeug.utils import secure_filename

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
    }
}

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
        "mappings": workspace_state["mappings"]
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
            "mappings": workspace_state["mappings"]
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

    return jsonify({
        "message": f"Auto-matched {len(auto_map)} field tags!",
        "mapping": auto_map
    })

@app.route('/api/template/pdf-preview', methods=['GET'])
def get_pdf_preview():
    doc_type = request.args.get('type', 'template')
    target_path = workspace_state["template_file"] if doc_type == 'template' else workspace_state["transmittal_template"]
    excel_path = workspace_state["excel_file"]
    mapping = workspace_state["mappings"].get(doc_type, {})

    if not target_path or not os.path.exists(target_path):
        return jsonify({"error": "Template file not loaded"}), 400

    try:
        pdf_name = f"preview_{doc_type}"
        pdf_path = render_sample_pdf_preview(excel_path, target_path, mapping, app.config['TEMP_FOLDER'], pdf_name)
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

@app.route('/api/output/open', methods=['POST'])
def open_output_dir():
    out_dir = app.config['OUTPUT_FOLDER']
    try:
        subprocess.Popen(['xdg-open', out_dir])
        return jsonify({"message": f"Opened output folder: {out_dir}"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    print("Starting PCIC Form Generator Web Server on http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)
