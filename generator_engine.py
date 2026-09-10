import pandas as pd
from docxtpl import DocxTemplate
from docx.shared import Mm
import os
import re
import subprocess
import concurrent.futures
import logging
import shutil
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from pypdf import PdfWriter
import traceback

_UNSAFE_CHARS = str.maketrans({
    '/': '_', '\\': '_', ':': '_', '*': '_',
    '?': '_', '"': '_', '<': '_', '>': '_', '|': '_',
})

def sanitize_key(k: str) -> str:
    """Convert an Excel column header into a valid Jinja2 variable name."""
    s = str(k).strip()
    s = s.replace(' ', '_').replace('.', '').replace('-', '_').replace('/', '_')
    return ''.join(c for c in s if c.isalnum() or c == '_')

def safe_filename(name: str) -> str:
    """Replace all OS-reserved / path-unsafe characters in a filename segment."""
    return str(name).translate(_UNSAFE_CHARS).strip()

def _get_underline_str(tag: str) -> str:
    tag_lower = tag.lower()
    if any(k in tag_lower for k in ['area', 'suffix', 'lot', 'sex', 'gender', 'civil']):
        return "__________"
    return "________________"

def _build_context(row_dict: dict, mapping: dict | None = None) -> dict:
    """Convert a raw row dictionary into a clean Jinja2 rendering context using mapped tags."""
    context = {}
    for k, v in row_dict.items():
        if pd.isna(v):
            clean_v = ""
        elif isinstance(v, str):
            clean_v = v.strip()
            if clean_v.lower() in {'nan', 'none', 'nat', 'null'}:
                clean_v = ""
        else:
            clean_v = v
        context[sanitize_key(k)] = clean_v

    if mapping:
        for tag, excel_col in mapping.items():
            if not excel_col:
                continue
            if isinstance(excel_col, str) and (excel_col == "__BLANK_UNDERLINE__" or excel_col == "UNDERLINE:"):
                context[tag] = _get_underline_str(tag)
            elif isinstance(excel_col, str) and excel_col.startswith("STATIC:"):
                context[tag] = excel_col[7:]
            elif excel_col in row_dict:
                val = row_dict[excel_col]
                if pd.isna(val):
                    clean_v = ""
                elif isinstance(val, str):
                    clean_v = val.strip()
                    if clean_v.lower() in {'nan', 'none', 'nat', 'null'}:
                        clean_v = ""
                else:
                    clean_v = val
                context[tag] = clean_v

    return context

def _ensure_table_borders(table):
    try:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        tblPr = table._tbl.tblPr
        tblBorders = tblPr.first_child_found_in("w:tblBorders")
        if tblBorders is None:
            tblBorders = OxmlElement('w:tblBorders')
            tblPr.append(tblBorders)

        borders = {
            'top': {'val': 'single', 'sz': '4', 'space': '0', 'color': '000000'},
            'left': {'val': 'single', 'sz': '4', 'space': '0', 'color': '000000'},
            'bottom': {'val': 'single', 'sz': '4', 'space': '0', 'color': '000000'},
            'right': {'val': 'single', 'sz': '4', 'space': '0', 'color': '000000'},
            'insideH': {'val': 'single', 'sz': '4', 'space': '0', 'color': '000000'},
            'insideV': {'val': 'single', 'sz': '4', 'space': '0', 'color': '000000'}
        }
        for border_name, border_props in borders.items():
            border_el = tblBorders.find(qn(f'w:{border_name}'))
            if border_el is None:
                border_el = OxmlElement(f'w:{border_name}')
                tblBorders.append(border_el)
            for k, v in border_props.items():
                border_el.set(qn(f'w:{k}'), str(v))
    except Exception:
        pass

def _get_farmer_val(farmer: dict, tag_key: str, mapping: dict | None = None) -> str:
    if not tag_key or not farmer:
        return ""

    mapped_col = mapping.get(tag_key) if mapping else None
    if mapped_col:
        if isinstance(mapped_col, str) and (mapped_col == "__BLANK_UNDERLINE__" or mapped_col == "UNDERLINE:"):
            return "____________"
        if isinstance(mapped_col, str) and mapped_col.startswith("STATIC:"):
            return mapped_col[7:]
        val = farmer.get(mapped_col, farmer.get(sanitize_key(mapped_col)))
        if val is not None and not pd.isna(val) and str(val).lower() not in {'nan', 'none', 'null'}:
            return str(val)

    if tag_key in farmer and farmer[tag_key] is not None and not pd.isna(farmer[tag_key]):
        return str(farmer[tag_key])

    tag_clean = sanitize_key(tag_key).lower()
    farmer_lower = {sanitize_key(k).lower(): v for k, v in farmer.items() if v is not None and not pd.isna(v)}
    if tag_clean in farmer_lower:
        return str(farmer_lower[tag_clean])

    if 'name' in tag_clean:
        if 'last' in tag_clean and 'last_name' in farmer_lower:
            return str(farmer_lower['last_name'])
        if 'first' in tag_clean and 'first_name' in farmer_lower:
            return str(farmer_lower['first_name'])
        if 'middle' in tag_clean and 'middle_name' in farmer_lower:
            return str(farmer_lower['middle_name'])
        if 'full_name' in farmer_lower:
            return str(farmer_lower['full_name'])

    if 'ref' in tag_clean or 'id' in tag_clean or 'no' in tag_clean:
        for r_key in ['rsbsa_no', 'reference_no', 'reference_no.', 'id', 'farmer_id', 'ref_no']:
            if r_key in farmer_lower:
                return str(farmer_lower[r_key])

    if 'birth' in tag_clean or 'bday' in tag_clean:
        for b_key in ['birthdate', 'birth_date', 'birthday', 'bday']:
            if b_key in farmer_lower:
                return str(farmer_lower[b_key])

    if 'gender' in tag_clean or 'sex' in tag_clean:
        for g_key in ['gender', 'sex']:
            if g_key in farmer_lower:
                return str(farmer_lower[g_key])

    if 'brgy' in tag_clean or 'barangay' in tag_clean:
        for br_key in ['barangay', 'brgy', 'home_barangay', 'farm_barangay']:
            if br_key in farmer_lower:
                return str(farmer_lower[br_key])

    return ""


def _get_row_sort_name(row_dict: dict, mapping: dict | None = None) -> str:
    """
    Extracts the normalized farmer name from a row dictionary for alphabetical sorting.
    Supports user tag mappings, split Last/First name columns, and common header variants.
    """
    # 1. Check user-defined tag mapping first
    if mapping:
        for tag in ['Full_Name', 'Fullname', 'Farmer_Name', 'Farmers_Name', 'Name', 'Beneficiary_Name', 'Beneficiary', 'FARMER_NAME']:
            mapped_col = mapping.get(tag)
            if mapped_col and mapped_col in row_dict and not str(mapped_col).startswith("STATIC:"):
                val = row_dict.get(mapped_col)
                if val is not None and not pd.isna(val) and str(val).strip().lower() not in {'nan', 'none', 'null'}:
                    return str(val).strip().upper()

    # 2. Case-insensitive dictionary of column headers
    lower_to_orig = {str(k).strip().lower().replace(' ', '_').replace('.', '').replace('-', '_'): k for k in row_dict.keys()}

    # Check split Last Name + First Name for standard Last, First sorting
    last_key = None
    first_key = None
    for l_cand in ['last_name', 'lastname', 'surname', 'family_name', 'lname']:
        if l_cand in lower_to_orig:
            last_key = lower_to_orig[l_cand]
            break
    for f_cand in ['first_name', 'firstname', 'given_name', 'fname']:
        if f_cand in lower_to_orig:
            first_key = lower_to_orig[f_cand]
            break

    if last_key:
        l_val = str(row_dict.get(last_key, '')).strip().upper()
        f_val = str(row_dict.get(first_key, '')).strip().upper() if first_key else ""
        if l_val and l_val not in {'NAN', 'NONE', 'NULL'}:
            return f"{l_val}, {f_val}".strip()

    # Check unified full name column headers
    for n_cand in ['full_name', 'fullname', 'farmer_name', 'farmername', 'farmers_name', 'farmersname', 'beneficiary_name', 'beneficiary', 'name']:
        if n_cand in lower_to_orig:
            val = str(row_dict.get(lower_to_orig[n_cand], '')).strip().upper()
            if val and val not in {'NAN', 'NONE', 'NULL'}:
                return val

    # Fallback: check any column with 'name' in it (excluding non-person columns)
    for clean_k, orig_k in lower_to_orig.items():
        if 'name' in clean_k and not any(ex in clean_k for ex in ['group', 'lender', 'farm', 'program']):
            val = str(row_dict.get(orig_k, '')).strip().upper()
            if val and val not in {'NAN', 'NONE', 'NULL'}:
                return val

    return ""


def _calculate_smart_column_widths(active_cols: list[dict], farmers_list: list[dict], total_width_mm: float = 160.0) -> list[float]:
    """
    Calculates content-aware column widths based on header length and max data row text length.
    Ensures short columns (Gender, Suffix, ID) take less space and long columns (Names, Address, Reference No.) get more space.
    Strictly normalizes all column widths so their sum is EXACTLY equal to total_width_mm (160mm),
    preventing any table overflow beyond the A4 page margins.
    """
    if not active_cols:
        return []

    scores = []
    for col_info in active_cols:
        h_name = str(col_info.get("header", ""))
        h_clean = h_name.lower().strip()
        max_val_len = 0
        for farmer in farmers_list[:50]:
            val = _get_farmer_val(farmer, h_name, None)
            if val:
                max_val_len = max(max_val_len, len(str(val)))

        effective_len = max(max_val_len, len(h_name))

        if 'gender' in h_clean or 'sex' in h_clean or 'suffix' in h_clean:
            score = 6.0
        elif 'ref' in h_clean or 'rsbsa' in h_clean or 'id' in h_clean or max_val_len >= 16:
            score = 19.0  # Cap score for Reference No so it gets ~38-40mm
        elif 'date' in h_clean or 'birth' in h_clean:
            score = 9.0   # ~18mm for Date
        elif 'name' in h_clean or 'last' in h_clean or 'first' in h_clean or 'middle' in h_clean:
            score = max(effective_len, 12) * 1.25 # Generous ~24-26mm for Names
        else:
            score = max(effective_len, 8) * 1.05

        scores.append(score)

    total_score = sum(scores)
    if total_score <= 0:
        total_score = 1.0

    raw_widths = [(s / total_score) * total_width_mm for s in scores]

    final_widths = []
    for idx, col_info in enumerate(active_cols):
        h_clean = str(col_info.get("header", "")).lower().strip()
        w = raw_widths[idx]
        if 'gender' in h_clean or 'sex' in h_clean or 'suffix' in h_clean:
            w = min(w, 12.0)
            w = max(w, 10.0)
        elif 'ref' in h_clean or 'rsbsa' in h_clean:
            w = min(w, 40.0)
            w = max(w, 35.0)
        elif 'date' in h_clean or 'birth' in h_clean:
            w = min(w, 20.0)
            w = max(w, 17.0)
        elif 'name' in h_clean:
            w = max(w, 22.0)
        final_widths.append(w)

    final_sum = sum(final_widths)
    normalized_widths = [(w / final_sum) * total_width_mm for w in final_widths]

    return normalized_widths

def _apply_xml_table_column_widths(table, col_widths_mm: list[float]):
    """
    Rebuilds the Word document XML <w:tblGrid> and <w:tcW> for every cell in every row.
    Also tightens cell margins (tcMar) to 1mm so text has maximum horizontal room without wrapping.
    """
    if not col_widths_mm:
        return

    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    widths_dxa = [int(w * 56.6929) for w in col_widths_mm]

    # 1. Rebuild <w:tblGrid>
    tblGrid = table._tbl.tblGrid
    tblGrid.clear()
    for dxa in widths_dxa:
        gridCol = OxmlElement('w:gridCol')
        gridCol.set(qn('w:w'), str(dxa))
        tblGrid.append(gridCol)

    # 2. Update <w:tcW> and set tight cell padding on every cell
    for row in table.rows:
        for col_idx, dxa in enumerate(widths_dxa):
            if col_idx < len(row.cells):
                cell = row.cells[col_idx]
                tcPr = cell._tc.get_or_add_tcPr()
                
                tcW = tcPr.find(qn('w:tcW'))
                if tcW is None:
                    tcW = OxmlElement('w:tcW')
                    tcPr.append(tcW)
                tcW.set(qn('w:w'), str(dxa))
                tcW.set(qn('w:type'), 'dxa')

                # Tighten cell padding (left/right margins) to 60 dxa (~1mm)
                tcMar = tcPr.find(qn('w:tcMar'))
                if tcMar is None:
                    tcMar = OxmlElement('w:tcMar')
                    tcPr.append(tcMar)
                for margin_name in ['left', 'right']:
                    node = tcMar.find(qn(f'w:{margin_name}'))
                    if node is None:
                        node = OxmlElement(f'w:{margin_name}')
                        tcMar.append(node)
                    node.set(qn('w:w'), '60')
                    node.set(qn('w:type'), 'dxa')

def _populate_transmittal_table(table, farmers_list: list[dict], transmittal_mapping: dict | None = None, transmittal_columns: list[dict] | None = None):
    if not farmers_list:
        return

    header_row_idx = 0
    sample_row_idx = 1 if len(table.rows) > 1 else 0

    try:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        row0_texts = set([c.text.strip() for c in table.rows[0].cells])
        if len(row0_texts) == 1 and len(table.rows) > 2:
            header_row_idx = 1
            sample_row_idx = 2
            header_row_indices = [0, 1]
        else:
            header_row_indices = [0]

        for r_idx in header_row_indices:
            trPr = table.rows[r_idx]._tr.get_or_add_trPr()
            if trPr.find(qn('w:tblHeader')) is None:
                trPr.append(OxmlElement('w:tblHeader'))
    except Exception:
        pass

    # Check for active transmittal columns
    active_cols = []
    if transmittal_columns:
        active_cols = [c for c in transmittal_columns if c.get('enabled', True)]
        def safe_order(c):
            val = c.get('order')
            try:
                return int(val)
            except Exception:
                return 999
        active_cols.sort(key=safe_order)

    if active_cols:
        target_count = len(active_cols)
        col_widths_mm = _calculate_smart_column_widths(active_cols, farmers_list, total_width_mm=160.0)

        # 1. If table has more columns than target_count, trim extra cells from all rows
        for row in table.rows:
            while len(row.cells) > target_count:
                tc = row.cells[-1]._tc
                tc.getparent().remove(tc)

        tblGrid = table._tbl.tblGrid
        if tblGrid is not None:
            while len(tblGrid) > target_count:
                tblGrid.remove(tblGrid[-1])

        # 2. If table has fewer columns than target_count, add extra columns
        while len(table.rows[0].cells) < target_count:
            col_idx = len(table.rows[0].cells)
            w_mm = col_widths_mm[min(col_idx, target_count - 1)]
            table.add_column(Mm(w_mm))

        header_cells = table.rows[header_row_idx].cells
        for col_i, col_info in enumerate(active_cols):
            if col_i < len(header_cells):
                header_cells[col_i].text = str(col_info["header"])

        has_sample_row = len(table.rows) > sample_row_idx
        col_formats = []
        for col_idx in range(target_count):
            fmt = {'name': None, 'size': None, 'bold': None}
            try:
                cell = table.cell(sample_row_idx, col_idx)
                if cell.paragraphs and cell.paragraphs[0].runs:
                    r = cell.paragraphs[0].runs[0]
                    fmt['name'] = r.font.name
                    fmt['size'] = r.font.size
                    fmt['bold'] = r.bold if has_sample_row else False
            except Exception:
                pass
            col_formats.append(fmt)

        for i, farmer in enumerate(farmers_list):
            if i == 0 and has_sample_row:
                row_cells = table.rows[sample_row_idx].cells
            else:
                row_cells = table.add_row().cells

            for col_i, col_info in enumerate(active_cols):
                if col_i >= len(row_cells):
                    break
                header_name = col_info["header"]
                val = _get_farmer_val(farmer, header_name, transmittal_mapping)
                row_cells[col_i].text = str(val) if val is not None else ""

        # Dynamic font size scaling based on column count to ensure zero wrapping
        from docx.shared import Pt
        if target_count <= 5:
            auto_font_size = Pt(9.5)
        elif target_count <= 7:
            auto_font_size = Pt(8.5)
        elif target_count <= 9:
            auto_font_size = Pt(7.5)
        else:
            auto_font_size = Pt(6.5)

        for row in table.rows:
            for col_idx, cell in enumerate(row.cells):
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    for r in p.runs:
                        r.font.size = auto_font_size
                        if col_idx < len(col_formats) and col_formats[col_idx]['name']:
                            r.font.name = col_formats[col_idx]['name']

        # Rebuild XML tblGrid and cell widths to force Word & LibreOffice to respect dynamic column widths
        _apply_xml_table_column_widths(table, col_widths_mm)

        _ensure_table_borders(table)
        return

    col_tags = {}
    header_cells = table.rows[header_row_idx].cells
    for col_i, cell in enumerate(header_cells):
        raw_text = cell.text.strip()
        matches = re.findall(r'\{\{\s*([^\}\s]+)\s*\}\}', raw_text)
        if matches:
            tag_name = matches[0]
            col_tags[col_i] = tag_name
            mapped_val = transmittal_mapping.get(tag_name) if transmittal_mapping else None
            if mapped_val:
                display_label = mapped_val[7:] if mapped_val.startswith("STATIC:") else mapped_val
            else:
                display_label = tag_name.replace('_', ' ').title()
            cell.text = display_label
        else:
            col_tags[col_i] = raw_text

    has_sample_row = len(table.rows) > sample_row_idx
    col_formats = []
    for col_idx in range(len(table.columns)):
        fmt = {'name': None, 'size': None, 'bold': None}
        try:
            cell = table.cell(sample_row_idx, col_idx)
            if cell.paragraphs and cell.paragraphs[0].runs:
                r = cell.paragraphs[0].runs[0]
                fmt['name'] = r.font.name
                fmt['size'] = r.font.size
                fmt['bold'] = r.bold if has_sample_row else False
        except Exception:
            pass
        col_formats.append(fmt)

    for i, farmer in enumerate(farmers_list):
        if i == 0 and has_sample_row:
            row_cells = table.rows[sample_row_idx].cells
        else:
            row_cells = table.add_row().cells

        for col_i in range(len(table.columns)):
            if col_i >= len(row_cells):
                break
            tag_key = col_tags.get(col_i, "")
            val = _get_farmer_val(farmer, tag_key, transmittal_mapping)
            row_cells[col_i].text = str(val) if val is not None else ""

        for col_idx, cell in enumerate(row_cells):
            if col_idx < len(col_formats):
                fmt = col_formats[col_idx]
                for p in cell.paragraphs:
                    for r in p.runs:
                        if fmt['name']:
                            r.font.name = fmt['name']
                        if fmt['size']:
                            r.font.size = fmt['size']
                        if fmt['bold'] is not None:
                            r.bold = fmt['bold']

    _ensure_table_borders(table)

def _find_libreoffice() -> str | None:
    lo_bin = shutil.which("libreoffice") or shutil.which("soffice")
    if not lo_bin:
        for loc in [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]:
            if os.path.exists(loc):
                lo_bin = loc
                break
    return lo_bin


def _render_one_docx(args: tuple):
    """
    Parallel worker function (pure Python, 100% thread/process safe):
      1. Fills the .docx template for a single farmer row or transmittal using docxtpl.
      2. Injects transmittal table dynamically if farmers list is present.
      3. Saves filled .docx to temp_docx_path.
    """
    worker_idx, context, template_file, temp_docx_path_str = args
    temp_docx_path = Path(temp_docx_path_str)
    try:
        doc = DocxTemplate(template_file)
        doc.init_docx()

        if 'farmers' in context and doc.docx.tables:
            _populate_transmittal_table(
                doc.docx.tables[0],
                context['farmers'],
                context.get('_transmittal_mapping'),
                context.get('_transmittal_columns')
            )

        doc.render(context)

        for section in doc.docx.sections:
            section.page_width  = Mm(210)
            section.page_height = Mm(297)

        doc.save(temp_docx_path)
        return worker_idx, temp_docx_path, None
    except Exception as exc:
        return worker_idx, None, str(exc)


def _convert_one_lo(args: tuple):
    """
    Parallel worker function for LibreOffice conversion with strict timeout.
    """
    worker_idx, docx_path_str, temp_dir_str, lo_bin, lo_profile_url, temp_pdf_path_str = args
    docx_path = Path(docx_path_str)
    temp_pdf_path = Path(temp_pdf_path_str)
    temp_dir = Path(temp_dir_str)
    try:
        result = subprocess.run(
            [
                lo_bin,
                f"-env:UserInstallation={lo_profile_url}",
                "--headless",
                "--convert-to", "pdf",
                str(docx_path.resolve()),
                "--outdir", str(temp_dir.resolve()),
            ],
            capture_output=True,
            text=True,
            timeout=300,
            check=False,
        )
        if result.returncode == 0 and temp_pdf_path.exists() and temp_pdf_path.stat().st_size > 0:
            return worker_idx, temp_pdf_path, None
        return worker_idx, None, f"LibreOffice rc={result.returncode}: {result.stderr.strip()}"
    except subprocess.TimeoutExpired:
        return worker_idx, None, "LibreOffice conversion timed out after 300s"
    except Exception as exc:
        return worker_idx, None, str(exc)


def _compose_group_forms(args: tuple):
    """
    Combines individual farmer application forms for a single group into one consolidated multi-page DOCX.
    Sets page_break_before on the first paragraph of every appended document to ensure that each
    form strictly begins at the top of a new page, preventing headers from spilling onto previous pages.
    """
    group_idx, form_paths, out_bundle_docx_str = args
    if not form_paths:
        return group_idx, None, "No forms in group"
    try:
        from docxcompose.composer import Composer
        from docx import Document
        from docx.shared import Pt
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        master = Document(form_paths[0])
        # Minimize trailing empty paragraph height in master
        if master.paragraphs and not master.paragraphs[-1].text.strip():
            master.paragraphs[-1].paragraph_format.space_before = Pt(0)
            master.paragraphs[-1].paragraph_format.space_after = Pt(0)
            master.paragraphs[-1].paragraph_format.line_spacing = Pt(1)

        if len(form_paths) > 1:
            composer = Composer(master)
            for p in form_paths[1:]:
                doc_to_append = Document(p)
                # Minimize trailing empty paragraph height in appended document
                if doc_to_append.paragraphs and not doc_to_append.paragraphs[-1].text.strip():
                    doc_to_append.paragraphs[-1].paragraph_format.space_before = Pt(0)
                    doc_to_append.paragraphs[-1].paragraph_format.space_after = Pt(0)
                    doc_to_append.paragraphs[-1].paragraph_format.line_spacing = Pt(1)

                # Ensure the appended document strictly starts at the top of a new page
                if doc_to_append.paragraphs:
                    doc_to_append.paragraphs[0].paragraph_format.page_break_before = True
                elif doc_to_append.element.body and doc_to_append.element.body[0].tag.endswith('tbl'):
                    p_el = OxmlElement('w:p')
                    pPr = OxmlElement('w:pPr')
                    pPr.append(OxmlElement('w:pageBreakBefore'))
                    sp = OxmlElement('w:spacing')
                    sp.set(qn('w:before'), '0')
                    sp.set(qn('w:after'), '0')
                    pPr.append(sp)
                    p_el.append(pPr)
                    doc_to_append.element.body.insert(0, p_el)

                composer.append(doc_to_append)

            composer.save(out_bundle_docx_str)
        else:
            master.save(out_bundle_docx_str)
        return group_idx, Path(out_bundle_docx_str), None
    except Exception as exc:
        return group_idx, None, str(exc)


def _convert_all_word_com(
    docx_tasks: list[tuple[any, Path, Path]],
    log_func,
    progress_callback=None,
    total_jobs: int = 1,
    start_percent: int = 35,
    end_percent: int = 85
):
    """
    High-performance, deadlock-free MS Word COM conversion:
    Uses a single persistent background Word instance with screen updating,
    pagination, and spelling checks disabled for maximum conversion throughput.
    """
    import win32com.client as win32
    import pythoncom

    pythoncom.CoInitialize()
    word = None
    pdf_results: dict[any, Path] = {}
    error_count = 0
    done = 0

    def start_word():
        try:
            w = win32.DispatchEx("Word.Application")
        except Exception:
            try:
                import win32com
                gen_dir = getattr(win32com, '__gen_path__', None)
                if gen_dir and os.path.exists(gen_dir):
                    shutil.rmtree(gen_dir, ignore_errors=True)
            except Exception:
                pass
            w = win32.DispatchEx("Word.Application")

        w.Visible = False
        w.DisplayAlerts = 0
        w.ScreenUpdating = False
        try:
            w.Options.Pagination = False
            w.Options.CheckSpellingAsYouType = False
            w.Options.CheckGrammarAsYouType = False
            w.Options.SaveNormalPrompt = False
        except Exception:
            pass
        try:
            w.NormalTemplate.Saved = True
        except Exception:
            pass
        return w

    try:
        word = start_word()
        for i, (task_id, docx_path, pdf_path) in enumerate(docx_tasks):
            # Recycle Word every 300 conversions to ensure pristine memory state
            if i > 0 and i % 300 == 0:
                try:
                    if word:
                        try:
                            word.NormalTemplate.Saved = True
                        except Exception:
                            pass
                        word.Quit(0)
                except Exception:
                    pass
                word = start_word()

            done += 1
            if not docx_path or not Path(docx_path).exists():
                error_count += 1
                log_func(f"Conversion [{done}/{total_jobs}] FAILED: Docx file was not created", "warning")
                continue

            try:
                doc = word.Documents.Open(
                    str(Path(docx_path).resolve()),
                    ReadOnly=True,
                    ConfirmConversions=False,
                    AddToRecentFiles=False
                )
                doc.SaveAs(str(Path(pdf_path).resolve()), FileFormat=17)
                doc.Close(0)

                if Path(pdf_path).exists() and Path(pdf_path).stat().st_size > 0:
                    pdf_results[task_id] = Path(pdf_path)
                else:
                    error_count += 1
                    log_func(f"Conversion [{done}/{total_jobs}] FAILED: PDF output not generated", "warning")
            except Exception as conv_err:
                error_count += 1
                log_func(f"Conversion [{done}/{total_jobs}] FAILED: {conv_err}", "warning")
                try:
                    if word:
                        word.Quit(0)
                except Exception:
                    pass
                word = start_word()

            if progress_callback:
                pct = start_percent + int((done / max(total_jobs, 1)) * (end_percent - start_percent))
                progress_callback({
                    "type": "progress",
                    "current": 0,
                    "total": total_jobs,
                    "percent": pct,
                    "failed": error_count,
                    "status_text": f"Converting PDF bundles ({done}/{total_jobs})..."
                })

    finally:
        if word:
            try:
                try:
                    word.NormalTemplate.Saved = True
                except Exception:
                    pass
                word.Quit(0)
            except Exception:
                pass
        pythoncom.CoUninitialize()

    return pdf_results, error_count


def run_batch_generation(
    excel_file: str,
    template_file: str,
    transmittal_template: str,
    output_dir: str = "output",
    temp_dir: str = "temp_docs",
    primary_group_col: str | None = None,
    secondary_group_col: str | None = None,
    bundle_group_col: str | None = None,
    test_limit: int | None = None,
    max_workers: int = 4,
    template_mapping: dict | None = None,
    transmittal_mapping: dict | None = None,
    transmittal_columns: list[dict] | None = None,
    progress_callback=None
):
    """
    Executes high-speed consolidated batch generation.
    Option 2 Pipeline:
      1. Parallel .docx rendering of all farmer application forms and transmittals.
      2. High-speed docx bundle consolidation per Barangay (forms share identical A4 margins).
      3. Minimal PDF conversions (only 2 conversions per Barangay: Transmittal + Forms Bundle).
      4. Instant PyPDF packaging of final output bundles.
    """
    out_path = Path(output_dir)
    tmp_path = Path(temp_dir)
    out_path.mkdir(parents=True, exist_ok=True)
    tmp_path.mkdir(parents=True, exist_ok=True)

    def log(msg, status="info"):
        if progress_callback:
            progress_callback({
                "type": "log",
                "status": status,
                "message": msg,
                "timestamp": datetime.now().strftime("%H:%M:%S")
            })

    log(f"Starting batch generation from '{Path(excel_file).name}'")
    df = pd.read_excel(excel_file)

    if test_limit and test_limit > 0:
        log(f"Test Mode: Limiting to first {test_limit} rows.")
        df = df.head(test_limit)

    for col in df.columns:
        if 'date' in str(col).lower() or 'birth' in str(col).lower():
            df[col] = pd.to_datetime(df[col], errors='coerce')
    for col in df.select_dtypes(include=['datetime64[ns]']).columns:
        df[col] = df[col].dt.strftime('%m/%d/%Y')

    df = df.fillna("")
    col_dict = {str(c).strip().lower(): str(c) for c in df.columns}

    def resolve_col(user_choice, default_names):
        if user_choice and user_choice.strip().lower() in col_dict:
            return col_dict[user_choice.strip().lower()]
        for d in default_names:
            if d.lower() in col_dict:
                return col_dict[d.lower()]
        return None

    p_col = resolve_col(primary_group_col, ["province", "region", "district"])
    s_col = resolve_col(secondary_group_col, ["municipality", "city", "town"])
    b_col = resolve_col(bundle_group_col, ["barangay", "association", "group", "cooperative"])

    group_keys = [c for c in [p_col, s_col, b_col] if c is not None]

    if not group_keys:
        log("No valid grouping columns found. Treating dataset as single bundle.", "warning")
        grouped = [("All_Records", df)]
    else:
        log(f"Grouping dataset by: {' > '.join(group_keys)}")
        grouped = df.groupby(group_keys, sort=True)

    group_specs: list[dict] = []
    total_farmers = 0
    all_render_jobs: list[tuple] = []

    for g_idx, (group_key_val, group_df) in enumerate(grouped):
        if not isinstance(group_key_val, tuple):
            group_key_tuple = (group_key_val,)
        else:
            group_key_tuple = group_key_val

        prov_name = str(group_key_tuple[0]) if len(group_key_tuple) > 0 else ""
        muni_name = str(group_key_tuple[1]) if len(group_key_tuple) > 1 else ""
        brgy_name = str(group_key_tuple[2]) if len(group_key_tuple) > 2 else prov_name

        if not any(group_key_tuple):
            continue

        group_df = group_df.copy()
        group_df['_sort_key'] = group_df.apply(
            lambda r: _get_row_sort_name(r.to_dict(), template_mapping or transmittal_mapping),
            axis=1
        )
        group_df = group_df.sort_values(by='_sort_key', ascending=True, kind='mergesort')
        group_df = group_df.drop(columns=['_sort_key'])

        farmers_list = []
        for _, row in group_df.iterrows():
            r_d = row.to_dict()
            ctx = _build_context(r_d, transmittal_mapping)
            ctx['Barangay'] = brgy_name
            if 'Full_Name' not in ctx:
                ctx['Full_Name'] = ctx.get('Fullname', ctx.get('Name', ''))
            if 'Birthday' not in ctx:
                ctx['Birthday'] = ctx.get('Birthdate', ctx.get('Birth_Date', ''))
            if 'Reference_No' not in ctx:
                ctx['Reference_No'] = ctx.get('Reference_No.', ctx.get('ID', ''))
            for r_k, r_v in r_d.items():
                if r_k not in ctx:
                    ctx[r_k] = r_v
            farmers_list.append(ctx)

        farmers_list.sort(key=lambda x: str(x.get('Full_Name') or x.get('Fullname') or x.get('Name') or '').strip().upper())

        trans_docx = tmp_path / f"trans_{g_idx:04d}.docx"
        trans_pdf = tmp_path / f"trans_{g_idx:04d}.pdf"
        trans_ctx = {
            'barangay': brgy_name,
            'municipality': muni_name,
            'province': prov_name,
            'farmers': farmers_list,
            '_transmittal_mapping': transmittal_mapping,
            '_transmittal_columns': transmittal_columns
        }
        if transmittal_mapping:
            for tag, val in transmittal_mapping.items():
                if val.startswith('STATIC:'):
                    trans_ctx[tag] = val[7:]

        all_render_jobs.append((f"trans_{g_idx}", trans_ctx, transmittal_template, str(trans_docx.resolve())))

        form_docx_paths = []
        for f_idx, (_, row) in enumerate(group_df.iterrows()):
            f_docx = tmp_path / f"form_{g_idx:04d}_{f_idx:04d}.docx"
            form_docx_paths.append(str(f_docx.resolve()))
            all_render_jobs.append((
                f"form_{g_idx}_{f_idx}",
                _build_context(row.to_dict(), template_mapping),
                template_file,
                str(f_docx.resolve())
            ))

        forms_bundle_docx = tmp_path / f"forms_bundle_{g_idx:04d}.docx"
        forms_bundle_pdf = tmp_path / f"forms_bundle_{g_idx:04d}.pdf"

        group_specs.append({
            "group_idx": g_idx,
            "prov": prov_name,
            "muni": muni_name,
            "brgy": brgy_name,
            "trans_docx": trans_docx,
            "trans_pdf": trans_pdf,
            "form_docx_paths": form_docx_paths,
            "forms_bundle_docx": forms_bundle_docx,
            "forms_bundle_pdf": forms_bundle_pdf,
            "count": len(group_df)
        })
        total_farmers += len(group_df)

    total_render_tasks = len(all_render_jobs)
    error_count = 0
    lo_bin = _find_libreoffice()

    log(f"Phase 1/4: Rendering {total_render_tasks} templates across {max_workers} parallel workers...")
    if progress_callback:
        progress_callback({
            "type": "progress",
            "current": 0,
            "total": total_farmers,
            "percent": 0,
            "failed": 0,
            "status_text": "Filling document templates..."
        })

    with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(_render_one_docx, job): job[0] for job in all_render_jobs}
        r_done = 0
        for future in concurrent.futures.as_completed(futures):
            job_id, docx_path, err = future.result()
            r_done += 1
            if err:
                log(f"Template [{r_done}/{total_render_tasks}] FAILED: {err}", "warning")
                error_count += 1
            if progress_callback and (r_done % 25 == 0 or r_done == total_render_tasks):
                progress_callback({
                    "type": "progress",
                    "current": 0,
                    "total": total_farmers,
                    "percent": int((r_done / max(total_render_tasks, 1)) * 25),
                    "failed": error_count,
                    "status_text": f"Filling templates ({r_done}/{total_render_tasks})..."
                })

    log(f"Phase 2/4: Composing forms into {len(group_specs)} consolidated Barangay bundles...")
    compose_jobs = [
        (g["group_idx"], g["form_docx_paths"], str(g["forms_bundle_docx"].resolve()))
        for g in group_specs if g["form_docx_paths"]
    ]

    with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as pool:
        c_futures = {pool.submit(_compose_group_forms, cj): cj[0] for cj in compose_jobs}
        c_done = 0
        for future in concurrent.futures.as_completed(c_futures):
            g_idx, b_path, err = future.result()
            c_done += 1
            if err:
                log(f"Bundle composition warning: {err}", "warning")
            if progress_callback and (c_done % 5 == 0 or c_done == len(compose_jobs)):
                pct = 25 + int((c_done / max(len(compose_jobs), 1)) * 10)
                progress_callback({
                    "type": "progress",
                    "current": 0,
                    "total": total_farmers,
                    "percent": pct,
                    "failed": error_count,
                    "status_text": f"Composing bundles ({c_done}/{len(compose_jobs)})..."
                })

    conv_tasks: list[tuple[any, Path, Path]] = []
    for g in group_specs:
        if g["trans_docx"].exists():
            conv_tasks.append((f"trans_{g['group_idx']}", g["trans_docx"], g["trans_pdf"]))
        if g["forms_bundle_docx"].exists():
            conv_tasks.append((f"forms_{g['group_idx']}", g["forms_bundle_docx"], g["forms_bundle_pdf"]))

    log(f"Phase 3/4: Converting {len(conv_tasks)} consolidated bundle files to PDF...")

    if lo_bin:
        lo_profile_base = (tmp_path / "lo_profiles").resolve()
        lo_profile_base.mkdir(parents=True, exist_ok=True)
        with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as pool:
            lo_futures = {}
            for t_id, d_path, p_path in conv_tasks:
                lo_profile_dir = lo_profile_base / f"lo_{t_id}"
                lo_profile_dir.mkdir(parents=True, exist_ok=True)
                lo_args = (
                    t_id,
                    str(d_path.resolve()),
                    str(tmp_path.resolve()),
                    lo_bin,
                    lo_profile_dir.as_uri(),
                    str(p_path.resolve())
                )
                lo_futures[pool.submit(_convert_one_lo, lo_args)] = t_id

            conv_done = 0
            for future in concurrent.futures.as_completed(lo_futures):
                t_id, pdf_path, err = future.result()
                conv_done += 1
                if err:
                    log(f"Conversion [{conv_done}/{len(conv_tasks)}] FAILED: {err}", "warning")
                    error_count += 1
                if progress_callback and (conv_done % 2 == 0 or conv_done == len(conv_tasks)):
                    pct = 35 + int((conv_done / max(len(conv_tasks), 1)) * 50)
                    progress_callback({
                        "type": "progress",
                        "current": 0,
                        "total": total_farmers,
                        "percent": pct,
                        "failed": error_count,
                        "status_text": f"Converting PDF bundles ({conv_done}/{len(conv_tasks)})..."
                    })
    else:
        word_results, word_errors = _convert_all_word_com(
            docx_tasks=conv_tasks,
            log_func=log,
            progress_callback=progress_callback,
            total_jobs=len(conv_tasks),
            start_percent=35,
            end_percent=85
        )
        error_count += word_errors

    log(f"Phase 4/4: Packaging {len(group_specs)} Barangay PDF bundles...")
    total_merged = 0

    for b_idx, g in enumerate(group_specs, 1):
        safe_prov = safe_filename(g["prov"])
        safe_muni = safe_filename(g["muni"])
        safe_brgy = safe_filename(g["brgy"])

        if safe_prov and safe_muni:
            target_dir = out_path / f"{safe_prov}, {safe_muni}"
        elif safe_prov:
            target_dir = out_path / safe_prov
        else:
            target_dir = out_path

        target_dir.mkdir(parents=True, exist_ok=True)
        merged_path = target_dir / f"Bundle_{safe_brgy}_{g['count']}_Forms.pdf"
        log(f"Packaging bundle [{b_idx}/{len(group_specs)}]: {merged_path.name}")

        merger = PdfWriter()
        if g["trans_pdf"].exists():
            merger.append(str(g["trans_pdf"]))
        if g["forms_bundle_pdf"].exists():
            merger.append(str(g["forms_bundle_pdf"]))

        with open(merged_path, "wb") as f:
            merger.write(f)

        total_merged += g["count"]

        if progress_callback:
            pct = 85 + int((b_idx / max(len(group_specs), 1)) * 14)
            progress_callback({
                "type": "progress",
                "current": total_merged,
                "total": total_farmers,
                "percent": min(pct, 99),
                "failed": error_count,
                "status_text": f"Packaging bundle {b_idx}/{len(group_specs)}..."
            })

    if progress_callback:
        progress_callback({
            "type": "progress",
            "current": total_farmers,
            "total": total_farmers,
            "percent": 100,
            "failed": error_count,
            "status_text": "Batch Processing Complete!"
        })

    if error_count == 0:
        log("No errors recorded — cleaning up temporary files.")
        shutil.rmtree(tmp_path, ignore_errors=True)

    log(f"Batch execution finished! {total_merged} application forms merged into '{out_path}/'.", "success")

    return {
        "total_rendered": total_farmers - error_count,
        "total_failed": error_count,
        "total_merged": total_merged,
        "output_directory": str(out_path.resolve())
    }
