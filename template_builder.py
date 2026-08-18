import re
import os
import subprocess
from pathlib import Path
import pandas as pd
import docx
from docxtpl import DocxTemplate

def sanitize_key(k: str) -> str:
    """Convert an Excel column header into a valid Jinja2 variable name."""
    s = str(k).strip()
    s = s.replace(' ', '_').replace('.', '').replace('-', '_').replace('/', '_')
    return ''.join(c for c in s if c.isalnum() or c == '_')

def extract_excel_headers(excel_path: str) -> list[dict]:
    """Reads Excel file and returns column metadata with sanitized Jinja keys."""
    try:
        df = pd.read_excel(excel_path, nrows=2)
        headers = []
        for col in df.columns:
            orig = str(col).strip()
            sanitized = sanitize_key(orig)
            headers.append({
                "original": orig,
                "sanitized": sanitized,
                "tag": f"{{{{ {sanitized} }}}}"
            })
        return headers
    except Exception as e:
        raise ValueError(f"Failed to read Excel headers: {str(e)}")

def extract_docx_tags(docx_path: str) -> list[str]:
    """Inspects a Word document and extracts all {{ variable }} tags in top-to-bottom document order."""
    try:
        doc = docx.Document(docx_path)
        found_tags = []
        seen = set()

        def extract_text_tags(text: str):
            matches = re.findall(r'\{\{\s*([^\}\s]+)\s*\}\}', text)
            for tag in matches:
                if tag not in seen:
                    seen.add(tag)
                    found_tags.append(tag)

        # 1. Body paragraphs & tables
        for element in doc.element.body:
            if element.tag.endswith('p'):
                p = docx.text.paragraph.Paragraph(element, doc)
                extract_text_tags(p.text)
            elif element.tag.endswith('tbl'):
                table = docx.table.Table(element, doc)
                for row in table.rows:
                    for cell in row.cells:
                        for cell_p in cell.paragraphs:
                            extract_text_tags(cell_p.text)

        # 2. Section headers & footers
        for section in doc.sections:
            if section.header:
                for p in section.header.paragraphs:
                    extract_text_tags(p.text)
                for t in section.header.tables:
                    for row in t.rows:
                        for cell in row.cells:
                            for cell_p in cell.paragraphs:
                                extract_text_tags(cell_p.text)
            if section.footer:
                for p in section.footer.paragraphs:
                    extract_text_tags(p.text)
                for t in section.footer.tables:
                    for row in t.rows:
                        for cell in row.cells:
                            for cell_p in cell.paragraphs:
                                extract_text_tags(cell_p.text)

        return found_tags
    except Exception as e:
        raise ValueError(f"Failed to inspect docx tags: {str(e)}")

def auto_match_tags(excel_headers: list[str], docx_tags: list[str]) -> dict[str, str]:
    """
    Intelligently connects Word document Jinja tags to Excel column headers.
    Returns a dictionary mapping { docx_tag: excel_header_name }.
    """
    mapping = {}
    
    # Custom explicit mappings for standard insurance fields
    aliases = {
        'birthdate': ['birth date', 'bday', 'birthdate', 'dob'],
        'contact_number': ['contact no', 'contact number', 'phone', 'mobile'],
        'home_barangay': ['barangay', 'brgy'],
        'home_municipality': ['municipality', 'city', 'town'],
        'home_province': ['province'],
        'home_sitiopurok': ['sitio', 'purok', 'street'],
        'farm_barangay': ['barangay', 'brgy'],
        'farm_municipality': ['municipality'],
        'farm_province': ['province'],
        'farm_sitiopurok': ['sitio', 'purok'],
        'cfitf_no': ['reference no', 'reference no.', 'cfitf no'],
        'rsbsa_no': ['farmer id', 'rsbsa no'],
        'area': ['coconutarea', 'area'],
        'pb_bday': ['birth date'],
        'pb_rel': ['civil status']
    }

    for tag in docx_tags:
        t_clean = tag.lower().strip()
        matched = False
        
        # 1. Check exact match
        for h in excel_headers:
            h_clean = h.lower().strip()
            if t_clean == h_clean or t_clean == h_clean.replace(' ', '_'):
                mapping[tag] = h
                matched = True
                break
        if matched:
            continue

        # 2. Check aliases
        if t_clean in aliases:
            for alias in aliases[t_clean]:
                for h in excel_headers:
                    if alias in h.lower():
                        mapping[tag] = h
                        matched = True
                        break
                if matched:
                    break
        if matched:
            continue

        # 3. Check fuzzy containment
        t_norm = t_clean.replace('_', '').replace(' ', '')
        for h in excel_headers:
            h_norm = h.lower().replace('_', '').replace(' ', '').replace('.', '')
            if t_norm == h_norm or t_norm in h_norm or h_norm in t_norm:
                mapping[tag] = h
                matched = True
                break

    return mapping

def render_docx_to_pdf_preview(docx_path: str, temp_dir: str, output_name: str) -> str:
    """Converts docx to high-fidelity PDF via LibreOffice for 1:1 exact visual preview."""
    try:
        out_dir = Path(temp_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        pdf_path = out_dir / f"{output_name}.pdf"

        profile_dir = (out_dir / f"profile_{output_name}").resolve()
        profile_dir.mkdir(parents=True, exist_ok=True)
        profile_url = profile_dir.as_uri()

        subprocess.run(
            [
                "libreoffice",
                f"-env:UserInstallation={profile_url}",
                "--headless",
                "--convert-to", "pdf",
                docx_path,
                "--outdir", str(out_dir),
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=True
        )

        generated_pdf = out_dir / f"{Path(docx_path).stem}.pdf"
        if generated_pdf.exists() and generated_pdf != pdf_path:
            if pdf_path.exists():
                os.remove(pdf_path)
            os.rename(generated_pdf, pdf_path)

        return str(pdf_path)
    except Exception as e:
        raise ValueError(f"Failed to convert docx to preview PDF: {str(e)}")

def render_sample_pdf_preview(
    excel_path: str,
    docx_path: str,
    mapping: dict[str, str],
    temp_dir: str,
    output_name: str
) -> str:
    """
    Renders sample data from Excel into the Word template using docxtpl,
    then converts to PDF for live 1:1 previewing in the browser.
    """
    try:
        docx_tags = extract_docx_tags(docx_path)
        context = {tag: "" for tag in docx_tags}

        df_sample = None
        row_dict = {}
        if excel_path and os.path.exists(excel_path):
            try:
                df_sample = pd.read_excel(excel_path, nrows=5)
                if not df_sample.empty:
                    row_dict = df_sample.iloc[0].to_dict()
            except Exception:
                pass

        for tag in docx_tags:
            mapped_val = mapping.get(tag)
            if mapped_val:
                if isinstance(mapped_val, str) and mapped_val.startswith("STATIC:"):
                    context[tag] = mapped_val[7:]
                elif mapped_val in row_dict:
                    val = row_dict[mapped_val]
                    context[tag] = str(val) if pd.notna(val) else ""

        # Build sample farmers if rendering transmittal template
        if 'transmittal' in output_name.lower() or 'transmittal' in docx_path.lower():
            farmers = []
            if df_sample is not None and not df_sample.empty:
                for _, s_row in df_sample.iterrows():
                    r_d = s_row.to_dict()
                    def find_val(candidates):
                        for c in candidates:
                            for r_k, r_v in r_d.items():
                                if str(r_k).strip().lower() == c.lower() and pd.notna(r_v) and str(r_v).strip():
                                    return str(r_v).strip()
                        return ""
                    
                    fn = find_val(['Full Name', 'Full_Name', 'Farmer Name', 'Name'])
                    if not fn:
                        fname = find_val(['First Name', 'First_Name'])
                        lname = find_val(['Last Name', 'Last_Name'])
                        if fname or lname:
                            fn = f"{lname}, {fname}".strip(', ')
                    
                    bd = find_val(['Birth Date', 'Birth_Date', 'Birthdate', 'Birthday', 'DOB', 'Bday'])
                    if bd and hasattr(s_row.get('Birth Date', ''), 'strftime'):
                        bd = s_row.get('Birth Date').strftime('%m/%d/%Y')
                    
                    farmers.append({
                        'Barangay': find_val(['Barangay', 'Brgy']),
                        'Full_Name': fn or 'DELA CRUZ, JUAN A.',
                        'Birthday': str(bd) if bd else '01/01/1980',
                        'Gender': find_val(['Gender', 'Sex']) or 'Male',
                        'Reference_No': find_val(['Reference No.', 'Reference No', 'Reference_No', 'Farmer ID', 'ID']) or 'REF-00001'
                    })
            context['farmers'] = farmers

        tpl = DocxTemplate(docx_path)
        tpl.render(context)

        # Inject sample table rows if farmers context exists
        if 'farmers' in context and tpl.docx.tables:
            from generator_engine import populate_transmittal_table
            populate_transmittal_table(tpl.docx, context['farmers'])

        out_dir = Path(temp_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        filled_docx = out_dir / f"{output_name}_sample.docx"
        tpl.save(filled_docx)

        return render_docx_to_pdf_preview(str(filled_docx), temp_dir, output_name)
    except Exception as e:
        raise ValueError(f"Failed to render sample PDF preview: {str(e)}")

