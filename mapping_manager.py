import os
import re
import json
import hashlib
from datetime import datetime
from pathlib import Path
from template_builder import auto_match_tags

MAPPINGS_DIR = os.path.abspath("mappings")
Path(MAPPINGS_DIR).mkdir(parents=True, exist_ok=True)

def normalize_header(h: str) -> str:
    """Standardizes a column header string for robust similarity matching."""
    s = str(h).strip().lower()
    s = re.sub(r'[\s_\-\.\/]+', ' ', s).strip()
    return s

def compute_header_similarity(current_headers: list[str], saved_headers: list[str]) -> float:
    """
    Computes a similarity score between 0.0 and 1.0 between two sets of Excel column headers.
    Uses normalized token matching and Dice overlap coefficient.
    """
    if not current_headers or not saved_headers:
        return 0.0

    norm_curr = [normalize_header(h) for h in current_headers if str(h).strip()]
    norm_saved = [normalize_header(h) for h in saved_headers if str(h).strip()]

    if not norm_curr or not norm_saved:
        return 0.0

    # Match each current header against saved headers
    matched_saved_indices = set()
    matches_count = 0

    for c in norm_curr:
        matched = False
        # 1. Exact normalized match
        for s_idx, s in enumerate(norm_saved):
            if s_idx in matched_saved_indices:
                continue
            if c == s:
                matched_saved_indices.add(s_idx)
                matches_count += 1
                matched = True
                break
        if matched:
            continue

        # 2. Substring / partial match (for length >= 4)
        for s_idx, s in enumerate(norm_saved):
            if s_idx in matched_saved_indices:
                continue
            if len(c) >= 4 and len(s) >= 4 and (c in s or s in c):
                matched_saved_indices.add(s_idx)
                matches_count += 1
                matched = True
                break

    # Dice coefficient: 2 * |A ∩ B| / (|A| + |B|)
    dice_score = (2.0 * matches_count) / (len(norm_curr) + len(norm_saved))
    return round(dice_score, 4)

def _generate_profile_id(filename: str, headers: list[str]) -> str:
    base = Path(filename).stem if filename else "mapping"
    clean_base = re.sub(r'[^a-zA-Z0-9]+', '_', base).strip('_').lower()
    headers_sig = hashlib.md5("||".join(sorted([normalize_header(h) for h in headers])).encode('utf-8')).hexdigest()[:8]
    return f"{clean_base}_{headers_sig}"

def save_mapping_profile(
    excel_filename: str,
    headers: list[str],
    template_mapping: dict | None = None,
    transmittal_mapping: dict | None = None,
    transmittal_columns: list[dict] | None = None,
    profile_id: str | None = None,
    profile_name: str | None = None
) -> dict:
    """
    Persists a mapping profile to disk in mappings/<profile_id>.json.
    Merges with any existing profile data so partial updates don't overwrite other configurations.
    """
    Path(MAPPINGS_DIR).mkdir(parents=True, exist_ok=True)

    raw_headers = [h["original"] if isinstance(h, dict) else str(h) for h in headers]

    if not profile_id:
        profile_id = _generate_profile_id(excel_filename, raw_headers)

    filepath = os.path.join(MAPPINGS_DIR, f"{profile_id}.json")
    existing_data = {}
    if os.path.exists(filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                existing_data = json.load(f)
        except Exception:
            existing_data = {}

    final_name = profile_name or existing_data.get("name") or (Path(excel_filename).name if excel_filename else "Unnamed Profile")
    final_tpl_mapping = template_mapping if template_mapping is not None else existing_data.get("template_mapping", {})
    final_trans_mapping = transmittal_mapping if transmittal_mapping is not None else existing_data.get("transmittal_mapping", {})
    final_trans_cols = transmittal_columns if transmittal_columns is not None else existing_data.get("transmittal_columns", [])

    profile = {
        "id": profile_id,
        "name": final_name,
        "source_file": os.path.basename(excel_filename) if excel_filename else existing_data.get("source_file", ""),
        "excel_headers": raw_headers or existing_data.get("excel_headers", []),
        "template_mapping": final_tpl_mapping,
        "transmittal_mapping": final_trans_mapping,
        "transmittal_columns": final_trans_cols,
        "updated_at": datetime.now().isoformat()
    }

    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(profile, f, indent=2, ensure_ascii=False)

    return profile

def list_saved_profiles(current_headers: list[str] | None = None) -> list[dict]:
    """
    Lists all saved mapping profiles from the mappings/ directory.
    If current_headers is provided, calculates similarity score for each profile.
    """
    profiles = []
    if not os.path.exists(MAPPINGS_DIR):
        return profiles

    raw_curr = [h["original"] if isinstance(h, dict) else str(h) for h in (current_headers or [])]

    for filename in os.listdir(MAPPINGS_DIR):
        if not filename.endswith('.json'):
            continue
        filepath = os.path.join(MAPPINGS_DIR, filename)
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
                score = compute_header_similarity(raw_curr, data.get("excel_headers", [])) if raw_curr else 0.0
                data["similarity"] = score
                profiles.append(data)
        except Exception:
            continue

    # Sort primarily by similarity (descending), then by updated_at (descending)
    profiles.sort(key=lambda p: (p.get("similarity", 0.0), p.get("updated_at", "")), reverse=True)
    return profiles

def find_best_matching_profile(current_headers: list[str], min_threshold: float = 0.55) -> tuple[dict | None, float]:
    """
    Finds the saved mapping profile with the highest similarity to the current headers.
    Returns (best_profile, similarity_score) if score >= min_threshold, else (None, 0.0).
    """
    raw_curr = [h["original"] if isinstance(h, dict) else str(h) for h in current_headers]
    if not raw_curr:
        return None, 0.0

    profiles = list_saved_profiles(raw_curr)
    if not profiles:
        return None, 0.0

    best = profiles[0]
    best_score = best.get("similarity", 0.0)
    if best_score >= min_threshold:
        return best, best_score

    return None, best_score

def apply_mapping_to_headers(
    profile: dict,
    current_headers: list[str],
    current_template_tags: list[str],
    current_transmittal_tags: list[str] | None = None
) -> dict:
    """
    Reconstructs template_mapping, transmittal_mapping, and transmittal_columns
    from a remembered profile for the newly provided Excel headers and docx tags.
    """
    raw_curr = [h["original"] if isinstance(h, dict) else str(h) for h in current_headers]
    norm_map = {normalize_header(h): h for h in raw_curr}

    saved_tpl_map = profile.get("template_mapping", {})
    saved_trans_map = profile.get("transmittal_mapping", {})
    saved_trans_cols = profile.get("transmittal_columns", [])

    # 1. Adapt Template Mapping
    reconstructed_tpl = {}
    all_tpl_tags = set(current_template_tags or []) | set(saved_tpl_map.keys())

    for tag in all_tpl_tags:
        if tag in saved_tpl_map:
            val = saved_tpl_map[tag]
            # Handle static text or blank underline
            if val.startswith("STATIC:") or val in ("__BLANK_UNDERLINE__", "UNDERLINE:"):
                reconstructed_tpl[tag] = val
            elif val in raw_curr:
                reconstructed_tpl[tag] = val
            else:
                # Fuzzy match to new headers
                norm_val = normalize_header(val)
                if norm_val in norm_map:
                    reconstructed_tpl[tag] = norm_map[norm_val]
                else:
                    # Check partial
                    matched_h = None
                    for nh, h in norm_map.items():
                        if (len(norm_val) >= 4 and norm_val in nh) or (len(nh) >= 4 and nh in norm_val):
                            matched_h = h
                            break
                    if matched_h:
                        reconstructed_tpl[tag] = matched_h

    # Fallback to auto-match for any unmapped tags in current template
    if current_template_tags:
        unmapped_tags = [t for t in current_template_tags if t not in reconstructed_tpl]
        if unmapped_tags:
            auto_tpl = auto_match_tags(raw_curr, unmapped_tags)
            reconstructed_tpl.update(auto_tpl)

    # 2. Adapt Transmittal Tag Mapping
    reconstructed_trans_map = {}
    all_trans_tags = set(current_transmittal_tags or []) | set(saved_trans_map.keys())

    for tag in all_trans_tags:
        if tag in saved_trans_map:
            val = saved_trans_map[tag]
            if val.startswith("STATIC:") or val in ("__BLANK_UNDERLINE__", "UNDERLINE:"):
                reconstructed_trans_map[tag] = val
            elif val in raw_curr:
                reconstructed_trans_map[tag] = val
            else:
                norm_val = normalize_header(val)
                if norm_val in norm_map:
                    reconstructed_trans_map[tag] = norm_map[norm_val]
        # Fallback
        if current_transmittal_tags and tag in current_transmittal_tags and tag not in reconstructed_trans_map:
            auto_trans = auto_match_tags(raw_curr, [tag])
            if tag in auto_trans:
                reconstructed_trans_map[tag] = auto_trans[tag]

    # 3. Adapt Transmittal Columns
    # Create a lookup of saved columns
    saved_cols_by_norm = {}
    for c in saved_trans_cols:
        h_name = c.get("header", "")
        saved_cols_by_norm[normalize_header(h_name)] = c

    reconstructed_cols = []
    used_headers = set()

    for h in raw_curr:
        nh = normalize_header(h)
        if nh in saved_cols_by_norm:
            saved_c = saved_cols_by_norm[nh]
            reconstructed_cols.append({
                "header": h,
                "enabled": bool(saved_c.get("enabled", True)),
                "order": int(saved_c.get("order", 999))
            })
            used_headers.add(h)

    # For any new columns not in saved transmittal columns, append them with enabled=False
    max_order = max([c["order"] for c in reconstructed_cols], default=0)
    for h in raw_curr:
        if h not in used_headers:
            max_order += 1
            reconstructed_cols.append({
                "header": h,
                "enabled": False,
                "order": max_order
            })

    def safe_order(c):
        try:
            return int(c.get("order", 999))
        except Exception:
            return 999

    reconstructed_cols.sort(key=safe_order)

    # Re-normalize orders sequentially
    for idx, c in enumerate(reconstructed_cols, start=1):
        c["order"] = idx

    return {
        "template_mapping": reconstructed_tpl,
        "transmittal_mapping": reconstructed_trans_map,
        "transmittal_columns": reconstructed_cols
    }

def delete_mapping_profile(profile_id: str) -> bool:
    """Deletes a saved profile JSON file."""
    filepath = os.path.join(MAPPINGS_DIR, f"{profile_id}.json")
    if os.path.exists(filepath):
        try:
            os.remove(filepath)
            return True
        except Exception:
            return False
    return False
