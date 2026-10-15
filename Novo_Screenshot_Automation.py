import os
import time
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk, messagebox, END
from PIL import ImageGrab
import cv2
import easyocr
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import glob
import re

# =========================
#  GLOBAL DATA STORAGE
# =========================
extracted_data = []  # Stores all OCR results from multiple captures
last_export_index = 0

def format_code(code):
    clean = code.replace("-", "")
    if len(clean) < 14:
        return code
    return f"{clean[:4]}-{clean[4:9]}-{clean[9:12]}-{clean[12:14]}"

reader = easyocr.Reader(['en'], gpu=False)


# =========================
#  IMAGE PREPROCESS
# =========================
def preprocess(image):
    warm = image.copy()
    b, g, r = cv2.split(warm)
    r = cv2.add(r, 20)
    b = cv2.subtract(b, 15)
    warm = cv2.merge([b, g, r])

    gray = cv2.cvtColor(warm, cv2.COLOR_BGR2GRAY)
    denoised = cv2.fastNlMeansDenoising(gray, h=15)
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(denoised)
    return enhanced


def is_ds_number(text):
    return "(DS)" in text.upper().replace(" ", "")


def is_rdesc(text):
    t = text.upper().replace(" ", "")
    return t in ["RDESCRIPTION", "#DESCRIPTION", "LDESCRIPTION", "~DESCRIPTION", "DESCRIPTION"]


# =========================
#  ATA TOKENIZATION HELPERS (for CA Code)
# =========================
def _tokenize_rows(rows):
    """
    Split each OCR cell text into tokens, splitting on '|' but preserving hyphens.
    Returns a flat list of tokens in roughly left-to-right order.
    """
    tokens = []
    for r in rows:
        r_sorted = sorted(r, key=lambda b: b["x"])
        for c in r_sorted:
            t = (c.get("text") or "").strip()
            if not t:
                continue
            parts = [p.strip() for p in re.split(r"\s*\|\s*", t) if p.strip()]
            tokens.extend(parts)
    return tokens


def _extract_ata_related_tokens(tokens, ata_chapter="53"):
    """
    From a token list, extract ATA-related tokens IF the capture shows ATA context.
    ATA context if ANY token contains 'ATA' (case-insensitive).

      - Keep tokens with 'ATA' (e.g., 'SA-S17-ATA53')
      - Keep 'SA-...-<chapter>-...' (default chapter 53)
      - Keep equipment-like codes 'E53C73383000' (E + 2 digits + letter + >=8 digits)
    """
    if not tokens:
        return []

    has_ata_context = any("ATA" in t.upper() for t in tokens)
    if not has_ata_context:
        return []

    ata_rel = []
    chap = re.escape(ata_chapter)

    sa_chapter_pat = re.compile(rf"^SA-[A-Z0-9]+-{chap}-[A-Z0-9]+$", re.IGNORECASE)
    equip_pat = re.compile(r"^[Ee]\d{2}[A-Za-z]\d{8,}$")

    for t in tokens:
        if not t:
            continue

        if "ATA" in t.upper() or sa_chapter_pat.match(t) or equip_pat.match(t):
            ata_rel.append(t)

    # De-duplicate while preserving order (case-insensitive key)
    seen = set()
    out = []
    for t in ata_rel:
        k = t.upper()
        if k in seen:
            continue
        seen.add(k)
        out.append(t)

    return out


def extract_ata_context_from_text(text, ata_chapter="53"):
    """
    Extract ATA-related tokens *only if* the text has 'ATA' context.
    Returns a list of tokens in display order, deduped (case-insensitive).
    (Used as fallback if precomputed tokens are not available.)
    """
    if not text or "ATA" not in text.upper():
        return []

    raw_segments = [seg.strip(" \t\n\r:;,") for seg in re.split(r"\|", text) if seg.strip()]
    candidates = []
    for seg in raw_segments:
        parts = [p.strip(" \t\n\r:;,") for p in seg.split() if p.strip()]
        if seg not in parts:
            parts.append(seg)
        candidates.extend(parts)

    seen = set()
    out = []
    chap = re.escape(ata_chapter)
    sa_chapter_pat = re.compile(rf"(?i)^SA-[A-Z0-9]+-{chap}-[A-Z0-9]+$")
    equip_pat      = re.compile(r"(?i)^E\d{{2}}[A-Z]\d{{8,}}$")  # e.g., E53C73383000

    for t in candidates:
        tu = t.upper()
        if ("ATA" in tu) or sa_chapter_pat.match(t) or equip_pat.match(t):
            if tu not in seen:
                seen.add(tu)
                out.append(t)

    return out


# =========================
#  OCR + EXTRACT
# =========================
def ocr_and_extract(img_path, suppress_headers=False):
    """
    Runs OCR pipeline and returns formatted text for storage.
    If suppress_headers=True, do not include dividers or 'Data Extracted from:' line.
    """
    image = cv2.imread(img_path)
    if image is None:
        return None

    processed = preprocess(image)
    results = reader.readtext(processed)

    items = []
    for bbox, text, conf in results:
        x_coords = [p[0] for p in bbox]
        y_coords = [p[1] for p in bbox]
        items.append({
            "text": text,
            "x": int(min(x_coords)),
            "y": int(min(y_coords)),
            "bbox": bbox
        })

    if not items:
        return None

    # Group into rows
    rows = []
    current_row = []
    threshold = 18
    items_sorted = sorted(items, key=lambda b: b["y"])
    for b in items_sorted:
        if not current_row:
            current_row.append(b)
            continue
        if abs(b["y"] - current_row[-1]["y"]) < threshold:
            current_row.append(b)
        else:
            rows.append(current_row)
            current_row = [b]
    if current_row:
        rows.append(current_row)

    lines = []
    if not suppress_headers:
        lines.append("=" * 50)
        lines.append(f"Data Extracted from: {os.path.basename(img_path)}")
        lines.append("=" * 50)
        lines.append("")

    for row_idx, row in enumerate(rows, start=1):
        row_sorted = sorted(row, key=lambda b: b["x"])
        part_name_printed = False

        ds_index = None
        r_description_index = None
        for i, cell in enumerate(row_sorted):
            if is_ds_number(cell["text"]):
                ds_index = i
            if is_rdesc(cell["text"]):
                r_description_index = i

        last_index = len(row_sorted) - 1

        for col_idx, cell in enumerate(row_sorted, start=1):
            cell_index = col_idx - 1
            text = cell['text']
            label = "Normal"
            should_print = False

            if ds_index is not None:
                items_after_ds = len(row_sorted) - (ds_index + 1)

                if cell_index == ds_index:
                    label = "DS Number"
                    text_fixed = fix_leading_alpha(text)
                    text_core = text_fixed.replace("(DS)", "").strip()
                    text = format_code(text_core) + " (DS)"
                    should_print = True
                elif items_after_ds > 1:
                    if cell_index == ds_index + 1:
                        label = "Higher Assembly Number"
                        text = format_code(fix_leading_alpha(text))
                        should_print = True
                    elif cell_index == last_index:
                        label = "Part Number"
                        text = format_code(fix_leading_alpha(text))
                        should_print = True
                elif items_after_ds == 1:
                    if cell_index == ds_index + 1:
                        label = "Part Number"
                        text = format_code(fix_leading_alpha(text))
                        should_print = True

            if (r_description_index is not None and
                    cell_index == r_description_index + 1 and
                    not part_name_printed):
                label = "Part Name"
                should_print = True
                part_name_printed = True

            if should_print:
                lines.append(f"{label}: {text}")

    # === Precompute ATA Context tokens (for CA Code), but DO NOT display in UI ===
    all_tokens = _tokenize_rows(rows)
    ata_related = _extract_ata_related_tokens(all_tokens, ata_chapter="53")
    if ata_related:
        # Store in text so Summary/Export can derive CA reliably; we'll hide it in display.
        lines.append("ATA Context: " + " | ".join(ata_related))

    # Always end with a blank line so entries are separated by one empty line
    lines.append("")
    return "\n".join(lines)


def fix_leading_alpha(raw: str) -> str:
    """
    Fix only the very first character:
    - If the first non-space char is '5', treat it as 'S' (OCR confusion).
    - If it's an alphabet, uppercase it.
    - Do not touch the rest of the string (keep digits, dashes, spaces, '(DS)' etc. as-is).
    """
    if not raw:
        return raw

    s = str(raw)
    i = 0
    while i < len(s) and s[i].isspace():
        i += 1
    if i >= len(s):
        return s

    first = s[i]
    if first == '5':
        corrected = 'S'
    elif first == '€':
        corrected = 'E'
    elif first == '0':
        corrected = 'O'
    elif first == 'l' or first == '1':
        corrected = 'L'
    elif first == '8':
        corrected = 'B'
    elif first == '6':
        corrected = 'G'
    elif first.isalpha():
        corrected = first.upper()
    else:
        corrected = first

    if corrected == first:
        return s
    return s[:i] + corrected + s[i+1:]


def extract_data_fields(capture_data):
    """
    Extract Part Number, DS Number, and Higher Assembly Number from captured data.
    Returns a dictionary with these fields.
    """
    fields = {
        "Part Number": "",
        "DS Number": "",
        "Higher Assembly Number": ""
    }

    lines = capture_data.split('\n')
    for line in lines:
        line = line.strip()
        if not line or ':' not in line:
            continue

        label, value = line.split(':', 1)
        label = label.strip()
        value = value.strip()

        if label in fields:
            fields[label] = value

    return fields


def find_latest_excel_file():
    """
    Find the most recent Novo_Screenshot_Data_*.xlsx file in current directory.
    Returns the file path if found, None otherwise.
    """
    pattern = "Novo_Screenshot_Data_*.xlsx"
    files = glob.glob(pattern)
    if not files:
        return None
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

def _remove_default_empty_sheets(wb, keep_names):
    """
    Remove default blank sheets like 'Sheet', 'Sheet1', etc. if they are empty.
    keep_names: set of sheet names that must never be removed.
    """
    def is_empty_sheet(ws):
        # If there is any cell with a value, it's not empty
        for row in ws.iter_rows(values_only=True):
            if any(cell is not None and str(cell).strip() != "" for cell in row):
                return False
        return True

    # iterate over a copy since we may remove sheets
    for name in list(wb.sheetnames):
        if name in keep_names:
            continue
        # remove only default-like sheet names
        if name.lower().startswith("sheet"):
            ws = wb[name]
            if is_empty_sheet(ws):
                wb.remove(ws)

def export_to_excel():
    """
    Create or append to Novo_Screenshot_Data_{timestamp}.xlsx.
    Only appends newly captured data since last export.

    Sheet 1: OCR Data
      Row 1: [Part Number,  , DS Number, CA Code, Captured At]
      Row 2: [Higher Assembly Number or "",  , "", "", ""]

    Sheet 2: Parts_DS_CA_Summary
      One row per capture:
      [Part Number, DS Number, Higher Assembly Number (blank if not found), CA Code, Captured At]

    Also:
      - Forces headers to Row 1 (prevents headers starting at Row 2)
      - Prevents duplicate "Part Number" header
      - Removes default empty sheet(s) like "Sheet", "Sheet1" if present and empty
    """
    global last_export_index

    # --- Safety: if pointer is stale (e.g., list was cleared), reset it ---
    if last_export_index > len(extracted_data):
        last_export_index = 0

    if not extracted_data or last_export_index >= len(extracted_data):
        messagebox.showinfo("No Data", "No new OCR data to export. Capture more data first.")
        return

    new_captures = extracted_data[last_export_index:]

    OCR_SHEET_NAME = "OCR Data"
    SUMMARY_SHEET_NAME = "Parts_DS_CA_Summary"

    PN_PATTERN = re.compile(r"\b[DE]\d{3,}-\d{5}-\d{3}-\d{2}\b", re.IGNORECASE)

    # --------------------------
    # Helpers
    # --------------------------
    def _ensure_headers_row1(sheet, headers, widths=None):
        """Force headers into row 1 (do not use append)."""
        for col, h in enumerate(headers, start=1):
            sheet.cell(row=1, column=col, value=h)

        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=12)

        for col in range(1, len(headers) + 1):
            cell = sheet.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        if widths:
            for idx, w in enumerate(widths, start=1):
                col_letter = chr(64 + idx)  # A,B,C...
                sheet.column_dimensions[col_letter].width = w

    def _fix_duplicate_part_number_header(sheet):
        """
        If an older export created duplicate 'Part Number' headers in row 1,
        clear duplicates beyond column 1.
        """
        for col in range(2, sheet.max_column + 1):
            v = sheet.cell(row=1, column=col).value
            if v == "Part Number":
                sheet.cell(row=1, column=col, value="")

    def _remove_default_empty_sheets(wb, keep_names):
        """
        Remove default blank sheets like 'Sheet', 'Sheet1', etc. if they are empty.
        keep_names: set of sheet names that must never be removed.
        """
        def is_empty_sheet(ws):
            for row in ws.iter_rows(values_only=True):
                if any(cell is not None and str(cell).strip() != "" for cell in row):
                    return False
            return True

        for name in list(wb.sheetnames):
            if name in keep_names:
                continue
            if name.lower().startswith("sheet"):
                ws_tmp = wb[name]
                if is_empty_sheet(ws_tmp):
                    wb.remove(ws_tmp)

    def _pick_ca_code_from_ata_tokens(ata_tokens, default_chapter="53"):
        if not ata_tokens:
            return None

        tokens = [t.strip(" \t,:;|").replace("\u200b", "") for t in ata_tokens if t and t.strip()]
        # de-dup case-insensitive preserving order
        seen = set()
        dedup = []
        for t in tokens:
            k = t.upper()
            if k not in seen:
                seen.add(k)
                dedup.append(t)
        tokens = dedup

        contains_ata_code = re.compile(r"(?i)^(?!ATA$)[A-Z0-9\-]*ATA[A-Z0-9\-]*$")
        sa_chapter_pat = re.compile(r"(?i)^SA-[A-Z0-9]+-(\d{2})-[A-Z0-9]+$")
        ca_pat = re.compile(r"(?i)^[Ee]{1,2}\d{2}[A-Za-z]?\d{6,}[A-Za-z0-9]*$")
        ata_chapter_in_token = re.compile(r"(?i)ATA\s*([0-9]{2})")

        anchor_idx = None
        chapter = None

        # Prefer anchor token containing ATA and not bare ATA
        for i, tok in enumerate(tokens):
            if contains_ata_code.match(tok):
                anchor_idx = i
                m = ata_chapter_in_token.search(tok)
                if m:
                    chapter = m.group(1)
                if not chapter:
                    m2 = sa_chapter_pat.match(tok)
                    if m2:
                        chapter = m2.group(1)
                break

        # Fallback: structural SA-...-XX-...
        if anchor_idx is None:
            for i, tok in enumerate(tokens):
                m2 = sa_chapter_pat.match(tok)
                if m2:
                    anchor_idx = i
                    chapter = m2.group(1)
                    break

        # Last resort: bare ATA
        if anchor_idx is None:
            for i, tok in enumerate(tokens):
                if tok.upper() == "ATA":
                    anchor_idx = i
                    chapter = chapter or default_chapter
                    break

        # If still no anchor, pick first CA-looking token
        if anchor_idx is None:
            for tok in tokens:
                norm = tok.replace(" ", "")
                if ca_pat.match(norm):
                    return norm
            return None

        # Look ahead for CA candidates
        LOOKAHEAD = 4
        after = tokens[anchor_idx + 1:]
        window = after[:LOOKAHEAD] + after[LOOKAHEAD:]

        candidates = []
        for tok in window:
            norm = tok.replace(" ", "")
            if ca_pat.match(norm):
                candidates.append(norm)

        if not candidates:
            return None

        # Prefer candidate matching chapter
        if chapter:
            for c in candidates:
                m = re.match(r"(?i)^E{1,2}(\d{2})", c)
                if m and m.group(1) == chapter:
                    return c

        return candidates[0]

    def _compute_ca_code_from_capture(capture_data):
        # Prefer precomputed "ATA Context:" tokens if present
        ata_tokens = []
        for s in capture_data.splitlines():
            if s.lower().startswith("ata context:"):
                payload = s.split(":", 1)[1].strip()
                ata_tokens = [t.strip() for t in payload.split("|") if t.strip()]
                ata_tokens = [t for t in ata_tokens if t.upper() != "ATA"]
                break

        # Fallback derive tokens from raw text
        if not ata_tokens:
            raw_segments = [seg.strip(" \t\n\r:;,") for seg in re.split(r"\|", capture_data) if seg.strip()]
            candidates = []
            for seg in raw_segments:
                parts = [p.strip(" \t\n\r:;,") for p in seg.split() if p.strip()]
                if seg not in parts:
                    parts.append(seg)
                candidates.extend(parts)

            chap_pat = re.compile(r"(?i)^SA-[A-Z0-9]+-\d{2}-[A-Z0-9]+$")
            equip_pat = re.compile(r"(?i)^E\d{2}[A-Z]\d{6,}[A-Z0-9]*$")
            for t in candidates:
                tu = t.upper()
                if ("ATA" in tu) or chap_pat.match(t) or equip_pat.match(t):
                    if tu != "ATA":
                        ata_tokens.append(t)

        return _pick_ca_code_from_ata_tokens(ata_tokens) or ""

    # --------------------------
    # Main logic
    # --------------------------
    try:
        existing_file = find_latest_excel_file()

        # Load or create workbook
        if existing_file:
            wb = load_workbook(existing_file)
        else:
            wb = Workbook()

        # Ensure OCR Data sheet exists (do NOT use wb.active)
        if OCR_SHEET_NAME in wb.sheetnames:
            ws = wb[OCR_SHEET_NAME]
        else:
            ws = wb.create_sheet(OCR_SHEET_NAME, 0)

        # Ensure Summary sheet exists
        if SUMMARY_SHEET_NAME in wb.sheetnames:
            ws_summary = wb[SUMMARY_SHEET_NAME]
        else:
            ws_summary = wb.create_sheet(SUMMARY_SHEET_NAME)

        # Remove default empty sheet(s) like "Sheet" if present
        _remove_default_empty_sheets(wb, keep_names={OCR_SHEET_NAME, SUMMARY_SHEET_NAME})

        # Force headers row 1 for OCR Data (and fix duplicate header)
        _ensure_headers_row1(
            ws,
            ["Part Number", "", "DS Number", "CA Code", "Captured At"],
            widths=[25, 20, 25, 25, 22]
        )
        _fix_duplicate_part_number_header(ws)

        # Force headers row 1 for Summary
        _ensure_headers_row1(
            ws_summary,
            ["Part Number", "DS Number", "Higher Assembly Number", "CA Code", "Captured At"],
            widths=[25, 25, 28, 25, 22]
        )

        valid_rows = 0
        summary_rows_added = 0

        # Export new captures
        for capture_data in new_captures:
            fields = extract_data_fields(capture_data)
            pn_raw = fields.get("Part Number", "")
            ds_raw = fields.get("DS Number", "")
            ha_raw = fields.get("Higher Assembly Number", "")

            n = fix_leading_alpha(pn_raw).replace("-", "") if pn_raw else ""
            ds = fix_leading_alpha(ds_raw).replace("-", "").replace("(DS)", "").strip() if ds_raw else ""
            ha = fix_leading_alpha(ha_raw).replace("-", "") if ha_raw else ""

            ca = _compute_ca_code_from_capture(capture_data)

            # If all empty, try PN regex fallback
            if not any([n, ds, ha, ca]):
                m = PN_PATTERN.search(capture_data)
                if m:
                    n = m.group(0).replace("-", "")
                else:
                    continue

            current_time = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")

            # Sheet 1: OCR Data (2-row export)
            ws.append([n, "", ds, ca, current_time])
            ws.append([ha if ha else "", "", "", "", ""])  # HA blank if not found
            valid_rows += 2

            # Sheet 2: Summary (1-row export)
            ws_summary.append([n, ds, ha if ha else "", ca, current_time])
            summary_rows_added += 1

        # Save
        if existing_file:
            save_name = existing_file
        else:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            save_name = f"Novo_Screenshot_Data_{timestamp}.xlsx"

        wb.save(save_name)

        try:
            os.startfile(save_name)
        except Exception as e:
            messagebox.showwarning("Open Excel", f"File saved but could not be opened automatically:\n{e}")

        messagebox.showinfo(
            "Export Successful",
            f"Data exported to:\n{save_name}\n\n"
            f"{valid_rows} rows exported to '{OCR_SHEET_NAME}'\n"
            f"{summary_rows_added} rows exported to '{SUMMARY_SHEET_NAME}'\n"
            f"(from {len(new_captures)} new captures)"
        )

        last_export_index = len(extracted_data)

    except Exception as e:
        messagebox.showerror("Export Error", f"Failed to export data:\n{str(e)}")


# =========================
#  GUI FUNCTIONS
# =========================
def show_info():
    info = tk.Toplevel(root)
    info.title("User Guide")
    info.geometry("450x350")

    text = tk.Text(info, wrap="word")
    text.pack(fill="both", expand=True, padx=10, pady=10)

    guide = """
    USER GUIDE

    • Click 'Screenshot' multiple times to capture & store data
    • Click 'Select Images' multiple times to process multiple image files
    • All extracted data is stored automatically
    • Click 'Summary' to view ALL stored data in new window
    • In Summary window: Click 'Export Excel' (green) to save/append data
    • ℹ️ icon shows this guide anytime.
    """

    text.insert("1.0", guide)
    text.config(state="disabled")


def clear_all_data():
    """Clear all data from main window"""
    global extracted_data
    extracted_data = []
    status_label.config(text="Ready | 0 captures stored")


def clear_both(left_text):
    """Clear text area in result window"""
    global extracted_data
    extracted_data = []
    left_text.delete("1.0", END)
    left_text.insert("1.0", "All data cleared.")
    if 'status_label' in globals():
        status_label.config(text="Ready | 0 captures stored")


def open_result_window():
    if not extracted_data:
        messagebox.showinfo("No Data", "No OCR data captured yet. Take screenshots or select images first.")
        return

    win = tk.Toplevel(root)
    win.title("Data Summary - All Screenshots")
    win.geometry("700x550")
    win.minsize(600, 400)

    header_frame = tk.Frame(win)
    header_frame.pack(fill="x", pady=(10, 5), padx=15)

    title_label = tk.Label(header_frame, text=f"Data Summary - {get_capture_count()} captures",
                           font=("Arial", 14, "bold"), fg="#2d3748")
    title_label.pack(side="left")

    credit_label = tk.Label(header_frame, text="Built by ME_Digi Team", font=("Arial", 9), fg="#3452eb")
    credit_label.pack(side="right")

    main_frame = tk.Frame(win)
    main_frame.pack(fill="both", expand=True, padx=15, pady=10)

    # LEFT TEXT AREA (ONLY)
    left_label = tk.Label(main_frame, text="All Extracted Data", font=("Arial", 10, "bold"))
    left_label.pack(anchor="w")

    left_text = scrolledtext.ScrolledText(main_frame, wrap="word", font=("Consolas", 10), bg="#f8f9fa", fg="#2d3748",
                                          selectbackground="#4CAF50", selectforeground="white", state="normal",
                                          cursor="hand2")
    left_text.pack(fill="both", expand=True)

    # --- helper to pick CA code robustly (local to display) ---
    def pick_ca_code_from_ata_tokens(ata_tokens, default_chapter="53"):
        """
        Robust CA Code picker:
          - Prefer anchor = first *code-like* token containing 'ATA' (e.g., 'SA-S17-ATA53'), not bare 'ATA'.
          - Fallback anchor = first structural chapter token 'SA-...-<chapter>-...'.
          - Last resort anchor = bare 'ATA'.
          - From the anchor, search a short lookahead window for the first valid CA code (E...),
            prioritizing those whose chapter matches the anchor's chapter.

        Returns normalized CA code (spaces removed) or None.
        """
        if not ata_tokens:
            return None

        tokens = [t.strip(" \t,:;|").replace("\u200b", "") for t in ata_tokens if t and t.strip()]
        # Remove exact duplicates while preserving order
        seen = set()
        dedup_tokens = []
        for t in tokens:
            k = t.upper()
            if k not in seen:
                seen.add(k)
                dedup_tokens.append(t)
        tokens = dedup_tokens

        contains_ata_code = re.compile(r"(?i)^(?!ATA$)[A-Z0-9\-]*ATA[A-Z0-9\-]*$")
        sa_chapter_pat    = re.compile(r"(?i)^SA-[A-Z0-9]+-(\d{2})-[A-Z0-9]+$")
        ca_pat            = re.compile(r"(?i)^[Ee]{1,2}\d{2}[A-Za-z]?\d{6,}[A-Za-z0-9]*$")
        ata_chapter_in_token = re.compile(r"(?i)ATA\s*([0-9]{2})")

        anchor_idx = None
        chapter = None

        for i, tok in enumerate(tokens):
            if contains_ata_code.match(tok):
                anchor_idx = i
                m = ata_chapter_in_token.search(tok)
                if m:
                    chapter = m.group(1)
                if not chapter:
                    m2 = sa_chapter_pat.match(tok)
                    if m2:
                        chapter = m2.group(1)
                break

        if anchor_idx is None:
            for i, tok in enumerate(tokens):
                m2 = sa_chapter_pat.match(tok)
                if m2:
                    anchor_idx = i
                    chapter = m2.group(1)
                    break

        if anchor_idx is None:
            for i, tok in enumerate(tokens):
                if tok.upper() == "ATA":
                    anchor_idx = i
                    if not chapter:
                        chapter = default_chapter
                    break

        if anchor_idx is None:
            for tok in tokens:
                norm = tok.replace(" ", "")
                if ca_pat.match(norm):
                    return norm
            return None

        LOOKAHEAD = 4
        after = tokens[anchor_idx + 1:]
        window = after[:LOOKAHEAD] + after[LOOKAHEAD:]

        candidates = []
        for tok in window:
            norm = tok.replace(" ", "")
            if ca_pat.match(norm):
                candidates.append(norm)

        if not candidates:
            return None

        if chapter:
            for c in candidates:
                m = re.match(r"(?i)^E{1,2}(\d{2})", c)
                if m and m.group(1) == chapter:
                    return c

        return candidates[0]

    # BUILD CONTENT
    display_content = []
    for capture_idx, capture_data in enumerate(extracted_data, 1):
        # Special heading line for multi-select sets
        if capture_data.strip().startswith("Affected Parts"):
            display_content.append("\tAffected Parts")
            display_content.append("")  # one blank line after heading
            continue

        lines = capture_data.split('\n')
        has_header = any('Data Extracted from:' in line for line in lines)

        # For single/screenshot mode only: show the filename header
        if has_header:
            filename_line = next((line for line in lines if 'Data Extracted from:' in line), "")
            if filename_line:
                display_content.append(filename_line.strip())
                display_content.append("")

        # Collect fields
        part_names = []
        part_number = None
        ds_number = None

        for line in lines:
            s = line.strip()
            if not s or s.startswith('=') or ':' not in s:
                continue
            label, value = s.split(':', 1)
            label = label.strip()
            value = value.strip()

            if label == "Part Name":
                part_names.append(value)
            elif label == "Part Number":
                part_number = value
            elif label == "DS Number":
                ds_number = value

        # For multi-upload entries (no header): " PN: PartName"
        if not has_header and part_names:
            for p in part_names:
                if part_number:
                    display_content.append(f" {part_number}: {p}")
                elif ds_number:
                    display_content.append(f" {ds_number}: {p}")

        # In single/screenshot entries, keep "Affected Parts" block
        if has_header and part_names:
            display_content.append(" Affected Parts:")
            display_content.append("")
            for p in part_names:
                if part_number:
                    display_content.append(f" {part_number}: {p}")
                elif ds_number:
                    display_content.append(f" {p}: {ds_number}")

        # Collect and print other label:value lines with enforced order (HA above DS)
        ha_line = None
        ds_line = None

        for line in lines:
            s = line.strip()
            if (s and ':' in s and
                    'Data Extracted from:' not in s and
                    not s.startswith('Part Name')):

                # Do not show ATA Context in the UI
                if s.lower().startswith('ata context:'):
                    continue

                # Remove "Part Number: ..." for BOTH screenshots and uploads
                if s.startswith('Part Number:'):
                    continue

                # Capture and normalize "Higher Assembly Number" → append ": ASSY" neatly
                if s.lower().startswith('higher assembly number:'):
                    val = s[len('Higher Assembly Number:'):].strip()
                    val_clean = val.strip(' "\'')

                    if not re.search(r'\bASSY\b', val_clean, flags=re.IGNORECASE):
                        val_clean = f"{val_clean}: ASSY"
                    elif not re.search(r':\s*ASSY\s*$', val_clean, flags=re.IGNORECASE):
                        val_clean = re.sub(r'ASSY["\']?$', 'ASSY', val_clean, flags=re.IGNORECASE)

                    ha_line = f"Higher Assembly Number: {val_clean}"
                    continue  # don't append now; we'll enforce order later

                # Capture "DS Number" to print after HA
                if s.lower().startswith('ds number:'):
                    ds_line = s
                    continue  # don't append now; we'll enforce order later

                # Any other lines print as-is here (except ATA Context which we already skip)
                display_content.append(f" {s}")

        # ✅ Enforce the desired order at the end: HA first, then DS
        if ha_line:
            display_content.append(f" {ha_line}")
        if ds_line:
            display_content.append(f" {ds_line}")

        # === Display CA Code (derived from precomputed ATA Context line or via fallback) ===
        ata_tokens = []
        for s in lines:
            if s.lower().startswith('ata context:'):
                payload = s.split(':', 1)[1].strip()
                ata_tokens = [t.strip() for t in payload.split('|') if t.strip()]
                ata_tokens = [t for t in ata_tokens if t.upper() != 'ATA']  # avoid bare 'ATA'
                break
        # Fallback derive if not present
        if not ata_tokens:
            try:
                ata_tokens = extract_ata_context_from_text(capture_data, ata_chapter="53")
                ata_tokens = [t for t in ata_tokens if t.upper() != 'ATA']
            except Exception:
                ata_tokens = []

        ca_code = pick_ca_code_from_ata_tokens(ata_tokens, default_chapter="53")
        if ca_code:
            display_content.append(f" CA Code: {ca_code}")

        # - Single/screenshot entries keep the old dashed divider
        # - Batch entries get just one blank line between items
        display_content.append("---------------------------------" if has_header else "")

    full_content = "\n".join(display_content)
    left_text.insert("1.0", full_content)
    left_text.focus_set()

    # BUTTON ROW - Export Excel (LEFT) + Clear All Data (RIGHT)
    button_frame = tk.Frame(win)
    button_frame.pack(fill="x", pady=10, padx=15)

    export_btn = tk.Button(button_frame, text="Export Excel", command=export_to_excel, bg="#4CAF50",
                           fg="white", font=("Arial", 10, "bold"), relief="raised", bd=2, cursor="hand2")
    export_btn.pack(side="right", padx=(0, 10))

    clear_btn = tk.Button(button_frame, text="Clear All Data", command=lambda: clear_both(left_text),
                          bg="#ff6b6b", fg="white", font=("Arial", 10, "bold"), relief="raised", bd=2, cursor="hand2")
    clear_btn.pack(side="left", padx=(0, 10))


# Markers you use as headings inside extracted_data
BATCH_HEADING_MARKERS = ("Affected Parts", "Data of current Set")

def is_batch_heading(entry) -> bool:
    """Returns True if the entry is just a heading marker we inserted for display."""
    return isinstance(entry, str) and entry.strip().startswith(BATCH_HEADING_MARKERS)

def get_capture_count() -> int:
    """Counts only real OCR captures (excludes heading markers)."""
    return sum(1 for e in extracted_data if not is_batch_heading(e))


def process_image(img_path, suppress_headers=False):
    progress_bar["value"] = 0
    status_label.config(text="Processing...")
    root.update_idletasks()

    def worker():
        try:
            result = ocr_and_extract(img_path, suppress_headers=suppress_headers)
            if result:
                extracted_data.append(result)
                status_label.config(text=f"Stored! Total: {get_capture_count()} captures")
            else:
                status_label.config(text="No text detected")

            for i in range(1, 101):
                progress_bar["value"] = i
                root.update_idletasks()
                time.sleep(0.005)

            progress_bar["value"] = 0
            status_label.config(text=f"Ready | {get_capture_count()} captures stored")

        except Exception as e:
            status_label.config(text="Error occurred")
            messagebox.showerror("OCR Error", str(e))
        finally:
            progress_bar["value"] = 0

    t = threading.Thread(target=worker, daemon=True)
    t.start()


def get_next_screenshot_name():
    """
    Returns a filename like screenshot1.png, screenshot2.png, ...
    based on the highest existing screenshotN.png in the current folder.
    """
    pattern = re.compile(r"^screenshot(\d+)\.png$", re.IGNORECASE)
    max_num = 0

    for fname in os.listdir("."):
        match = pattern.match(fname)
        if match:
            num = int(match.group(1))
            if num > max_num:
                max_num = num

    next_num = max_num + 1
    return f"screenshot{next_num}.png"


def take_screenshot():
    root.iconify()
    root.update()
    time.sleep(0.45)
    img = ImageGrab.grab()

    # auto-incrementing screenshot name
    path = get_next_screenshot_name()
    img.save(path)

    root.deiconify()
    process_image(path)


# === Process selected images strictly in the order selected ===
def process_images_in_selection_order(paths, suppress_headers=False):
    """
    Processes a list of image paths sequentially (one background thread),
    ensuring extracted_data appends in the same order as 'paths'.
    """
    status_label.config(text=f"Processing {len(paths)} images in selected order...")
    root.update_idletasks()

    def worker():
        try:
            total = len(paths)
            for idx, path in enumerate(paths, 1):
                status_label.config(text=f"Processing {idx}/{total}: {os.path.basename(path)}")
                root.update_idletasks()

                result = ocr_and_extract(path, suppress_headers=suppress_headers)
                if result:
                    extracted_data.append(result)

                for i in range(1, 101):
                    progress_bar["value"] = i
                    root.update_idletasks()
                    time.sleep(0.005)

                progress_bar["value"] = 0

            status_label.config(text=f"Ready | {get_capture_count()} captures stored")
        except Exception as e:
            messagebox.showerror("Batch OCR Error", str(e))
            status_label.config(text="Error occurred")
        finally:
            progress_bar["value"] = 0

    threading.Thread(target=worker, daemon=True).start()


def select_images():
    paths = filedialog.askopenfilenames(
        title="Select Image Files",
        filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff")]
    )
    if paths:
        paths = list(paths)  # preserve dialog-provided order
        status_label.config(text=f"Queued {len(paths)} images...")
        root.update_idletasks()

        batch_mode = len(paths) > 1

        # For multi-select only: add heading once at the start of this set
        if batch_mode:
            extracted_data.append("Affected Parts\n")
            process_images_in_selection_order(paths, suppress_headers=True)
        else:
            process_image(paths[0], suppress_headers=False)


def run_summary():
    progress_bar["value"] = 0
    for i in range(1, 101):
        progress_bar["value"] = i
        root.update_idletasks()
        time.sleep(0.01)
    open_result_window()


# =========================
#  MAIN GUI (unchanged)
# =========================
root = tk.Tk()
root.title("NOVO Screenshot Automation Tool")
root.geometry("450x110")
root.resizable(False, False)

status_row = tk.Frame(root)
status_row.pack(fill="x", pady=(5, 0), padx=10)

status_label = tk.Label(status_row, text="Ready | 0 captures stored", font=("Arial", 9), fg="#666")
status_label.pack(side="left")

info_button = tk.Button(status_row, text="ℹ️", font=("Arial", 8, "bold"), command=show_info, bg="#e6e6e6",
                        relief="raised", bd=2, padx=2, pady=0, width=2)
info_button.pack(side="right")

button_row = tk.Frame(root)
button_row.pack(fill="x", pady=5)

center_frame = tk.Frame(button_row)
center_frame.pack(expand=True)

ss_btn = ttk.Button(center_frame, text="Screenshot", command=take_screenshot)
ss_btn.pack(side="left", padx=2)

select_btn = ttk.Button(center_frame, text="Select Images", command=select_images)
select_btn.pack(side="left", padx=2)

summary_btn = ttk.Button(center_frame, text="Summary", command=run_summary)
summary_btn.pack(side="left", padx=2)

clear_main_btn = tk.Button(center_frame, text="Clear", command=clear_all_data, bg="#ff6b6b", fg="white",
                           font=("Arial", 8, "bold"), relief="raised", bd=2, cursor="hand2")
clear_main_btn.pack(side="left", padx=2)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=320, mode="determinate")
progress_bar.pack(pady=8)

root.mainloop()