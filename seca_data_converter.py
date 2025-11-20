"""SECA Data Converter
=======================

This script opens GUI dialogs to let a user select one or more SECA PDF
reports and a destination folder.  It then extracts the patient metadata and
measurement values defined in the project requirements and stores them in an
Excel workbook (one row per PDF).

Usage::

    python seca_data_converter.py

Dependencies:
    - pdfplumber
    - pytesseract (requires the Tesseract OCR binary)
    - pandas (which also requires openpyxl for Excel output)
    - tkinter (bundled with most Python distributions)
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import pdfplumber
import pytesseract
from tkinter import Tk, messagebox
from tkinter import filedialog

# --- OCR region configuration ---

# Base page size in pixels that coordinates were measured on
BASE_PAGE_WIDTH = 652
BASE_PAGE_HEIGHT = 922

# Region in that base coordinate system (left, top, right, bottom)
RAW_OCR_BOX = (440, 239, 568, 875)

def get_scaled_ocr_box(image_size):
    """Scale RAW_OCR_BOX from the 652x922 coordinate system to the actual rendered size."""
    img_w, img_h = image_size
    sx = img_w / BASE_PAGE_WIDTH
    sy = img_h / BASE_PAGE_HEIGHT
    x0, y0, x1, y1 = RAW_OCR_BOX
    return (
        int(x0 * sx),
        int(y0 * sy),
        int(x1 * sx),
        int(y1 * sy),
    )

# Tell pytesseract where tesseract.exe lives
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Regex that captures floats/integers including optional comma decimal marks.
NUMBER_PATTERN = re.compile(r"-?\d+(?:[.,]\d+)?")

PATIENT_FIELDS = {
    "Patient ID": re.compile(r"ID[:\s]+([A-Za-z0-9-]+)", re.IGNORECASE),
    # Use a word boundary to avoid matching the "age" portion inside other words
    # such as "Average", which previously resulted in incorrect ages (e.g. "1").
    "Age": re.compile(r"\bAge[:\s]+(\d+)", re.IGNORECASE),
}

def normalize_number(token: str) -> float:
    token = token.replace(",", ".")
    return float(token)


def collapse_whitespace(text: str) -> str:
    """Return *text* with all whitespace collapsed to single spaces."""

    return " ".join(text.split())


def extract_page_text(page: "pdfplumber.page.Page", include_ocr: bool = True) -> str:
    """Return textual content for a page, optionally augmented with OCR."""

    parts: List[str] = []
    text = page.extract_text() or ""
    if text.strip():
        parts.append(text)

    if include_ocr:
        ocr_text = ""
        try:
            # Render full page, then crop to the scaled region
            pil_image = page.to_image(resolution=300).original
            crop_box = get_scaled_ocr_box(pil_image.size)
            cropped = pil_image.crop(crop_box)

            # DEBUG (optional): save the cropped region once to visually verify
            # cropped.save("debug_cropped_page.png")

            ocr_text = pytesseract.image_to_string(cropped)
        except Exception:
            ocr_text = ""
        if ocr_text.strip():
            parts.append(ocr_text)

    return "\n".join(parts)

def extract_pdf_text(pdf_path: Path) -> str:
    """Extract ONLY OCR text from the cropped measurement region."""
    parts: List[str] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            try:
                pil_image = page.to_image(resolution=300).original
                crop_box = get_scaled_ocr_box(pil_image.size)
                cropped = pil_image.crop(crop_box)
                ocr_text = pytesseract.image_to_string(cropped)
            except Exception:
                ocr_text = ""
            parts.append(ocr_text)

    return "\n".join(parts)

def extract_text_layer(pdf_path: Path) -> str:
    """Extract ONLY the PDF's embedded text layer (no OCR)."""
    parts: List[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            parts.append(text)
    return "\n".join(parts)

def parse_patient_metadata(text: str) -> Dict[str, Optional[str]]:
    metadata: Dict[str, Optional[str]] = {
        "Patient ID": None,
        "Sex": None,
        "Age": None,
        "Collection Date": None,
        "Collection Time": None,
    }

    for field, pattern in PATIENT_FIELDS.items():
        match = pattern.search(text)
        if match:
            metadata[field] = match.group(1).strip()

    sex_match = re.search(r"\b(Male|Female)\b", text, re.IGNORECASE)
    if sex_match:
        metadata["Sex"] = sex_match.group(1).title()

    age_fallback = re.search(r"\b(\d{1,3})\s+(Male|Female)\b", text, re.IGNORECASE)
    if metadata["Age"] is None and age_fallback:
        metadata["Age"] = age_fallback.group(1)

    date_match = re.search(r"(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})", text)
    if date_match:
        metadata["Collection Date"] = date_match.group(1)

    time_match = re.search(r"(\d{1,2}:\d{2}\s?(?:AM|PM)?)", text, re.IGNORECASE)
    if time_match:
        metadata["Collection Time"] = time_match.group(1)

    return metadata


def parse_measurements_from_seca_ocr(ocr_text: str) -> Dict[str, Optional[float]]:
    """
    Parse SECA measurement values from OCR text based purely on the
    order of numeric values in the cropped OCR region (no date anchor).
    """

    # Extract all numbers (ints or floats)
    nums = re.findall(r"\d+(?:\.\d+)?", ocr_text)
    values = [float(n) for n in nums]

    # List of fields, in exact order they appear in the cropped OCR block
    field_defs: List[str] = [
        "Fat Mass (kg)",
        "Fat Mass (%)",
        "Fat Mass Index (kg/m^2)",
        "Fat-Free Mass (kg)",
        "Fat-Free Mass (%)",
        "Fat-Free Mass Index (kg/m^2)",
        "Skeletal Muscle Mass (kg)",
        "Right Arm (kg)",
        "Left Arm (kg)",
        "Right Leg (kg)",
        "Left Leg (kg)",
        "Torso (kg)",
        "Visceral Adipose Tissue",
        "Body Mass Index (kg/m^2)",
        "Height (m)",
        "Weight (kg)",
        "Total Body Water (L)",
        "Total Body Water (%)",
        "Extracellular Water (L)",
        "Extracellular Water (%)",
        "ECW/TBW (%)",
        "Resting Energy Expenditure (kcal/day)",
        "Energy Consumption (kcal/day)",
        "Phase Angle (deg)",
        "Phase Angle Percentile",
        "Resistance (Ohm)",
        "Reactance (Ohm)",
        "Physical Activity Level",
    ]

    measurements = {name: None for name in field_defs}

    # Fill what we can
    for i, name in enumerate(field_defs):
        if i < len(values):
            measurements[name] = values[i]

    return measurements

def extract_pdf_data(pdf_path: Path) -> Dict[str, Optional[float]]:

    # --- 1. HEADER TEXT (for Patient ID, Sex, Age, Date, Time) ---
    text_layer = extract_text_layer(pdf_path)
    normalized_header_text = collapse_whitespace(text_layer)

    # --- 2. OCR TEXT (cropped region, numbers only) ---
    ocr_text = extract_pdf_text(pdf_path)

    # --- 3. DEBUG OUTPUT (optional) ---
    debug_txt = pdf_path.with_suffix(".ocr.txt")
    debug_txt.write_text(ocr_text, encoding="utf-8")

    # --- 4. Build the row ---
    row: Dict[str, Optional[float]] = {}
    row.update(parse_patient_metadata(normalized_header_text))   # header from TEXT layer
    row.update(parse_measurements_from_seca_ocr(ocr_text))       # numbers from OCR
    row["Source File"] = pdf_path.name

    return row

def select_pdf_files() -> List[Path]:
    root = Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(
        title="Select SECA PDF files", filetypes=[("PDF files", "*.pdf")]
    )
    root.update()
    return [Path(path) for path in file_paths]


def select_output_path() -> Optional[Path]:
    root = Tk()
    root.withdraw()
    directory = filedialog.askdirectory(title="Select download folder")
    root.update()
    if not directory:
        return None
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return Path(directory) / f"seca_measurements_{timestamp}.xlsx"


def show_message(title: str, message: str) -> None:
    root = Tk()
    root.withdraw()
    messagebox.showinfo(title, message)
    root.destroy()


def main() -> None:
    pdf_files = select_pdf_files()
    if not pdf_files:
        show_message("SECA Data Converter", "No PDF files were selected.")
        return

    output_path = select_output_path()
    if output_path is None:
        show_message("SECA Data Converter", "No download folder was selected.")
        return

    rows = []
    for pdf in pdf_files:
        try:
            rows.append({"Source File": pdf.name, **extract_pdf_data(pdf)})
        except Exception as exc:  # pragma: no cover - user feedback path
            show_message(
                "Parsing error",
                f"Could not parse '{pdf.name}'.\nError: {exc}",
            )
            return

    df = pd.DataFrame(rows)
    df.to_excel(output_path, index=False)
    show_message(
        "SECA Data Converter",
        f"Successfully saved data for {len(rows)} file(s) to:\n{output_path}",
    )


if __name__ == "__main__":
    main()
