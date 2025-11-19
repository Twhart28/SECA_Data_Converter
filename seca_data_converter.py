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
            pil_image = page.to_image(resolution=300).original
            ocr_text = pytesseract.image_to_string(pil_image)
        except Exception:
            # OCR is best-effort; fall back to whatever text layer we have.
            ocr_text = ""
        if ocr_text.strip():
            parts.append(ocr_text)

    return "\n".join(parts)


def extract_pdf_text(pdf_path: Path) -> str:
    """Extract text from PDF using pdfplumber + Tesseract OCR."""
    parts: List[str] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Text layer (very minimal in SECA)
            text = page.extract_text() or ""
            parts.append(text)

            # OCR layer
            try:
                pil_image = page.to_image(resolution=300).original
                ocr_text = pytesseract.image_to_string(pil_image)
            except Exception:
                ocr_text = ""
            parts.append(ocr_text)

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


def parse_measurements_from_seca_ocr(full_text: str) -> Dict[str, Optional[float]]:
    """
    Parse SECA measurement values from OCR text by assuming a fixed order
    of numeric values after the '10/7/2025' style date line.
    """
    # Find the first date like 10/7/2025 or 7/10/2025
    date_match = re.search(r"\d{1,2}/\d{1,2}/\d{4}", full_text)
    if not date_match:
        return {}

    start = date_match.start()

    # Try to cut off before the footer "Page 1"
    page_match = re.search(r"Page\s+\d+", full_text)
    if page_match:
        region = full_text[start:page_match.start()]
    else:
        region = full_text[start:]

    # Extract all numbers in that region
    nums = re.findall(r"\d+(?:\.\d+)?", region)
    # First three numbers are the date pieces (month, day, year)
    if len(nums) <= 3:
        return {}

    values = [float(n) for n in nums[3:]]  # skip the date

    field_defs: List[tuple[str, int]] = [
        ("Fat Mass (kg)", 1),
        ("Fat Mass (%)", 1),
        ("Fat Mass Index (kg/m^2)", 1),
        ("Fat-Free Mass (kg)", 1),
        ("Fat-Free Mass (%)", 1),
        ("Fat-Free Mass Index (kg/m^2)", 1),
        ("Skeletal Muscle Mass (kg)", 1),
        ("Right Arm (kg)", 1),
        ("Left Arm (kg)", 1),
        ("Right Leg (kg)", 1),
        ("Left Leg (kg)", 1),
        ("Torso (kg)", 1),
        ("Visceral Adipose Tissue", 1),
        ("Body Mass Index (kg/m^2)", 1),
        ("Height (m)", 1),
        ("Weight (kg)", 1),
        ("Total Body Water (L)", 1),
        ("Total Body Water (%)", 1),
        ("Extracellular Water (L)", 1),
        ("Extracellular Water (%)", 1),
        ("ECW/TBW (%)", 1),
        ("Resting Energy Expenditure (kcal/day)", 1),
        ("Energy Consumption (kcal/day)", 1),
        ("Phase Angle (deg)", 1),
        ("Phase Angle Percentile", 1),
        ("Resistance (Ohm)", 1),
        ("Reactance (Ohm)", 1),
        ("Physical Activity Level", 1),
    ]

    total_needed = sum(count for _, count in field_defs)
    measurements: Dict[str, Optional[float]] = {
        name: None for name, _ in field_defs
    }

    if len(values) < total_needed:
        # Not enough numbers; bail out gracefully
        return measurements

    idx = 0
    for name, count in field_defs:
        # All counts are 1 in this layout
        measurements[name] = values[idx]
        idx += count

    return measurements

def extract_pdf_data(pdf_path: Path) -> Dict[str, Optional[float]]:
    full_text = extract_pdf_text(pdf_path)

    # For metadata, we only need a whitespace-collapsed version
    normalized_text = collapse_whitespace(full_text)

    row: Dict[str, Optional[float]] = {}

    # Patient metadata (ID, Sex, Age, Date, Time)
    row.update(parse_patient_metadata(normalized_text))

    # SECA measurement values from OCR
    row.update(parse_measurements_from_seca_ocr(full_text))

    # Always include source file name so you can trace back
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
