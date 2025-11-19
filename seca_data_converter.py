"""SECA Data Converter
=======================

This script opens GUI dialogs to let a user select one or more SECA PDF
reports and a destination folder.  It then extracts the patient metadata and
measurement values defined in the project requirements and stores them in an
Excel workbook (one row per PDF).

Usage:
    python seca_data_converter.py

Dependencies:
    - pdfplumber
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
from tkinter import Tk, messagebox
from tkinter import filedialog


# Regex that captures floats/integers including optional comma decimal marks.
NUMBER_PATTERN = re.compile(r"-?\d+(?:[.,]\d+)?")


@dataclass
class MeasurementSpec:
    """Defines how to capture measurements for a label."""

    label: str
    field_names: List[str]

    def expected_count(self) -> int:
        return len(self.field_names)


PATIENT_FIELDS = {
    "Patient ID": re.compile(r"ID[:\s]+([A-Za-z0-9-]+)", re.IGNORECASE),
    "Age": re.compile(r"Age[:\s]+(\d+)", re.IGNORECASE),
}


MEASUREMENT_SPECS: List[MeasurementSpec] = [
    MeasurementSpec("Fat Mass", ["Fat Mass (kg)", "Fat Mass (%)"]),
    MeasurementSpec("Fat Mass Index", ["Fat Mass Index (kg/m^2)"]),
    MeasurementSpec("Fat-Free Mass", ["Fat-Free Mass (kg)", "Fat-Free Mass (%)"]),
    MeasurementSpec("Fat-Free Mass Index", ["Fat-Free Mass Index (kg/m^2)"]),
    MeasurementSpec("Skeletal Muscle Mass", ["Skeletal Muscle Mass (kg)"]),
    MeasurementSpec("right arm", ["Right Arm SMM (kg)"]),
    MeasurementSpec("left arm", ["Left Arm SMM (kg)"]),
    MeasurementSpec("right leg", ["Right Leg SMM (kg)"]),
    MeasurementSpec("left leg", ["Left Leg SMM (kg)"]),
    MeasurementSpec("torso", ["Torso SMM (kg)"]),
    MeasurementSpec("Visceral Adipose Tissue", ["Visceral Adipose Tissue (L)"]),
    MeasurementSpec("Body Mass Index", ["Body Mass Index (kg/m^2)"]),
    MeasurementSpec("Height", ["Height (m)"]),
    MeasurementSpec("Weight", ["Weight (kg)"]),
    MeasurementSpec("Total Body Water", ["Total Body Water (L)", "Total Body Water (%)"]),
    MeasurementSpec(
        "Extracellular Water",
        ["Extraceullar Water (L)", "Extracellular Water (%)"],
    ),
    MeasurementSpec("ECW/TBW", ["ECW/TBW (%)"]),
    MeasurementSpec(
        "Resting Energy Expenditure",
        ["Resting Energy Expenditure (kcal/day)"],
    ),
    MeasurementSpec("Energy Consumption", ["Energy Consumption (kcal/day)"]),
    MeasurementSpec("Phase Angle", ["Phase Angle (°)"]),
    MeasurementSpec("Resistance", ["Resistance (Ω)"]),
    MeasurementSpec("Reactance", ["Reactance (Ω)"]),
    MeasurementSpec("Physical Activity Level", ["Physical Activity Level"]),
]


def normalize_number(token: str) -> float:
    token = token.replace(",", ".")
    return float(token)


def extract_numbers_near_label(text: str, label: str, count: int) -> List[float]:
    """Return up to *count* numbers that appear shortly after *label*."""

    text_lower = text.lower()
    idx = text_lower.find(label.lower())
    if idx == -1:
        return []
    snippet = text[idx : idx + 200]
    return [normalize_number(match) for match in NUMBER_PATTERN.findall(snippet)[:count]]


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

    date_match = re.search(r"(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})", text)
    if date_match:
        metadata["Collection Date"] = date_match.group(1)

    time_match = re.search(r"(\d{1,2}:\d{2}\s?(?:AM|PM)?)", text, re.IGNORECASE)
    if time_match:
        metadata["Collection Time"] = time_match.group(1)

    return metadata


def parse_measurements(text: str) -> Dict[str, Optional[float]]:
    measurements: Dict[str, Optional[float]] = {
        field: None
        for spec in MEASUREMENT_SPECS
        for field in spec.field_names
    }

    for spec in MEASUREMENT_SPECS:
        values = extract_numbers_near_label(text, spec.label, spec.expected_count())
        for (field_name, value) in zip(spec.field_names, values):
            measurements[field_name] = value

    return measurements


def extract_pdf_data(pdf_path: Path) -> Dict[str, Optional[float]]:
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join(filter(None, (page.extract_text() for page in pdf.pages)))

    row: Dict[str, Optional[float]] = {}
    row.update(parse_patient_metadata(full_text))
    row.update(parse_measurements(full_text))
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
