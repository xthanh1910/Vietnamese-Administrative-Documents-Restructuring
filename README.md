# Vietnamese Administrative Documents Restructuring

This project automates the reconstruction of Vietnamese legal and administrative documents from images/scans into fully formatted Microsoft Word (.docx) files. By combining **PaddleOCR**, a fine-tuned **ViT5** language model, and intelligent layout analysis, the system ensures that the output text is both accurate and compliant with official government standards.

---

## Video Demo
<div align="center">
  <a href="https://www.youtube.com/watch?v=JrPtboFN1zs">
    <img src="https://img.youtube.com/vi/JrPtboFN1zs/0.jpg" alt="Project Demo" style="width:100%;">
  </a>
  <p><i>Click the image above to watch the demonstration on YouTube</i></p>
</div>

---

## Key Features

* **High-Precision OCR**: Utilizes PaddleOCR to detect text boxes and extract raw coordinates.
* **AI-Powered Vietnamese Correction**: Integrates a fine-tuned **ViT5** (Seq2Seq) model to fix spelling, tone marks, and common OCR misinterpretations in legal Vietnamese text.
* **Intelligent Dictionary Post-processing**: Uses fuzzy matching to standardize agency names (Agencies) and document types (Titles).
* **Automated Layout Analysis**: 
    * Separates the Header (National Emblem, Agency Name, Document ID) from the Body.
    * Detects the primary Title and **Extended Title** (Summary of content) with automated line-balancing algorithms.
    * Groups text into paragraphs based on spacing and specific legal keywords (e.g., "Điều", "Chương", "Căn cứ").
* **Compliance with Decree 30/2020/NĐ-CP**: Automatically formats the Word output with specific margins (Top: 20mm, Bottom: 20mm, Left: 30mm, Right: 18mm), Times New Roman font (13pt/15pt), and proper line spacing.
* **Interactive Streamlit UI**: Supports multi-page processing and provides real-time layout analysis (Debug) images.

---

## Processing Pipeline

The system operates through four main stages defined in `backend.py`:

### 1. Stage 1: OCR & NLP Correction
* Scans the image using PaddleOCR to get bounding boxes.
* Processes text through the ViT5 model for initial correction.
* Applies a dictionary-based mapping for short text boxes to ensure official terminology is preserved.

### 2. Stage 2: Layout Segmentation
* Identifies the document **Title** based on text length and spatial centering.
* Determines the boundary between the Header and the Body.
* Defines the **Extended Title** range to capture content summaries.

### 3. Stage 3: Paragraph Grouping
* Merges single lines into coherent paragraphs based on four criteria: line width, vertical gaps (Normal Gap), indentation, and paragraph-break keywords.

### 4. Stage 4: Document Formatting
* Reconstructs the Header using a two-column table (3.5/6.5 ratio) for Agency and National Emblem sections.
* Applies standardized paragraph formatting (Line spacing 1.0, Space before/after 6pt).

---

## Tech Stack

* **Backend**: Python, PyTorch, Transformers (HuggingFace).
* **OCR Engine**: PaddleOCR.
* **Document Processing**: `python-docx`, `difflib`, `unicodedata`.
* **Image Processing & Visualization**: OpenCV, NumPy, Matplotlib.
* **Frontend**: Streamlit.

---

## Debugging Visuals
The system generates a visual analysis of each page to verify layout accuracy:
* **Red Boxes**: Primary Document Title.
* **Orange Boxes**: Extended Title / Content Summary.
* **Green Boxes**: Detected Paragraphs (labeled P1, P2, etc.).

---
