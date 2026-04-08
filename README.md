#  NCLT Legal Data Extraction & Matching Pipeline

##  Overview

This project implements an end-to-end data extraction and processing pipeline for National Company Law Tribunal (NCLT) cause lists.

It automates:

* Web scraping of cause list PDFs
* Structured data extraction from semi-structured documents
* Data normalization and transformation
* Entity resolution using fuzzy matching
* Generation of enriched Excel reports

---

##  Key Capabilities

###  Web Automation

* Automated scraping using Selenium
* Dynamic CAPTCHA handling
* Multi-court and date-range support

###  PDF Parsing

* Table extraction using pdfplumber
* Multi-line row reconstruction
* Robust handling of inconsistent PDF structures

###  Parallel Processing

* Concurrent PDF downloads using ThreadPoolExecutor
* Optimized I/O operations for faster execution

###  Entity Resolution

* Fuzzy matching using RapidFuzz
* Token-based similarity scoring
* Configurable match thresholds

###  Data Output

* Structured Excel reports using OpenPyXL
* Multi-sheet output (raw + matched data)
* Styled headers and formatted cells

###  UI Layer

* Tkinter-based GUI
* Date selection and court filtering
* Real-time logging interface

---

##  Architecture

```text
User Input (UI)
      ↓
Web Scraping (Selenium)
      ↓
PDF Download (Requests)
      ↓
PDF Parsing (pdfplumber)
      ↓
Data Cleaning & Structuring (Pandas)
      ↓
Fuzzy Matching (RapidFuzz)
      ↓
Excel Report Generation (OpenPyXL)
```

---

##  Installation

```bash
git clone https://github.com/your-username/NCLT-Legal-Data-Extractor.git
cd NCLT-Legal-Data-Extractor
pip install -r requirements.txt
```

---

##  Usage

```bash
python src/nclt_extractor.py
```

* Select date range
* Choose court bench
* Monitor progress via UI
* Output will be generated automatically

---

##  Output

* Structured Excel reports containing:

  * Extracted case data
  * Cleaned party names
  * Fuzzy-matched entities
  * Highlighted headers and formatted sheets

---

##  Configuration Notes

* Update file paths for:

  * `DOWNLOAD_FOLDER`
  * `MASTER_FILE`

* Ensure Chrome browser is installed for Selenium

---

##  Performance Considerations

* Uses multithreading for I/O-bound operations
* Handles large PDF batches efficiently
* Optimized for real-world legal datasets

---

##  Future Enhancements

* Headless execution mode
* API-based data ingestion
* Database integration (PostgreSQL / MongoDB)
* Dashboard layer (Streamlit / Power BI)
* Advanced NLP-based entity resolution

---

##  Tech Stack

* Python
* Selenium
* Pandas
* pdfplumber
* RapidFuzz
* OpenPyXL
* Tkinter


