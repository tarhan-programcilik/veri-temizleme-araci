# 📊 DataCleaner Studio

<div align="center">

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/deploy?repository=tarhan-programcilik/veri-temizleme-araci&branch=main&mainModule=app.py)
[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Pandas](https://img.shields.io/badge/Pandas-2.0%2B-150458?logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![OpenPyXL](https://img.shields.io/badge/OpenPyXL-3.1%2B-brightgreen)](https://openpyxl.readthedocs.io/)

**An intuitive, no-code web application for automated data cleaning, exploratory data analysis (EDA), and professionally styled Excel/CSV exporting.**

[Live Demo](#-live-demo) • [Key Features](#-key-features) • [Quick Start](#-quick-start) • [How to Deploy](#-how-to-deploy-to-streamlit-cloud) • [License](#-license)

</div>

---

## 🌟 Overview

Working with dirty data is one of the most tedious parts of any analytical workflow. **DataCleaner Studio** eliminates repetitive data preparation tasks by providing a visual, interactive interface to:
- Inspect and profile raw datasets instantly.
- Perform robust data cleaning operations without writing boilerplate code.
- Export clean datasets with professionally styled headers and auto-sized columns.
- Maintain full auditability with a real-time change log and 1-click reset.

---

## ✨ Key Features

### 1. 📂 Versatile File Ingestion
- **Excel Support (`.xlsx`, `.xls`)**: Multi-sheet file reader with interactive sheet selection.
- **CSV Support (`.csv`)**: Configurable delimiter (`,`, `;`, `\t`, `|`) and character encodings (`utf-8`, `utf-8-sig`, `latin1`, `iso-8859-9`, `cp1254`).

### 2. 📈 Instant Data Profiling & Summary
- **Key Metrics Dashboard**: Row count, column count, total nulls, missing value %, duplicate rows, and deep memory footprint.
- **Data Preview**: Browse top N, bottom N, random sample, or full dataset.
- **Column Health Table**: Per-column data types, non-null counts, missing value rates, unique value counts, and first sample values.
- **Missing Value Distribution**: Interactive bar charts highlighting columns with data gaps.
- **Statistical EDA**: Comprehensive distribution metrics (`count`, `mean`, `std`, `min`, `IQR`, `max`) for numeric and categorical attributes.

### 3. 🧹 Robust Data Cleaning Suite
- **Duplicate Removal**: Drop duplicate records across all columns or target specific subset columns (preserve first or last record).
- **Missing Value Handling (Imputation)**:
  - Drop missing rows (by `any`, `all`, or target subset).
  - Smart imputation: **Mean**, **Median**, **Mode**, **Forward-fill (ffill)**, **Backward-fill (bfill)**, or **Custom Constant Value**.
- **Column Management**:
  - Remove redundant columns.
  - Automatic column name normalization (converts to `snake_case`, strips whitespace, cleans non-standard characters).
  - Rename individual columns on the fly.
- **Type Casting**: Seamless conversion between Numeric/Float, Integer, String/Text, Datetime, and Categorical types.
- **String Cleaning**: Strip leading/trailing whitespace, convert text to lowercase, UPPERCASE, or Title Case.
- **Outlier Detection & Removal**: Filter extreme outliers using **IQR (Interquartile Range)** with customizable multipliers or custom min/max bounds.
- **Safe Sandbox**: Real-time action logging with a 1-click **"Reset to Original"** fallback.

### 4. 💾 Styled Export & Download
- **Styled Excel (`.xlsx`)**: Generated with `openpyxl` featuring dark navy header fills, bold white typography, centered alignments, and automated column width adjustments for clean presentation.
- **Standardized CSV (`.csv`)**: Export with custom delimiters and `utf-8-sig` encoding (ensures seamless opening in Microsoft Excel without character corruption).
- **Audit / Change Log**: Comprehensive timestamped audit trail of all performed operations.

---

## 🚀 Quick Start

### Prerequisites
- Python 3.10 or higher
- pip package manager

### Local Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/tarhan-programcilik/veri-temizleme-araci.git
   cd veri-temizleme-araci
   ```

2. **Install required dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Launch the application:**
   ```bash
   streamlit run app.py
   ```
   The application will automatically open in your default browser at `http://localhost:8501`.

---

## ☁️ How to Deploy to Streamlit Cloud

Deploying your own instance to **Streamlit Community Cloud** is 100% free and takes less than a minute:

1. Fork or push this repository to your GitHub account.
2. Visit **[share.streamlit.io](https://share.streamlit.io)** and log in with your GitHub account.
3. Click **"New app"** and configure:
   - **Repository:** `your-username/veri-temizleme-araci`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. Click **"Deploy!"**. Your app will be live with a public URL accessible worldwide!

---

## 🛠️ Tech Stack

- **Framework**: [Streamlit](https://streamlit.io/) (High-performance web app framework)
- **Data Wrangling**: [Pandas](https://pandas.pydata.org/) (High-performance data manipulation)
- **Spreadsheet Engine**: [OpenPyXL](https://openpyxl.readthedocs.io/) (Rich Excel spreadsheet styling & formatting)

---

## 🧪 Sample Datasets

The repository includes sample dirty datasets to test features immediately:
- `sample_data.xlsx` (Excel version with missing entries, duplicates, and outliers)
- `sample_data.csv` (CSV version with delimiter and whitespace inconsistencies)

---

## 📄 License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.
