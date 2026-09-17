# 📊 DataCleaner Studio (Portfolio Demo Edition)

<div align="center">

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/deploy?repository=tarhan-programcilik/veri-temizleme-araci&branch=main&mainModule=app.py)
[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Upwork](https://img.shields.io/badge/Upwork-Available%20for%20Hire-14A800?logo=upwork&logoColor=white)](https://www.upwork.com)

**An interactive no-code data preparation & exploratory data analysis (EDA) web application built for client portfolio showcase.**

[Live Demo](#-live-demo) • [Key Features](#-key-features) • [Demo Restrictions](#-portfolio-demo-restrictions) • [Hire Me on Upwork](#-hire-me-on-upwork) • [License](#-license)

</div>

---

## 🌟 Overview

Working with raw, messy datasets is one of the most time-consuming hurdles in any analytics workflow. **DataCleaner Studio** demonstrates how custom Python & Streamlit applications can automate tedious spreadsheet preparation, cleaning, and reporting tasks for businesses and data teams.

> 💼 **Note**: This repository contains the **Portfolio Showcase Edition**. For enterprise-scale implementations, automated scheduled ETL pipelines, cloud database sync, or custom internal SaaS dashboards, feel free to contact me directly on Upwork!

---

## 🔒 Portfolio Demo Restrictions

To showcase capabilities while preventing unauthorized commercial misuse, this live demo includes deliberate safeguard limits:
- **1-File Trial Quota**: Users can test and evaluate 1 dataset per session.
- **200 Rows Cap**: Datasets with more than 200 rows are automatically truncated to the first 200 rows for preview.
- **2 MB File Size Limit**: Uploads are restricted to 2MB.
- **Attribution Watermark**: Exported Excel spreadsheets include a demo portfolio signature.

---

## ✨ Key Features

### 1. 📂 Versatile File Ingestion
- **Excel (`.xlsx`, `.xls`)**: Multi-sheet file reader with interactive sheet selection.
- **CSV (`.csv`)**: Configurable delimiter (`,`, `;`, `\t`, `|`) and character encodings (`utf-8`, `utf-8-sig`, `latin1`, `iso-8859-9`, `cp1254`).

### 2. 📈 Instant Data Profiling & EDA
- **Key Metrics Dashboard**: Total rows, columns, null rate %, duplicates, and memory footprint.
- **Data Preview Modes**: First 10, last 10, random sample, or complete dataset.
- **Column Health Table**: Types, non-null counts, missing rates, unique value counts, and samples.
- **Missing Value Distribution**: Interactive visual bar chart showing missing data across all columns.
- **Statistical Summaries**: Distribution metrics (`count`, `mean`, `std`, `min`, `IQR`, `max`).

### 3. 🧹 Non-Destructive Data Cleaning Suite
- **Duplicate Removal**: Remove duplicate records across all columns or specific subsets (keep first or last).
- **Missing Value Management**:
  - Drop missing rows (by `any`, `all`, or target subset).
  - Smart imputation: **Mean**, **Median**, **Mode**, **Forward Fill (ffill)**, **Backward Fill (bfill)**, or **Custom Constant Value**.
- **Column Operations**: Drop redundant columns, auto-normalize headers to `snake_case`, and rename columns on the fly.
- **Type Casting**: Numeric/Float, Integer, String/Text, Datetime, and Categorical conversions.
- **String Cleaning**: Strip leading/trailing whitespace, convert text to lowercase, UPPERCASE, or Title Case.
- **Outlier Filtering**: Outlier elimination using **IQR (Interquartile Range)** with customizable multipliers or custom min/max bounds.
- **Action History**: Timestamped audit trail with a 1-click **"Reset to Original"** button.

### 4. 💾 Branded Export & Download
- **Styled Excel (`.xlsx`)**: Generated with `openpyxl` featuring dark navy header fills, bold white typography, centered alignments, and auto-adjusted column widths.
- **Standardized CSV (`.csv`)**: Export with custom delimiters and `utf-8-sig` encoding (ensures seamless opening in Microsoft Excel without character corruption).

---

## 💼 Hire Me on Upwork

Need custom data solutions, end-to-end automations, or internal company tools?

- 🛠️ **Custom Web Dashboards**: Streamlit, Dash, Gradio, React
- ⚙️ **Data Pipelines & ETL**: Python, Pandas, Polars, Airflow, Dagster
- 🌐 **Web Scraping & Automation**: Playwright, Selenium, Beautiful Soup, Scrapy
- 🤖 **AI & LLM Integration**: LangChain, LlamaIndex, OpenAI API, Anthropic, Gemini API
- 🗄️ **Database Engineering**: PostgreSQL, MySQL, Supabase, BigQuery, Snowflake

👉 **[Contact Me on Upwork to Discuss Your Project](https://www.upwork.com)**

---

## 🚀 Quick Start (Local Run)

```bash
# Clone the repository
git clone https://github.com/tarhan-programcilik/veri-temizleme-araci.git
cd veri-temizleme-araci

# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

---

## 📄 License

Distributed under the **MIT License** - see [LICENSE](LICENSE) for details.
