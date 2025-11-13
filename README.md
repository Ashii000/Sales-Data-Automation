# 🚀 Sales Data Automation & Cleaning Project

This project provides a **Python-based automated solution** for cleaning, analyzing, and generating professional reports from raw sales data. It transforms disorganized Excel sheets into multi-sheet analytical reports within seconds.

> **Goal:** Automate a manual 2-hour data cleaning task and reduce processing time to under **1 second**. ⏱️

---

## 📊 Project Overview — *Before vs. After*

The script converts messy, inconsistent sales data into clean and actionable business reports.

### Before: Raw, Unorganized Data

`![Raw Data Before](before.png)`
*(Example: raw_sales_data.xlsx — contains duplicates, missing values, and inconsistent formatting.)*

### After: Clean, Professional Report

`![Clean Report After](after.png)`
*(Example: cleaned_sales_data.xlsx — includes summary sheets and validated data.)*

---

## 🧩 Problem Statement

Manual cleaning of sales data is time-consuming, repetitive, and error-prone. Issues such as duplicate entries, missing fields, inconsistent formats, and invalid amounts make it difficult to perform reliable analysis.

---

## 💡 Automated Solution

The **Sales Data Automation Script** handles the complete **ETL (Extract, Transform, Load)** process:

1. **Extract:** Reads data from a raw `.xlsx` file.
2. **Transform:** Cleans, validates, and standardizes all data fields.
3. **Load:** Outputs a fully formatted Excel report with multiple summary sheets and a text-based analysis.

---

## ✅ Key Features

### 🔹 Data Cleaning

* Converts customer names to **Title Case**
* Removes duplicate records
* Fills missing entries in the *Notes* column with `"No notes"`
* Validates and filters out invalid (0 or negative) *Amount* values
* Standardizes *Date* formats across all records

### 🔹 Data Analysis

* Creates a **Financial Summary** (Total Sales, Average Sale, etc.)
* Generates pivot-based **Category Summary** (Sales by Category)
* Builds **Customer Summary** (Total Spent per Customer)
* Produces **Daily Report** (Sales by Date)

### 🔹 Output Files

1. **`cleaned_sales_data.xlsx`** – A professional Excel file with multiple sheets:

   * *Cleaned Data*
   * *Category Summary*
   * *Customer Summary*
   * *Daily Report*
2. **`sales_report.txt`** – A timestamped text summary of the key metrics.

---

## 🛠️ Tech Stack

* **Python 3.x**
* **pandas** – Data manipulation and analysis
* **xlsxwriter** – Excel report generation

---

## ⚙️ Setup Instructions

### 1. Clone the Repository

```bash
git clone https://github.com/Ashi000/Sales_Data_Automation.git
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Add Raw Data

Place your Excel file named `raw_sales_data.xlsx` in the project folder.
*(A sample dataset is already included for testing.)*

### 4. Run the Script

```bash
python sales_automation.py
```

### 5. Review Outputs

After execution, the following files are generated automatically:

* `cleaned_sales_data.xlsx`
* `sales_report.txt`

---

## 📄 License

This project is licensed under the **MIT License** — free for personal and commercial use with proper attribution.

---

## 👤 Author

**Ashiq Hussain**

* **GitHub:** [@Ashi000](https://github.com/Ashi000)
* **LinkedIn:** [ashiq-mari-5abb33277](https://www.linkedin.com/in/ashiq-mari-5abb33277)

---


