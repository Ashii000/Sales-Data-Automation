# Sales-Data-Automation
# 🚀 Sales Data Automation & Cleaning Project

This project is a complete Python-based solution for cleaning, analyzing, and reporting on raw sales data. It automates the entire process, transforming a messy Excel file into a multi-sheet, professional report in seconds.

## 📝 Problem
Manually cleaning sales data is slow, repetitive, and prone to errors. Data often contains duplicates, missing values, inconsistent formatting, and invalid entries, making analysis difficult.

## 💡 Solution
This Python script automates the entire "Extract, Transform, Load" (ETL) process:
1.  **Extract:** Loads data from a raw `.xlsx` file.
2.  **Transform:** Performs a deep cleaning and validation process.
3.  **Load:** Generates a final, clean multi-sheet Excel report and a text-based summary.

---

## ✅ Key Features

* **Data Cleaning:**
    * Standardizes customer names to `Title Case`.
    * Removes all duplicate transactions.
    * Fills missing values in the 'Notes' column with `No notes`.
    * Validates 'Amount' data, ensuring it is numeric and removing invalid (0 or negative) entries.
    * Converts all 'Date' entries to a consistent format.
* **Data Analysis:**
    * Generates a high-level financial summary (Total Sales, Avg. Sale, etc.).
    * Creates a pivot table for **Category Summary** (Total Sales, Avg. Sale, Count).
    * Creates a pivot table for **Customer Summary** (Total Spent, Purchases).
    * Creates a **Daily Report** (Total Sales, Count per day).
* **Professional Output:**
    1.  **`cleaned_sales_data.xlsx`:** A multi-sheet Excel file containing:
        * `Cleaned Data`: The full, clean dataset.
        * `Category Summary`: Pivot table of sales by category.
        * `Customer Summary`: Pivot table of sales by customer.
        * `Daily Report`: Pivot table of sales by date.
    2.  **`sales_report.txt`:** A text file with a timestamped summary of all findings, perfect for quick review.

---

## 🛠️ Tech Stack

* **Python**
* **pandas** (For data loading, manipulation, and analysis)
* **xlsxwriter** (For creating the multi-sheet Excel report)

---

## 🏃‍♂️ How to Run

1.  **Clone the repository:**
    ```bash
git clone https://github.com/Ashi000/Sales_Data_Automation.git
    ```
2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
3.  **Add your data:**
    * Place your raw sales file in the folder and name it `raw_sales_data.xlsx`.
    * *(A sample `raw_sales_data.xlsx` is included in this repository for testing.)*
4.  **Run the script:**
    ```bash
    python sales_automation.py
    ```
5.  **Check the output:**
    * `cleaned_sales_data.xlsx`
    * `sales_report.txt`
