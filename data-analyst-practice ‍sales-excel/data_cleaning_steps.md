# 🧹 Data Cleaning Steps

This document outlines the step-by-step process used to clean and prepare the dataset for analysis.  
The goal was to ensure accuracy, consistency, and readiness for visualization and reporting.

---

## 🗂️ 1. Data Import

- Imported the raw dataset into **Excel** (or Power BI / Python environment if applicable).
- Checked for incorrect file encodings and verified column headers.
- Ensured that all data types (Date, Text, Numeric) were correctly recognized.

---

## 🔍 2. Initial Data Inspection

- Reviewed the first few rows to understand structure and context.
- Verified column names and standardized them to **snake_case** (e.g., `product_id`, `product_name`).
- Identified missing values, inconsistent spellings, and potential outliers.

---

## 🧭 3. Removing Duplicates

- Used Excel’s **Remove Duplicates** feature or `=UNIQUE()` function.
- Checked for duplicate `product_id` and `customer_id`.
- Kept only the most relevant or latest entries.

---

## 🧼 4. Handling Missing Values

- Replaced missing text values with `"Unknown"` or `"N/A"`.
- For numeric columns (like `price` or `quantity`), filled missing values with:
  - The average/median of the column, or  
  - 0 if the field was optional.
- For date fields, used consistent format: **DD-MM-YYYY**.

---

## 🧾 5. Text Standardization

- Trimmed leading/trailing spaces using `=TRIM()`.
- Fixed inconsistent casing using:
  - `=PROPER()` for names
  - `=UPPER()` or `=LOWER()` where required.
- Replaced special characters and typos for uniformity.

---

## 🔢 6. Data Type Corrections

- Converted text-based numbers to numeric type using `=VALUE()`.
- Ensured all dates were true date values, not text.
- Checked for currency symbols or units within cells and removed them before analysis.

---

## 🔗 7. Creating Lookup Tables

Created reference (lookup) tables to avoid data repetition and maintain consistency:
- **Products Table:** `product_id`, `product_name`, `price_in_inr`
- **Customers Table:** `customer_id`, `customer_name`, `location`
- Linked all related datasets using formulas like:
  ```excel
  =INDEX(products[product_name], MATCH([@product_id], products[product_id], 0))

or modern:
=XLOOKUP([@product_id], products[product_id], products[product_name])

🧮 8. Derived Columns

Added calculated columns such as:

Total Amount = Quantity * Unit Price

Discounted Price (if applicable)

Profit Margin, Category, etc.

Applied consistent number formatting for currency and percentages.

📊 9. Data Validation

Applied Excel Data Validation to prevent incorrect entries.

Ensured IDs were unique and text entries used dropdown lists for categories.

Cross-checked totals with raw reports to ensure consistency.

🚀 10. Final Checks and Export

Rechecked data ranges and removed any helper columns.

Ensured table names were meaningful (e.g., orders, products, customers).

Saved the final cleaned dataset as:

cleaned_data.xlsx

cleaned_data.csv

Documented all transformations in this file.

✅ Summary
Step	  	Purpose
Import		Load and inspect the raw data
Cleaning	Fix inconsistencies and missing values
Standardization	Ensure consistent formats and structure
Linking		Create relationships across tables
Validation	Check for accuracy and integrity
Export		Prepare for analysis and visualization

Author: Razia Sultana
Project: Data Analysis and Cleaning in Excel
Last Updated: October 2025
