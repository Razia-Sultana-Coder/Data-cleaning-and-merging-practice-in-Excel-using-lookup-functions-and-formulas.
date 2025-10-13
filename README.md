# 🧮 Excel Data Cleaning & Merging Project — Product Sales

This project demonstrates essential **data cleaning and merging techniques** in Excel — one of the most common tasks for a **Data Analyst**.  
The dataset (`product_sales.xlsx`) contains three tables: `orders`, `products`, and `customer`.

---

## 🧰 Project Overview

### 🔹 Objective
Clean and merge raw data to prepare it for analysis.  
The workflow covers:
- Data type correction  
- Handling missing and inconsistent entries  
- Removing duplicates  
- Applying lookup functions (VLOOKUP, INDEX-MATCH, XLOOKUP)  
- Calculating derived columns (e.g., total price)

---

## 🧹 Task 1: Data Cleaning

### Orders Table
- Removed duplicate `order_id`
- Converted `product_id` to Text
- Fixed incorrect `qty` entries ending with “Q”
- Replaced blank `qty` values with **"Not Available"**

### Products Table
- Trimmed extra spaces in `product_name`
- Split `price (in Rs)` column to extract numeric values
- Renamed numeric column to **price_in_INR**

### Customer Table
- Converted all `customer_name` entries to lowercase
- Pasted values permanently (no formulas)

---

## 🔗 Task 2: Data Merging

| Function | Purpose | Example Formula |
|-----------|----------|----------------|
| `VLOOKUP()` | Retrieve customer names | `=VLOOKUP(C2, customer!A:B, 2, FALSE)` |
| `INDEX-MATCH()` | Retrieve product names | `=INDEX(products!B:B, MATCH(D2, products!A:A, 0))` |
| `XLOOKUP()` | Retrieve product prices | `=XLOOKUP(D2, products!A:A, products!C:C)` |
| `Formula` | Calculate total price | `=E2*G2` |

---

## 📊 Final Output

| order_id | date | customer_id | product_id | qty | customer_name | product_name | price_in_INR | total_price |
|-----------|------|--------------|-------------|------|----------------|----------------|----------------|---------------|

---

## 🖼️ Screenshots

See `/screenshots` folder for:
- Data cleaning process  
- Lookup formulas in action  
- Final merged dataset  

---

## 🧩 Tools Used
- Microsoft Excel / Google Sheets
- Functions: `TRIM()`, `VLOOKUP()`, `INDEX-MATCH()`, `XLOOKUP()`
- Data type formatting and cleaning tools

---

## 🧑‍💻 Author
**Razia Sultana**  
Aspiring Data Analyst | Excel • Power BI • SQL • Python  

📧 razia.uestc@gmail.com  
🌐 www.linkedin.com/in/razia-sultana-bd 
