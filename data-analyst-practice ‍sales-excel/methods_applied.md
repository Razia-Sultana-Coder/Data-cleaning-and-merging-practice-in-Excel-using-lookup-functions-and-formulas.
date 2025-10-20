# 📊 Product Sales Data Cleaning and Analysis

This project demonstrates the **end-to-end data cleaning and preparation process** for a product sales dataset using **Microsoft Excel**.  
It is designed as a hands-on exercise to strengthen **data analysis fundamentals** — including identifying bad data, applying cleaning techniques, and merging multiple tables.

---

## 📁 Dataset Overview

The dataset `product_sales.xlsx` contains three tables:

| Table | Description |
|--------|--------------|
| **orders** | Order details such as `order_id`, `order_date`, `customer_id`, `product_id`, and `qty`. |
| **products** | Product details including `product_name`, nutritional info, and `price (in Rs)`. |
| **customer** | Customer information including `customer_id` and `customer_name`. |

---

## 🧹 Data Cleaning Process

All cleaning steps were performed in **Excel**, focusing on consistency, accuracy, and data readiness.

### 🗂 Orders Table
- Removed duplicate `order_id` entries.  
- Converted `product_id` from Number → Text.  
- Fixed invalid entries in `qty` column that ended with “Q” using:
  ```excel
  =IF(RIGHT(A2,1)="Q",LEFT(A2,LEN(A2)-1),A2)
````

* Replaced blank `qty` values with `"Not Available"`.

### 🧃 Products Table

* Removed extra spaces using `=TRIM()`.
* Extracted numeric prices from text values such as `₹250`:

  ```excel
  =VALUE(SUBSTITUTE(A2,"₹",""))
OR
use Data Tab> Convert Text to Colomn> Delimited 
  ```
* Renamed the new column as `price_in_inr`.

### 👤 Customer Table

* Converted all customer names to lowercase using `=LOWER()`.
* Pasted as values to save cleaned results permanently.

---

## 🔗 Merging Tables

The cleaned tables were merged using Excel lookup functions.

| Task              | Formula                                                                         | Description                                 |
| ----------------- | ------------------------------------------------------------------------------- | ------------------------------------------- |
| Get customer name | `=VLOOKUP([@customer_id], customer, 2, FALSE)`                                  | Returns customer name from `customer` table |
| Get product name  | `=INDEX(products[product_name], MATCH([@product_id], products[product_id], 0))` | Returns product name                        |
| Get price         | `=XLOOKUP([@product_id], products[product_id], products[price_in_inr])`         | Returns product price                       |
| Calculate total   | `=[@qty] * [@price_in_inr]`                                                     | Calculates total amount per order           |

---

## 💾 Output Files

| File                         | Description                         |
| ---------------------------- | ----------------------------------- |
| `product_sales.xlsx`         | Original dataset (raw)              |
| `product_sales_cleaned.xlsx` | Cleaned and merged dataset          |
| `data_cleaning_tasks.md`     | Step-by-step cleaning documentation |
| `methods_applied.md`         | All methods applied in cleaning documentation |
| `README.md`                  | Project overview and summary        |

---

## 🧠 Key Learnings

* Identifying and correcting data quality issues
* Using Excel formulas (`VLOOKUP`, `INDEX-MATCH`, `XLOOKUP`) for table merging
* Applying structured references within Excel Tables
* Preparing a dataset for further analysis and visualization

---

## 🛠 Tools Used

* **Microsoft Excel 365**
* **Markdown (GitHub Documentation)**

---

**Author:** *Razia Sultana*
**Date:** *October 2025*
**Location:** *Dhaka, Bangladesh*
**Project Type:** *Data Cleaning & Analysis (Excel)*


