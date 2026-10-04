# 📊 Data Cleaning and Transformation Using Microsoft Excel

## 📌 Project Overview

This project demonstrates **data cleaning, preprocessing, transformation, and formatting using Microsoft Excel**.

Real-world datasets often contain missing values, duplicate records, inconsistent text formatting, spelling errors, and poorly structured fields. Before performing meaningful analysis, these data-quality issues need to be identified and corrected.

In this project, a **Product Dataset** was cleaned and transformed using Excel's built-in data-cleaning and formatting features.

The objective is to prepare the dataset for further **data analysis and visualization** by improving its accuracy, consistency, and usability.

---

## 🎯 Project Objectives

The main objectives of this project are:

* Identify and handle missing values
* Detect inconsistent text formatting
* Correct category spelling errors
* Identify and remove duplicate records
* Split structured Product ID information into separate columns
* Merge Brand Name and Product Name
* Apply appropriate number and date formatting
* Use conditional formatting to improve data readability
* Prepare a clean dataset for further analysis

---

## 📁 Dataset Description

The dataset contains information about different products.

### Dataset Attributes

| Column             | Description                                                       |
| ------------------ | ----------------------------------------------------------------- |
| Product ID         | Unique product identifier containing date and country information |
| Manufacturing Date | Manufacturing date extracted from Product ID                      |
| Country Code       | Country code extracted from Product ID                            |
| Product Name       | Name of the product                                               |
| Brand Name         | Brand/manufacturer of the product                                 |
| Product Brand      | Combined Product Name and Brand Name                              |
| Price ($)          | Product price                                                     |
| Quantity           | Available product quantity                                        |
| Category           | Product category                                                  |

### Product Categories

The dataset contains the following categories:

* Electronics
* Fashion
* Kitchen
* Accessories
* Outdoor

---

## 🔍 Dataset Summary

After inspecting the workbook:

* **Total product records:** 34
* **Columns used for analysis:** 9
* **Duplicate records identified:** 3 duplicate pairs
* **Missing values in populated records:** None
* **Minimum product price:** $30
* **Maximum product price:** $1,000
* **Average product price:** Approximately $306.23

### Category Distribution

| Category    | Records |
| ----------- | ------: |
| Electronics |      14 |
| Accessories |       7 |
| Fashion     |       6 |
| Kitchen     |       5 |
| Outdoor     |       2 |

---

# 🧹 Data Cleaning Process

## 1. Handling Missing Values

### Price Column

The `Price` column was checked for blank or missing values.

The appropriate approach for missing prices would be:

1. Identify products with missing prices.
2. Check whether the price can be obtained from a reliable source.
3. If a reliable value is unavailable, use an appropriate statistical method such as the category-wise or overall median/average, depending on the analysis requirement.
4. Document any imputed values.

For this dataset, there were **no missing price values among the populated product records**.

### Category Column

The `Category` column was also checked for missing values.

If a category were missing, possible approaches would include:

* Infer the category from the Product Name.
* Compare similar products and use the most appropriate category.
* Use the most frequent category only when justified.
* Leave it as `Unknown` if the category cannot be reliably determined.

The populated dataset did not contain missing category values.

---

# ✏️ 2. Correcting Inconsistent Data

## Product Name / Brand Formatting

The Product Name and Brand Name fields were reviewed for consistency.

Excel's **Find and Replace** feature can be used to correct formatting inconsistencies.

Example:

```text
Find: hp
Replace with: HP
```

Brand capitalization can also be standardized where necessary.

For example:

```text
Hp → HP
```

The objective is to ensure that the same product or brand is represented consistently throughout the dataset.

---

## Category Standardization

The `Category` column was reviewed for spelling mistakes and inconsistent values.

Expected standardized categories:

```text
Electronics
Fashion
Kitchen
Accessories
Outdoor
```

Excel's **Find & Replace** function can be used to correct any misspelled categories.

Example:

```text
Electroncs → Electronics
Accessorie → Accessories
```

The final dataset contains standardized category names.

---

# 🔁 3. Removing Duplicate Records

Duplicate rows were identified by comparing the **entire row**, including:

* Product ID
* Manufacturing Date
* Country Code
* Product Name
* Brand Name
* Product Brand
* Price
* Quantity
* Category

### Duplicate records identified

The dataset contains duplicate entries for:

* `17-JUN-IN – Laptop – HP`
* `16-APR-ES – Headphones – Bose`
* `21-AUG-CA – Laptop Bag – Samsonite`

These duplicate rows should be removed using:

**Excel → Data → Remove Duplicates**

The goal is to retain only one valid occurrence of each complete record.

---

# 🔀 4. Splitting and Merging Data

## Splitting Product ID

The original `Product ID` follows a structure similar to:

```text
28-JAN-US
```

The Product ID contains:

```text
28-JAN → Manufacturing Date
US     → Country Code
```

Therefore, the Product ID was separated into two fields:

### Manufacturing Date

Example:

```text
28-JAN
15-FEB
03-MAR
11-APR
```

### Country Code

Example:

```text
US
UK
IN
CA
DE
ES
AU
```

Excel techniques such as **Text to Columns**, `LEFT`, `RIGHT`, and `TEXTBEFORE`/`TEXTAFTER` can be used depending on the Excel version.

---

## Merging Brand and Product Name

The `Brand Name` and `Product Name` columns were combined into a new column called:

```text
Product Brand
```

Example:

```text
Product Name: Laptop
Brand Name: Dell

Product Brand: Laptop Dell
```

Possible Excel formula:

```excel
=D2&" "&E2
```

This creates a more convenient field for identifying the product and its brand together.

---

# 💰 5. Number Formatting

## Price Formatting

The `Price` column was formatted as currency.

Example:

```text
1000
```

becomes:

```text
$1,000.00
```

This improves readability and makes the financial information easier to understand.

---

## Manufacturing Date Formatting

The Manufacturing Date was formatted using the required:

```text
DD-MM-YYYY
```

format.

Example:

```text
28-JAN-2026
```

becomes:

```text
28-01-2026
```

This provides a consistent date format for future analysis.

---

# 🎨 6. Conditional Formatting

## Price Column

Conditional formatting was applied to the `Price` column using a:

* Data Bar, or
* Color Scale

This makes it easier to visually identify:

* Low-priced products
* Medium-priced products
* High-priced products

For example, longer data bars represent higher product prices.

---

## Electronics Category

A custom conditional formatting rule was created for the `Category` column.

### Rule

```excel
=I2="Electronics"
```

Cells containing:

```text
Electronics
```

are highlighted.

This allows Electronics products to be identified quickly within the dataset.

---

# 🛠️ Excel Features Used

The following Microsoft Excel features were used in this project:

* Data Cleaning
* Remove Duplicates
* Find & Replace
* Text to Columns
* Excel Formulas
* Text Manipulation
* Data Formatting
* Currency Formatting
* Date Formatting
* Conditional Formatting
* Data Bars
* Color Scales
* Custom Conditional Formatting Rules

---

# 📈 Data Transformation Workflow

The overall workflow followed in this project was:

```text
Raw Product Dataset
        ↓
Data Inspection
        ↓
Check Missing Values
        ↓
Identify Inconsistencies
        ↓
Standardize Text
        ↓
Correct Category Errors
        ↓
Identify Duplicate Records
        ↓
Remove Duplicates
        ↓
Split Product ID
        ↓
Create Manufacturing Date
        ↓
Create Country Code
        ↓
Merge Product + Brand
        ↓
Format Price and Date
        ↓
Apply Conditional Formatting
        ↓
Clean Dataset
        ↓
Ready for Data Analysis
```

---

# 📂 Project Structure

```text
Excel-Data-Cleaning-Transformation/
│
├── 📄 README.md
│
├── 📊 Assignment 2 - Data Cleaning and Transformation.xlsx
│
└── 📁 screenshots/
    ├── raw-data.png
    ├── missing-values.png
    ├── duplicate-removal.png
    ├── data-transformation.png
    └── conditional-formatting.png
```

---

# 📊 Final Dataset

After completing the cleaning and transformation process, the dataset is structured with the following fields:

```text
Product ID
Manufacturing Date
Country Code
Product Name
Brand Name
Product Brand
Price ($)
Quantity
Category
```

The cleaned dataset is now more suitable for:

* Exploratory Data Analysis (EDA)
* Excel dashboards
* Pivot Tables
* Data visualization
* Business reporting
* Power BI analysis
* Further statistical analysis

---

# 💡 Key Learnings

Through this project, I learned how to:

* Identify data-quality problems
* Handle missing values
* Detect and remove duplicate records
* Standardize inconsistent text
* Correct spelling errors
* Split structured data into multiple columns
* Combine columns using Excel formulas
* Format numerical and date fields
* Apply conditional formatting
* Prepare datasets for analysis

---

# 🚀 Future Improvements

This project can be extended by performing further analysis such as:

* Category-wise sales analysis
* Average price by category
* Quantity analysis
* Country-wise product distribution
* Most expensive products
* Top brands by quantity
* Price distribution
* Pivot Table analysis
* Excel dashboard creation
* Power BI visualization

---

# 👨‍💻 Tools & Technologies

**Tool:** Microsoft Excel

**Techniques:**

```text
Data Cleaning
Data Transformation
Data Standardization
Duplicate Removal
Text Manipulation
Conditional Formatting
Number Formatting
Date Formatting
Excel Formulas
```

---

# 🏁 Conclusion

This project demonstrates the importance of **data cleaning and preprocessing before performing data analysis**.

By identifying missing values, correcting inconsistencies, removing duplicate records, restructuring Product IDs, combining product information, and applying proper formatting, the raw product dataset was transformed into a cleaner and more analysis-ready dataset.

This project strengthened my practical understanding of **Microsoft Excel as a data-cleaning and preprocessing tool** and provides a foundation for progressing toward more advanced tools such as **SQL, Python, Power BI, and data analytics workflows**.
