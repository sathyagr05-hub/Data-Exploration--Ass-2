# 📊 Excel Data Cleaning and Transformation

## 📌 Project Overview

This project demonstrates the process of **data cleaning, preparation, and transformation using Microsoft Excel**. The project uses a Product Dataset containing information about product IDs, product names, brands, prices, quantities, and categories.

As an aspiring **Data Analyst**, I worked on this project to develop practical skills in preparing raw data for analysis. The dataset contains common real-world data quality issues such as missing values, inconsistent text formatting, category errors, and unstructured information within a single column.

The objective of this project is to transform the raw dataset into a **clean, standardized, and analysis-ready dataset** using Microsoft Excel.

---

## 🎯 Problem Statement

Raw datasets often contain incomplete, inconsistent, duplicated, or poorly formatted information. If these issues are not addressed, they can negatively affect the accuracy of data analysis and reporting.

In this project, I performed the following data preparation activities:

* Identified and handled missing values
* Standardized inconsistent product names
* Corrected category errors and typos
* Checked and removed duplicate records
* Split Product ID into meaningful fields
* Merged Brand Name and Product Name
* Applied currency formatting to prices
* Standardized manufacturing dates
* Applied conditional formatting for better data visualization

---

## 📂 Dataset

**Dataset:** Product Dataset

The original dataset contains the following attributes:

| Column       | Description                                                      |
| ------------ | ---------------------------------------------------------------- |
| Product ID   | Unique identifier containing manufacturing date and country code |
| Product Name | Name of the product                                              |
| Brand Name   | Brand associated with the product                                |
| Quantity     | Quantity of products                                             |
| Category     | Product category                                                 |
| Price ($)    | Product price                                                    |

### 🔗 Dataset Source

[Download the Product Dataset](https://github.com/GeethaGunasekaran1/Dataset_rep/raw/refs/heads/main/Assignment%202%20-%20Data%20Cleaning%20and%20Transformation.xlsx)

---

# 🛠️ Tools & Techniques Used

### Tool

* **Microsoft Excel**

### Excel Features Used

* Data Cleaning
* Find & Replace
* `COUNTBLANK`
* `AVERAGE`
* `IF`
* `PROPER`
* Remove Duplicates
* Text to Columns
* Data Formatting
* Currency Formatting
* Date Formatting
* Conditional Formatting
* Data Bars

---

# 🔍 Data Cleaning Process

## 1. Handling Missing Values

The dataset was checked for missing values in the **Price** and **Category** columns.

### Price

Missing price values were identified using Excel's `COUNTBLANK` function.

```excel
=COUNTBLANK(D2:D35)
```

The missing price values were handled using an **average-price imputation approach**, where appropriate.

Example formula:

```excel
=IF(ISBLANK(D2),AVERAGE($D$2:$D$32),D2)
```

This replaces a missing price with the average available price while keeping existing price values unchanged.

### Category

Missing category values were identified and handled using available product information.

Where the category could not be reliably determined, it was classified as:

```text
Unknown
```

This avoids making unsupported assumptions about the product category.

---

# 2. Correcting Inconsistent Data

## Product Name

The `Product Name` column contained inconsistent capitalization.

Examples included:

```text
laptop
Laptop
smartphone
Smartphone
headphones
Headphones
```

The product names were standardized using Excel's `PROPER` function.

Example:

```excel
=PROPER(B2)
```

This converts inconsistent capitalization into a standardized format.

### Example

```text
laptop → Laptop
smartphone → Smartphone
headphones → Headphones
```

---

## Category

The `Category` column contained inconsistent or incorrect category values.

For example:

```text
Electroni → Electronics
```

The **Find and Replace** feature was used to correct the identified category typo and standardize the data.

---

# 3. Removing Duplicates

The dataset was checked for duplicate records using Excel's **Remove Duplicates** feature.

The complete row was considered when identifying duplicates based on:

* Product ID
* Product Name
* Brand Name
* Price
* Quantity
* Category

Duplicate records, where present, were removed to prevent the same record from being counted more than once during future analysis.

---

# 4. Splitting and Merging Data

## Splitting Product ID

The `Product ID` contains multiple pieces of information.

For example:

```text
28-JAN-US
```

The Product ID was separated into:

| Manufacturing Date | Country Code |
| ------------------ | ------------ |
| 28-JAN             | US           |

The same process was applied to the remaining Product IDs.

### Result

The following columns were created:

* **Manufacturing Date**
* **Country Code**

This makes the information easier to analyze separately.

---

## Merging Brand Name and Product Name

The `Brand Name` and `Product Name` columns were combined into a new column called:

**Product Brand**

Example:

```text
Brand Name: Dell
Product Name: Laptop
```

Result:

```text
Dell Laptop
```

This creates a more descriptive product field that can be used in reporting and analysis.

---

# 5. Number Formatting

## Price Formatting

The `Price` column was formatted as a **currency value** to improve readability and consistency.

Example:

```text
1000 → $1,000.00
80   → $80.00
130  → $130.00
```

---

## Manufacturing Date Formatting

The Manufacturing Date was converted into the required:

```text
DD-MM-YYYY
```

format.

Example:

```text
28-JAN → 28-01-2026
15-FEB → 15-02-2026
03-MAR → 03-03-2026
```

This provides a consistent date format for further analysis.

---

# 6. Conditional Formatting

Conditional formatting was applied to make patterns in the dataset easier to identify visually.

## Price Column

A **Data Bar** / color-based conditional formatting rule was applied to the Price column.

This allows high and low product prices to be compared visually.

For example:

```text
Low Price     → Short Data Bar
Medium Price  → Medium Data Bar
High Price    → Long Data Bar
```

---

## Electronics Category

A custom conditional formatting rule was created for the `Category` column to highlight products where:

```text
Category = Electronics
```

This makes Electronics products easier to identify within the dataset.

---

# 📊 Project Workflow

The overall data cleaning workflow followed these steps:

```text
Raw Product Dataset
        ↓
Identify Missing Values
        ↓
Handle Missing Price & Category
        ↓
Correct Product Name Formatting
        ↓
Fix Category Errors
        ↓
Remove Duplicate Records
        ↓
Split Product ID
        ↓
Merge Brand + Product Name
        ↓
Format Price & Date
        ↓
Apply Conditional Formatting
        ↓
Clean & Analysis-Ready Dataset
```

---

# 📁 Workbook Structure

The Excel workbook contains separate sheets demonstrating each stage of the data preparation process.

| Sheet                            | Purpose                                           |
| -------------------------------- | ------------------------------------------------- |
| **Cleaned Data**                 | Handles missing values and creates cleaned fields |
| **Correcting Inconsistent Data** | Standardizes product names and categories         |
| **Removing Duplicates**          | Demonstrates duplicate checking/removal           |
| **Splitting and Merging Data**   | Splits Product ID and creates Product Brand       |
| **Number Formatting**            | Applies price and date formatting                 |
| **Conditional Formatting**       | Applies data bars and category highlighting       |

---

# 📈 Before vs After

## Before Cleaning

The raw dataset contained issues such as:

* Missing prices
* Missing categories
* Inconsistent capitalization
* Category spelling errors
* Unstructured Product ID information
* Different date representations
* Values requiring formatting

## After Cleaning

The dataset was transformed into a more structured format with:

* Missing values addressed
* Standardized product names
* Corrected category values
* Duplicate records checked
* Manufacturing Date extracted
* Country Code extracted
* Product Brand created
* Price formatted as currency
* Manufacturing Date formatted consistently
* Conditional formatting applied

---

# 🎯 Project Objectives

The main objectives of this project were:

* To understand the importance of data cleaning.
* To identify missing values in a dataset.
* To handle missing price and category information.
* To standardize inconsistent text data.
* To correct spelling errors and category inconsistencies.
* To identify and remove duplicate records.
* To restructure information stored within Product ID.
* To combine related columns into a new field.
* To apply appropriate number and date formatting.
* To use conditional formatting to improve data readability.
* To prepare a dataset for further analysis and visualization.

---

# 💡 Key Learning Outcomes

Through this project, I developed practical experience in:

* **Data Cleaning**
* **Data Preprocessing**
* **Data Transformation**
* **Missing Value Handling**
* **Data Standardization**
* **Duplicate Detection**
* **Text Transformation**
* **Find and Replace**
* **Excel Formulas**
* **Date Formatting**
* **Currency Formatting**
* **Conditional Formatting**
* **Data Preparation for Analysis**

---

# 🚀 Future Scope

After cleaning the dataset, it can be used for further analysis and visualization.

Possible future analysis includes:

* Product price analysis
* Category-wise product analysis
* Brand-wise analysis
* Quantity analysis
* Country-wise product analysis
* Price range classification
* Product performance analysis
* Interactive dashboard creation using **Power BI**

---

# 📂 Repository Structure

```text
excel-data-cleaning-transformation/
│
├── Assignment 2 - Data Cleaning and Transformation.xlsx
│
├── README.md
│
└── screenshots/
    ├── raw-data.png
    ├── cleaned-data.png
    ├── inconsistent-data.png
    ├── splitting-merging.png
    ├── number-formatting.png
    └── conditional-formatting.png
```

---

# 👨‍💻 About Me

I am an **Aspiring Data Analyst** building my professional portfolio through practical projects in **Microsoft Excel, Power BI, SQL, and Data Analysis**.

This project demonstrates my ability to work with a raw dataset, identify data quality issues, apply appropriate cleaning and transformation techniques, and prepare data for meaningful analysis.

---

## 🧰 Skills Demonstrated

**Microsoft Excel | Data Cleaning | Data Transformation | Data Preprocessing | Missing Value Handling | Find & Replace | Remove Duplicates | Text-to-Columns | Excel Formulas | Conditional Formatting | Data Formatting | Data Analysis**

---

## 📌 Project Status

**Completed ✅**

This project is part of my **Data Analyst Portfolio** and demonstrates practical experience in cleaning and preparing real-world style datasets using Microsoft Excel.
