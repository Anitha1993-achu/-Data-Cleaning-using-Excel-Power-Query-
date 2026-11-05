# Product Dataset – Data Cleaning and Preparation Using Excel & Power Query

## 1. Project Overview

This project focuses on cleaning, transforming, and preparing a Product Dataset using **Microsoft Excel and Power Query Editor**.

Real-world datasets often contain missing values, duplicate records, inconsistent text formats, spelling mistakes, and poorly structured columns.

The main objective of this project is to improve data quality, accuracy, consistency, and readability by applying Excel data-cleaning techniques and Power Query transformations.

## 2. Dataset Description

**Dataset Name:** Product Dataset

**Attributes:**

| Column Name | Description |
|---|---|
| Product ID | Unique product identification |
| Product Name | Name of the product |
| Brand Name | Product brand |
| Quantity | Available product quantity |
| Category | Product category |
| Price | Product price |

## 3. Tools and Technologies

- Microsoft Excel
- Power Query Editor
- Excel Data Cleaning Tools
- Find and Replace
- Remove Duplicates
- Text to Columns
- Merge Columns
- Conditional Formatting
- Number Formatting
- Power Query Transform Tools

## 4. Data Cleaning and Preparation Tasks

### Task 1: Handling Missing Values

identify and handle missing values in the Price and Category
 columns using Excel formulas.

•	Identify blank cells in the Price column using Filter. 

•	Create a new column named Cleaned Price.

•	Apply the IF formula:
  = (IF (ISBLANK ([@[Price ($)]]), AVERAGE ([Price ($)]), [@[Price ($)])
  
•	Identify blank cells in the Category column.

•	Create a new column named Cleaned Category.

•	Use the IF function to assign categories based on product names.

•	=IF(ISBLANK([@Category]),"Electronics",[@Category])



### Task 2: Correcting Inconsistent Data

**A. Product Name**

Examples of inconsistent product names:

| Original Product Name | Cleaned Product Name |
|---|---|
| laptop | Laptop |
| LAPTOP | Laptop |
| sneakers | Sneakers |
| HEADPHONES | Headphones |

**Using Excel:**
1. Select the Product Name column.
2. Press Ctrl + H.
3. Enter the incorrect text in Find what.
4. Enter the corrected text in Replace with.
5. Click Replace All.

**Using Power Query:**
1. Select Product Name.
2. Go to Transform → Format → Capitalize Each Word.
3. Use Replace Values to correct specific inconsistencies.
4. Apply Trim and Clean to remove unwanted spaces and characters.

**B. Category Corrections**

Examples:

| Original Category | Corrected Category |
|---|---|
| Electroni | Electronics |
| Electrnics | Electronics |
| Fashon | Fashion |
| Home Appliences | Home Appliances |

**Using Excel:**
- Press Ctrl + H.
- Replace incorrect category names with correct names.

**Using Power Query:**
- Select Category.
- Go to Transform → Replace Values.
- Replace spelling mistakes with the correct category.
- Apply Trim and Clean to standardize the text.

### Task 3: Removing Duplicate Records

**Using Excel:**
1. Select the entire dataset.
2. Go to Data → Remove Duplicates.
3. Select all columns.
4. Click OK.
5. Excel removes completely duplicated rows.

**Using Power Query:**
1. Open Power Query Editor.
2. Select all columns.
3. Go to Home → Remove Rows → Remove Duplicates.
4. Verify that duplicate records have been removed.

**Expected Result:** Only unique complete records remain.

### Task 4: Splitting and Merging Data

**A. Splitting Product ID**

Example Product ID:

`28-JAN-US`

The Product ID is separated into:

| Product ID | Date Part | Country Code |
|---|---|---|
| 28-JAN-US | 28-JAN | US |
| 15-FEB-US | 15-FEB | US |
| 07-JUN-UK | 07-JUN | UK |

**Using Excel:**
1. Select Product ID.
2. Go to Data → Text to Columns.
3. Select Delimited.
4. Choose Other and enter `-`.
5. Click Finish.
6. Combine the date-related parts if required and keep the country code separately.

**Using Power Query:**
1. Select Product ID.
2. Go to Transform → Split Column → By Delimiter.
3. Select Hyphen (-) as the delimiter.
4. Choose Each occurrence of the delimiter.
5. Rename the resulting columns.
6. Combine the date-related parts into Manufacturing Date.

**Important:** The example Product ID contains only a day and month, not a year. Obtain the correct manufacturing year from the source before creating a complete Manufacturing Date. Do not assume a year.

**B. Merging Brand Name and Product Name**

The Brand Name and Product Name columns are combined into a new column named Product Brand.

**Using Excel:**

Formula:

`=CONCAT(C2," ",B2)`

**Using Power Query:**
1. Select Brand Name and Product Name.
2. Go to Transform → Merge Columns.
3. Choose Space as the separator.
4. Set the new column name as Product Brand.
5. Click OK.

Example:

| Brand Name | Product Name | Product Brand |
|---|---|---|
| Dell | Laptop | Dell Laptop |
| Nike | Sneakers | Nike Sneakers |
| Sony | Headphones | Sony Headphones |

### Task 5: Number Formatting

**A. Price Column – Currency Format**

**Using Excel:**
1. Select Price.
2. Go to Home → Number Format.
3. Select Currency.
4. Choose the appropriate currency symbol.
5. Set the required decimal places.

**Using Power Query:**
1. Select Price.
2. Go to Transform → Data Type.
3. Choose Decimal Number or Fixed Decimal Number.
4. Load the cleaned data into Excel.
5. Apply Currency formatting in Excel.

**B. Manufacturing Date – DD-MM-YYYY**

**Using Excel:**
1. Select Manufacturing Date.
2. Press Ctrl + 1.
3. Select Custom.
4. Enter `dd-mm-yyyy`.
5. Click OK.

**Using Power Query:**
1. Select Manufacturing Date.
2. Set its data type to Date.
3. Load the data into Excel.
4. Apply the custom display format `dd-mm-yyyy`.

### Task 6: Conditional Formatting

**A. Price Column – Data Bars or Color Scales**

**Using Excel:**
1. Select the Price column.
2. Go to Home → Conditional Formatting.
3. Choose Data Bars or Color Scales.
4. Select a suitable formatting style.
5. Compare product prices visually.

**Using Power Query:**

Power Query does not directly support Excel conditional formatting. Load the transformed data into Excel and apply conditional formatting there.

**B. Highlight Electronics Category**

1. Select the Category column.
2. Go to Home → Conditional Formatting.
3. Select Highlight Cells Rules → Text That Contains.
4. Enter `Electronics`.
5. Choose a highlight color.
6. Click OK.

**Expected Result:** All Electronics category cells are highlighted.

## 5. Final Dataset Structure

The cleaned dataset will contain the following columns:

| Column Name | Cleaning / Transformation |
|---|---|
| Product ID | Original product identifier |
| Manufacturing Date | Extracted and converted to Date |
| Country Code | Extracted from Product ID |
| Product Brand | Merged Brand Name and Product Name |
| Quantity | Numeric data type |
| Category | Corrected and standardized |
| Price | Decimal number with currency formatting in Excel |

## 6. Final Outcome

After completing the cleaning and preparation process:

- Identified and handled missing Price values.
- Identified and handled missing Category values.
- Standardized inconsistent Product Names.
- Corrected spelling mistakes in Category.
- Removed completely duplicated records.
- Split Product ID into Manufacturing Date and Country Code.
- Merged Brand Name and Product Name into Product Brand.
- Applied currency formatting to Price.
- Applied DD-MM-YYYY formatting to Manufacturing Date.
- Applied Data Bars or Color Scales to Price.
- Highlighted Electronics using Conditional Formatting.

## 7. Conclusion

This project demonstrates practical data-cleaning and data-preparation techniques using Microsoft Excel and Power Query Editor.

Excel is useful for manual cleaning, formatting, and conditional formatting, while Power Query provides efficient and repeatable data transformation operations.

The final cleaned dataset is more accurate, consistent, and organized, making it suitable for further data analysis, reporting, and visualization.

## 8. Skills Demonstrated

- Data Cleaning
- Data Preprocessing
- Missing Value Handling
- Data Standardization
- Duplicate Removal
- Column Splitting and Merging
- Power Query Transformations
- Excel Formulas
- Conditional Formatting
- Data Formatting

## Author

**Role:** Anitha Data Analyst

**Tools:** Microsoft Excel | Power Query Editor

**Project:** Product Dataset – Data Cleaning and Preparation


