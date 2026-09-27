Data Cleaning and Transformation -- Assignment 2

📌 Project Overview

This project is a beginner-level Data Cleaning and Transformation
exercise using Microsoft Excel.

The objective is to take a product dataset and prepare it for analysis
by identifying missing or inconsistent values, standardizing text,
removing duplicates, formatting columns, splitting and merging
information, and applying conditional formatting.

📂 Dataset

The dataset contains product-related information such as:

Column                              Description

Product ID                          Unique product identifier
containing date/day and country
information

Manufacturing Date                  Date associated with the product

Country Code                        Country identifier

Product Brand                       Product and brand description

Product Name                        Name/type of the product

Brand Name                          Product brand

Price ($)                          Product price

Quantity                            Available quantity

Category                            Product category

🎯 Objectives

The main objectives of this assignment are:

Handle missing values.

Correct inconsistent text and category values.

Remove duplicate records.

Apply appropriate number and date formatting.

Split and merge columns.

Apply conditional formatting for easier data interpretation.

Prepare a clean dataset suitable for further analysis.

🧹 Data Cleaning Tasks

1. Handling Missing Values

The dataset should be checked for missing values, especially in the
Price and Category columns.

Possible approaches:

For missing numerical values such as Price, use an appropriate
statistical value such as the average/median when justified.

For missing categorical values, use a suitable category based on
available information or assign Unknown when the value cannot be
determined.

Avoid replacing values blindly; the chosen method should be
documented.

Example Excel formula:

=IF(ISBLANK(G2),AVERAGE($G$2:$G$35),G2)

For categorical values, an example approach is:

=IF(ISBLANK(I2),"Unknown",I2)

Note: In this dataset, some category values are already
represented as Unknown. These should be reviewed as part of the
data-quality check.

2. Correcting Inconsistent Data

The Product Name and Category columns should be reviewed for:

Extra spaces

Inconsistent capitalization

Unwanted characters

Typographical errors

Different text formats representing the same value

A useful Excel formula for cleaning product names is:

=PROPER(TRIM(CLEAN(E2)))

The Find and Replace feature can also be used to standardize known
typos or inconsistent category names.

3. Removing Duplicates

The dataset should be checked for duplicate rows based on the complete
record.

Steps in Excel:

Select the complete dataset.

Go to Data → Remove Duplicates.

Select all relevant columns.

Review the duplicate records.

Remove duplicates where appropriate.

This ensures that the same product record is not counted multiple times
during analysis.

4. Splitting and Merging Data

Split Product ID

The Product ID contains useful information that can be separated
into individual fields.

For example:

28-JAN-US

The assignment asks to extract information from the Product ID using
formulas such as:

=LEFT(A2,2)

and

=RIGHT(A2,2)

This can be used to create separate fields for the required day/date
component and Country Code.

Merge Brand Name and Product Name

The Brand Name and Product Name columns can be combined into a
new column named Product Brand.

Example:

=CONCAT(F2," ",E2)

Result:

Dell Laptop
Nike Sneakers
Samsung Smartphone

5. Number and Date Formatting

Price

The Price ($) column should be formatted as currency.

Example:

1000 → $1,000.00

Manufacturing Date

The Manufacturing Date column should use the following display
format:

DD-MM-YYYY

Example:

28-01-2026

This improves readability and consistency when the data is used for
analysis.

6. Conditional Formatting

Conditional formatting is used to make important values easier to
identify visually.

Price Column

Apply a Data Bar or Color Scale to the Price column.

This allows high and low product prices to be identified quickly.

Category Column

Create a custom conditional formatting rule to highlight records where:

Category = Electronics

This makes electronic products easier to identify within the dataset.

🔍 Data Quality Checks

Before considering the dataset ready for analysis, check:

Missing Price values

Missing or Unknown Category values

Duplicate rows

Extra spaces in text fields

Inconsistent capitalization

Category spelling/typing issues

Correct Price currency format

Correct Manufacturing Date format

Correct Country Code values

Correct Product Brand values after merging

🛠️ Tools Used

Microsoft Excel

Excel formulas

Find and Replace

Remove Duplicates

Number Formatting

Conditional Formatting

Text cleaning functions such as TRIM, CLEAN, and PROPER

📈 Expected Outcome

After completing the cleaning and transformation process, the dataset
should be:

Clean and consistent

Free from unnecessary duplicate records

Easier to read and understand

Properly formatted

Ready for further data analysis or visualization

📚 Key Skills Demonstrated

This assignment demonstrates beginner-level Data Analyst skills in:

Data cleaning

Data quality checking

Missing-value handling

Text transformation

Duplicate removal

Data formatting

Excel formulas

Conditional formatting

Data preparation
