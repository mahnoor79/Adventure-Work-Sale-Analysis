# Adventure-Work-Sale-Analysis
"In this project, I will share a comprehensive sales analysis project using Excel on Adventure Works dataset."
Tool Used: Microsoft Excel

# 👋 Warm Welcome
Hi there!
Before diving deep into this powerful Excel-based sales analysis project, let's take a moment to appreciate the journey. This backend work is not just about formulas and features — it's about learning how raw data transforms into meaningful insights. Whether you're a data enthusiast or a student stepping into the world of analytics, this guide walks you through every step of the backend process that powers a dynamic sales dashboard.

# 📁 Contents
1) Data Files Overview
2) Data Loading and Importing
3) Power Query Cleaning Techniques
4) Column Transformations
5) Custom & Conditional Columns
6) KPIs Creation (Power Query + Power Pivot)
7) Merge, Replace, Remove Operations
8) Relationships & Data Model
9) Developer Tools & Interactivity
10) Number Formatting & Graph Refinement

# Key Learnings

**1) 📂 Data Files**
The dataset includes 6 Excel files related to internet sales, products, customers, sales territories, geographies, and dates.

**2) 🔄 Data Loading**

**Method**

Open Excel → Go to Data → Click Get Data
Select:
From File
From Excel Workbook
Steps
After selecting "From Excel Workbook", a Navigator window appears.
Choose required sheets by enabling "Multiple" selections.
Click Load to → Choose the desired loading destination.

**3) 🧹 Data Cleaning (Power Query Editor)**
1) Fact Internet Sales Sheet
Select only important columns using "Choose Column"

**2) Remove unnecessary data for optimized processing**

a) Dim Sales Territory Sheet
b) Remove NA and null values

**3) Keep only clean and relevant columns**

a) Dim Product Sheet
b) Keep essential columns
c) Remove duplicates

**4) Replace unwanted values in the "Color" column**

a) Dim Geography Sheet
Choose only necessary columns

b) Dim Date Sheet
Remove unrelated columns

**5) Add new columns using "Add Column" → Date**

Year, Month Number, Month Name, Day Name

6) Create Conditional Columns

7) Change data types to Text or Date as needed

a) Dim Customer Sheet
b) Use the Merge Columns feature for full names
c) Retain only relevant data fields

# 🧮 KPIs Preparation (Power Query Editor)

1) Total Revenue = Order Quantity * Unit Price
2) Cost of Delivered Sales (CODS) = Order Quantity * Cost
3) Total Profit = Total Revenue - CODS
# 🧮 KPIs Preparation (Power Pivot - DAX Measures)

1) Transaction Count = =COUNTROWS(FactInternetSales)
2) All Products = =COUNTROWS(DimProduct)
3) Sold Products = =DISTINCTCOUNT(FactInternetSales[ProductKey])
4) Unsold Products = [All Products] - [Sold Products]
5) Profit Margin % = =DIVIDE([Total Profit], [Total Revenue], 0)

# 🔗 Manage Relationships

1) Once all sheets are cleaned and loaded:
2) Go to Power Pivot → Click Manage
3) Define relationships between Fact and Dimension tables
4) Open Diagram View to visually map relationships

# 🧰 Advanced Excel Tools (Developer Tab)

1) Insert interactive controls like Option Buttons
2) Link sheets via Format Control → Cell Link
3) Enable navigation and control via macros and slicers

# 🎨 Graph and Axis Formatting

1) Common Graph Enhancements:
Month Name sorted by Month Number

2) Proper number formatting using custom syntax:

[<999999]0.00," K";[<999999999]0.00 " M";0.00 " B"

Use design assets from AdobeColor.com and Flaticon.com

Customize slicers, captions, and dashboard elements for interactivity

# 🧠 My Learnings

1) How to import and load multiple Excel files
2) Data cleaning: Remove NA, nulls, and duplicates
3) Column selection and transformation
4) Creating KPIs using formulas and DAX
5) Power Query vs Power Pivot functions
6) Formatting numbers in K/M/B units
7) Linking sheets via developer tools
8) Creating a connected, filterable, user-friendly dashboard
9) Understanding slicers, macros, and interactive elements
10) Graph axis formatting and sorting for time-based visuals
11) Use of logical, lookup, and conditional formulas in dashboards

# 🙌 Final Words
This backend journey is more than just steps — it’s the foundation of meaningful data storytelling. By following each step with precision and creativity, you’re not only building dashboards, you’re crafting insights.

**Thanks for reading!**
Stay tuned for more advanced tips and tutorials.

**Mahnoor Naseer**
📧 mahnoornoorg57@gmail.com
