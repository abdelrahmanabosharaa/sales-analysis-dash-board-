# Sales Analysis Project (Excel, Power Query & Power Pivot)

An end-to-end data analysis project built entirely in Microsoft Excel. It starts from raw monthly order files and ends with an interactive dashboard showing sales by region, category, product, and customer.

## Table of Contents
- [Project Overview](#project-overview)
- [Tools Used](#tools-used)
- [Project Workflow](#project-workflow)
  - [Stage 1: Data Exploration in Excel](#stage-1-data-exploration-in-excel)
  - [Stage 2: Data Cleaning and Transformation (Power Query)](#stage-2-data-cleaning-and-transformation-power-query)
  - [Stage 3: Data Model / Schema (Power Pivot)](#stage-3-data-model--schema-power-pivot)
  - [Stage 4: Analysis with PivotTables](#stage-4-analysis-with-pivottables)
  - [Stage 5: Dashboard](#stage-5-dashboard)
- [Key Insights](#key-insights)
- [Repository Structure](#repository-structure)
- [How to Use](#how-to-use)
- [Author](#author)

---

## Project Overview
The goal of this project is to turn raw, messy order data into clear business insights. The data comes as **multiple monthly order files** (one file per month) plus two lookup tables: **Products** and **Customers**.

Questions this project answers:
- Which regions generate the most sales?
- Which product categories bring in the most revenue?
- Which products sell the most units?
- Who are the top 10 customers by revenue?

## Tools Used
| Tool | Purpose |
|------|---------|
| Microsoft Excel | Data exploration, PivotTables, dashboard |
| Power Query (M language) | Combining and cleaning the data |
| Power Pivot (Data Model, DAX) | Building relationships between tables |
| PivotTables and PivotCharts | Analysis and visuals |

---

## Project Workflow

### Stage 1: Data Exploration in Excel
Before changing anything, the raw data was reviewed to understand its structure and quality.

**Data sources**
| Table | Columns | Description |
|-------|---------|-------------|
| Orders | OrderID, CustomerID, ProductID, OrderDate, Quantity | One file per month, combined later |
| Products | ProductID, ProductName, Category, Price | Product catalog |
| Customers | CustomerID, First Name, Second Name, Region | Customer list |

**Problems found**
- Duplicate orders (repeated OrderID)
- Missing values in `Quantity` and `ProductID`
- Inconsistent date formats (different separators such as `/` and `-`, and different orders of day, month, and year)
- Month information stored only in the file name

---

### Stage 2: Data Cleaning and Transformation (Power Query)
All order files in one folder were combined and cleaned with Power Query. The full M code is in [`M_code_Data_Cleaning_steps.txt`](./M_code_Data_Cleaning_steps.txt).

| # | Step | What it does |
|---|------|--------------|
| 1 | Load from folder | `Folder.Files` reads every monthly order file |
| 2 | Filter hidden files | Skips hidden or system files |
| 3 | Combine files | A custom function transforms each file, then the tables are expanded into one |
| 4 | Set data types | Text for names, whole numbers for IDs and Quantity |
| 5 | Remove duplicates | Keeps one row per `OrderID` |
| 6 | Remove incomplete rows | Drops rows with empty `Quantity` or `ProductID` |
| 7 | Standardize dates | Replaces `/` with `-`, then splits the date into 3 parts |
| 8 | Detect the Year | Takes the part that is 2000 or greater |
| 9 | Extract the Month | Reads the month number from the file name (for example `Orders_3.csv` gives 3) |
| 10 | Detect the Day | Takes the remaining part that is 31 or less and is not the month |
| 11 | Build a clean date | Merges Day, Month, and Year into one `New Date` column of type Date |
| 12 | Final tidy-up | Removes helper columns, adds a new sequential `OrderID`, and reorders columns |

**Final Orders table:** `OrderID | CustomerID | ProductID | Quantity | Date`

---

### Stage 3: Data Model / Schema (Power Pivot)
The cleaned tables were loaded into the Excel Data Model and linked in **Diagram View** to form a **star schema**.

```
Products (1) ────────< Orders (*) >──────── (1) Customers
 ProductID               OrderID               CustomerID
 ProductName             CustomerID            First Name
 Category                ProductID             Second Name
 Price                   Date                  Region
                         Quantity
```

| Relationship | Type | Key |
|--------------|------|-----|
| Products to Orders | One-to-many | `ProductID` |
| Customers to Orders | One-to-many | `CustomerID` |

**Sales formula**
```
Sales = Orders[Quantity] * Products[Price]
```
Quantity comes from the Orders sheet and Price from the Products sheet.

> Screenshot: add your Power Pivot diagram here, for example `images/data_model.png`
> `![Data Model](images/data_model.png)`

---

### Stage 4: Analysis with PivotTables
PivotTables built on the Data Model were used to summarize the data:

| PivotTable | Rows | Values |
|------------|------|--------|
| Sales by Region | Region | Sum of Sales |
| Sales by Category | Category | Sum of Sales |
| Units by Product | ProductName | Sum of Quantity |
| Region Sales and Quantity | Region | Sum of Sales, Sum of Quantity |

**Overall totals**
- Total sales: **$101,111,318**
- Total quantity sold: **108,087 units**

> Screenshot: add your PivotTable image here, for example `images/pivot_tables.png`
> `![PivotTables](images/pivot_tables.png)`

---

### Stage 5: Dashboard
An Excel **Dashboard** sheet brings the analysis together with charts and a ranking table.

| Visual | Type | Shows |
|--------|------|-------|
| Region Sales | Bar chart | Sales by region |
| Category Sales $ | Column chart | Sales by product category |
| Top 5 Product Sales Units | Column chart | Best-selling products by quantity |
| Top 10 Customers by Revenue | Table | Customer, region, and total sales |

> Screenshot: add your dashboard image here, for example `images/dashboard.png`
> `![Dashboard](images/dashboard.png)`

---

## Key Insights

**Sales by region**
| Region | Sales |
|--------|-------|
| Mansoura | $21,015,212.93 |
| Aswan | $20,569,062.91 |
| Cairo | $20,452,188.95 |
| Giza | $20,412,092.82 |
| Alexandria | $18,662,760.43 |

**Sales by category**
| Category | Sales |
|----------|-------|
| Clothing | $29,101,756 |
| Toys | $25,377,189 |
| Electronics | $19,012,040 |
| Books | $17,185,365 |
| Furniture | $10,434,969 |

**Highlights**
- **Mansoura** is the top region, while **Alexandria** is the lowest.
- **Clothing** and **Toys** together make up more than half of total revenue.
- **Furniture** is the weakest category.
- Top 5 products by units sold: Tablet (4,634), Table (4,620), Ball (4,514), Camera (4,504), Jacket (4,493).
- The top customer is **Ahmed Ibrahim** (Alexandria) with 520,408 in sales, and 3 of the top 10 customers are from Alexandria.

---

## Repository Structure
```
.
├── README.md
├── Project.xlsx                      # Final workbook (data model, pivots, dashboard)
├── M_code_Data_Cleaning_steps.txt    # Power Query M code
├── Orders/                           # Raw monthly order files
├── Products.xlsx                     # Products lookup table
├── Customers.xlsx                    # Customers lookup table
└── images/
    ├── data_model.png
    ├── pivot_tables.png
    └── dashboard.png
```
Adjust the names to match your actual files.

## How to Use
1. Clone or download this repository.
2. Open `Project.xlsx` in Excel (Excel 2016 or later, or Microsoft 365, with Power Pivot enabled).
3. If you want to refresh the data, update the folder path in Power Query:
   **Data > Get Data > Query Options**, or edit the `Source` step in the Advanced Editor.
4. Click **Data > Refresh All** to update the tables, PivotTables, and dashboard.

## Author
**<Your Name>**
- GitHub: [@your-username](https://github.com/your-username)
- LinkedIn: <your link>

---
*Built as part of a data analysis course project using Excel, Power Query, and Power Pivot.*
