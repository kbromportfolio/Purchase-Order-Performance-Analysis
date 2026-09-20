# Purchase Order Performance & Supplier Analysis

## Stage 1 — Project Planning

Stage 1 of the project has been completed.

### Completed

* Defined the project aim and objectives.
* Confirmed the project scope and deliverables.
* Confirmed the dataset structure and project tools.
* Created the project timeline.
* Created the project management tracker.
* Created the local project directory and GitHub repository structure.

### Project Tools

* Microsoft Excel
* Git
* GitHub

The detailed **Project Proposal**, **Project Timeline**, and **Project Management Tracker** are provided as PDF documents in the repository.

**Status:** Stage 1 Completed

## Stage 2 – Dataset Preparation

I completed the dataset preparation stage using the 777-row purchase-order dataset.

The dataset contains 11 original fields covering purchase orders, suppliers, dates, categories, order status, quantity, pricing, defective units, and compliance.

I preserved the original data as raw data, imported the dataset into Excel, checked the structure and column headings, and prepared the working file for the next stage.

**Status:** Completed

**Next Stage:** Data Cleaning

## Stage 3 – Data Cleaning

The data-cleaning assessment was completed for the 777-record, 11-field procurement dataset.

### Cleaning actions and decisions

* Dates were checked and standardised to the `DD/MM/YYYY` format. `Order_Date` ranges from 01/01/2022 to 01/01/2024, and `Delivery_Date` ranges from 01/06/2022 to 01/12/2024.
* Text and category values were reviewed for consistency. The dataset contains 5 suppliers, 5 item categories, 4 order statuses, and 2 compliance values, with no duplicate category labels identified.
* Numeric fields were checked and classified as integer or decimal values: `Quantity`, `Unit_Price`, `Negotiated_Price`, and `Defective_Units`.
* No exact duplicate records were identified. `PO_ID` contains 777 unique purchase order IDs.
* `Delivery_Date` had 87 missing values (11.20%). These were filled using the order date plus the median observed delivery lag for the relevant order status: 13 days for `Cancelled`, 11 days for `Delivered`, 11 days for `Partially Delivered`, and 13 days for `Pending`.
* `Defective_Units` had 136 missing values (17.50%). These were filled with `0`, treating a blank as no recorded defective units for this personal project analysis.

The cleaned dataset is saved in `data/cleaned/Procurement KPI Analysis Dataset - Cleaned.csv`. The original raw data remains unchanged.

Completed Stage 3 data cleaning using Excel.

- Standardised date formats
- Reviewed text and category consistency
- Checked numeric fields
- Confirmed no duplicate records
- Filled missing Delivery_Date values
- Filled missing Defective_Units values with 0
- Preserved the original raw dataset

## Stage 4 – Data Validation

Data validation was completed in Excel using formulas, filters, and duplicate checks. The checks cover:

- `PO_ID` format and duplicate IDs
- `Order_Date` and `Delivery_Date` format and chronology
- Positive integer `Quantity`
- Positive `Unit_Price` and `Negotiated_Price`, with negotiated price not above unit price
- Non-negative integer `Defective_Units`, not above `Quantity`
- `Compliance` values restricted to `Yes` or `No`

The validation checked 777 rows and confirmed 777 unique PO IDs. All requested format and value checks passed except for one chronology issue: `PO-00101` has `Order_Date` `27/02/2022` and `Delivery_Date` `22/02/2022`. This record should be reviewed before the dataset is treated as fully validated.

## Stage 5 – Formula Development

Formula development was completed in Microsoft Excel using the cleaned and validated procurement dataset.

The following calculated fields were added to the workbook:

- `Original Cost` = `Quantity × Unit_Price`
- `Negotiated Cost` = `Quantity × Negotiated_Price`
- `Potential Savings` = `Original Cost − Negotiated Cost`
- `Discount %` = `Potential Savings ÷ Original Cost`
- `Defect Rate %` = `Defective_Units ÷ Quantity`
- `Delivery Days` = `Delivery_Date − Order_Date`

The formulas were applied to all 777 purchase-order records. Currency, percentage, and whole-number formats were applied to the calculated fields for clear analysis.

The completed Excel workbook is saved as `Excel/Procurement KPI Analysis Dataset.xlsx`.

**Status:** Stage 5 Completed

**Next Stage:** Supplier and Category Analysis

## Stage 6 – Supplier and Category Analysis

Supplier and category analysis was completed in Microsoft Excel using the calculated fields from Stage 5.

Two analysis sheets were added to the workbook:

- `Supplier Analysis` summarises each supplier.
- `Category Analysis` summarises each item category.

Both sheets analyse:

- Purchase-order count
- Quantity ordered
- Original spend
- Negotiated spend
- Potential savings
- Defective units
- Defect rate %
- Compliant purchase orders
- Compliance rate %

The analysis tables use Excel formulas linked to the source dataset, so the results update when the source data changes.

**Status:** Stage 6 Completed

**Next Stage:** PivotTable Analysis

## Stage 7 – PivotTable Analysis

PivotTable and PivotChart analysis was completed in Microsoft Excel using the calculated fields from Stage 5 and the cleaned procurement dataset.

The `PivotTable Analysis` sheet contains five native PivotTables and five PivotCharts covering:

- Supplier performance and spend
- Category performance and savings
- Order status and spend
- Supplier cost comparison
- Category quality and compliance

The PivotTables use the 777 purchase-order records as their source and can be refreshed in Excel when the source data changes.

**Status:** Stage 7 Completed

## Stage 8 – KPI Development

Core procurement and supplier-performance KPIs were calculated in Microsoft Excel using formulas linked to the source dataset.

The `KPI Development` sheet includes:

- Total purchase orders and quantity ordered
- Original spend, negotiated spend, potential savings, and savings rate
- Average discount and average delivery days
- Overall defect rate, compliance rate, and delivered-order rate
- Supplier-level purchase orders, spend, savings rate, delivery days, defect rate, and compliance rate

The workbook is configured to recalculate the KPI formulas when the source data changes.

**Status:** Stage 8 Completed



