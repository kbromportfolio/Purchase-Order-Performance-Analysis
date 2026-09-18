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

