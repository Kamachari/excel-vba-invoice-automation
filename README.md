# Excel VBA – Invoice Data Automation & Management System

An Excel VBA-based project that automates invoice record management through conditional data mapping, category-based classification, duplicate detection, and invoice search functionality. The system uses predefined business rules to organize records into structured worksheets and simplify invoice tracking and retrieval.

## 1. Demo Link

- Demo: Not publicly hosted.
- Screenshots: Available in the `Screenshots` folder.
- Source Code: Available in the repository.

## 2. Table of Contents

- Business Understanding
- Data Understanding
- Screenshots of Results
- Technologies
- Setup
- Approach
- Business Rules & Conditions
- Key Functions
- Project Status
- Credits

## 3. Business Understanding

Manual invoice management can lead to duplicate records, inconsistent data entry, and difficulties retrieving previously stored information.

The objective of this project is to develop an Excel VBA automation solution that simplifies invoice data processing through predefined business rules and conditional workflows.

The system automates:
- Invoice classification based on invoice type and category.
- Conditional data mapping between input fields and destination worksheets.
- Duplicate identification using unique reference numbers.
- Sequential record numbering and structured data storage.
- Invoice searching across multiple worksheets.
- Automatic navigation to matching invoice records.

The main challenge was implementing different data-mapping rules for different invoice categories while maintaining a common validation and retrieval process.

## 4. Data Understanding

The project uses a structured Excel workbook with an input interface and multiple destination worksheets.

**Workbook Structure:**

| Worksheet | Purpose |
|---|---|
| INPUT | Accepts invoice details and provides search functionality. |
| PO | Stores Purchase Order records. |
| NON PO | Stores Non-Purchase Order records. |
| TAXI | Stores travel-related expense records. |

The system processes user-entered information rather than relying on an external dataset.

Each record contains relevant information such as reference number, invoice date, vendor details, bill number, and bill amount.

## 5. Screenshots of Visualizations/Results

The following screenshots demonstrate the system's interface and automation results.

| Input Interface | Automated Record Storage |
|:---:|:---:|
| ![Input Interface](Screenshots/Input%20Interface.png) | ![Automated Record Storage](Screenshots/Automated%20Record%20Storage.png) |

| Duplicate Detection | Invoice Search Results |
|:---:|:---:|
| ![Duplicate Detection](Screenshots/Duplicate%20Detection.png) | ![Invoice Search Results](Screenshots/Invoice%20Search%20Results.png) |

Refer to the `Screenshots` folder for visual demonstrations.

## 6. Technologies

- Microsoft Excel
- Excel VBA (Visual Basic for Applications)
- Conditional Statements (If–Else)
- Worksheet and Range Operations
- Excel Find Method
- Data Validation
- String Manipulation
- Automated Data Mapping

## 7. Setup

1. Clone or download this repository.
2. Open the Excel workbook using Microsoft Excel.
3. Enable macros when prompted, provided the workbook is from a trusted source.
4. Ensure the workbook contains the required worksheets: `INPUT`, `PO`, `NON PO`, and `TAXI`.
5. Add the VBA modules to the workbook if they are not already included.
6. Enter invoice information through the input interface.
7. Use the available macro buttons to add, search, and open records.

**Note:** The workbook must be saved in `.xlsm` format to retain VBA macros.

## 8. Approach

The project follows a rule-based automation workflow:

1. **Input Collection:** Collect invoice information through a structured input worksheet.
2. **Classification:** Identify the invoice type and applicable category.
3. **Conditional Routing:** Select the appropriate destination worksheet based on predefined conditions.
4. **Duplicate Validation:** Search for an existing reference number across all destination worksheets.
5. **Data Mapping:** Transfer the relevant input fields to their corresponding destination columns.
6. **Record Generation:** Identify the next available row and generate a sequential serial number.
7. **Data Formatting:** Apply standardized date formatting to designated columns.
8. **Record Retrieval:** Search across all destination worksheets using the unique reference number.
9. **Record Navigation:** Activate the matching worksheet and highlight the corresponding record.

## 9. Business Rules & Conditions

The system implements the following rules:

**Invoice Classification**
- If invoice type is `PO`, store the record in the PO worksheet.
- If invoice type is `NON PO` and category is `TRAVELLING EXPENSES`, route the record to the TAXI worksheet.
- If invoice type is `NON PO` with another category, store it in the NON PO worksheet.
- Invalid invoice types are rejected.

**Duplicate Prevention**
- Every record must have a unique CID number.
- Before inserting a record, the system checks the CID number across all three destination worksheets.
- Existing CID numbers are rejected to prevent duplicate entries.

**Data Mapping**
- Different input fields are mapped depending on the invoice classification.
- Travel-related records receive a predefined TAXI classification.
- When a comma-separated value is encountered in the designated input field, only the first value is extracted.

**Record Management**
- New records are inserted into the next available row.
- Serial numbers are generated automatically.
- Date columns are formatted consistently.

**Search & Retrieval**
- Empty search requests are rejected.
- The system searches all three destination worksheets using an exact CID match.
- Matching invoice details are displayed in the input worksheet.
- Record navigation is allowed only after a successful search.
- The matching worksheet and complete record are automatically selected.

## 10. Key VBA Functions

| Function / Procedure | Responsibility |
|---|---|
| `AddInvoice()` | Classifies, validates, maps, and stores invoice records. |
| `CIDExists()` | Checks whether a CID exists in a worksheet. |
| `GetCIDSheet()` | Identifies the worksheet containing an existing CID. |
| `SearchInvoice()` | Searches records and displays matching invoice details. |
| `OpenInvoice()` | Navigates to and highlights the matching record. |
| `ClearSearchResults()` | Clears previously displayed search results. |

## 11. Project Status

**Status:** Completed

The core automation features, including conditional classification, duplicate prevention, data mapping, invoice searching, and record navigation, have been implemented using Excel VBA.

## 12. Credits

Developed independently as a practical Excel VBA automation project.

