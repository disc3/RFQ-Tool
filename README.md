# RFQ Tool - Project Documentation & Developer Guidelines

## 1. Project Overview
The **RFQ Tool** is an advanced Excel VBA application designed to streamline the Request for Quotation (RFQ) process for manufacturing. It bridges the gap between Sales and Engineering by automating Costing, Bill of Materials (BOM) definition, Routing creation, and SAP ERP integration.

**Primary Goals:**
* **Standardization:** Ensure consistent costing logic across different plants.
* **Automation:** Reduce manual entry via SAP Scripting and Power Automate integration.
* **Speed:** Enable mass creation of Variants and rapid BOM construction.

---

## 2. Workflow & Architecture

The application relies heavily on **Excel ListObjects (Tables)** for data storage and manipulation. It avoids hardcoded cell references wherever possible.

### Core Workflow:
1.  **Product Creation:** User defines a Final Product (`AddProductForm`).
2.  **BOM Definition:** User searches the local database (`Purchasing Info Records`) to add components (`ResultsForm`) or performs a bulk upload (`MassCopy`).
3.  **Routing Definition:** User assigns manufacturing operations based on the selected plant's capabilities (`RoutineForm`).
4.  **Variant Generation:** (Optional) User creates variations of a base product (`VariantCreationForm`, `frmRoutineVariantEditor`).
5.  **Validation:** The system checks for missing prices, quantities, or operations (`RFQValidation`).
6.  **Export/Integration:**
    * **Power Automate:** JSON payload sent to Azure Logic Apps for CRM integration.
    * **SAP:** Direct scripting to creating BOMs (`CS01`), Routings (`CA01`), and check stock (`CO09`).

---

## 3. Key Modules & Functions

### A. Core Logic
* **`Search_copy_components.bas`**: The central engine for BOM manipulation.
    * `AddFullComponent`: Composite sub that adds a row and performs `XLookup` against the database.
    * `UpdateComponentDetails`: Refreshes cost/description data for existing BOM lines.
    * `RefreshBOMData`: Iterates through the BOM to ensure all pricing is current.
* **`RFQValidation.bas`**: The gatekeeper module.
    * `ValidateAllComponentsAndProducts`: Performs a multi-step check (Component counts, Cost > 0, Quantity > 0) before allowing the RFQ to be sent.
* **`PowerAutomateAPI.cls`**:
    * `SendFullyValidatedRFQ`: Gathers Project Data, calculates totals, constructs a JSON payload, and sends it via `MSXML2.XMLHTTP`. Handles JSON escaping manually.

### B. SAP Automation
* **`SAP_CO09_Exporter.bas`**: Connects to SAP GUI to check "Provisional Free Stock" for components. Handles logic differences between HANA and legacy plants.
* **`CreateBOM.bas` / `CreateRouting.bas`**: Scripts that automate transaction codes `CS01` and `CA01` by iterating through Excel ranges and sending `SendKeys`/`findById` commands to SAP.

### C. User Forms
* **`AddProductForm`**: Entry point. Validates unique product numbers and initializes sheet variables.
* **`RoutineForm`**: Complex UI for selecting manufacturing steps. Loads available work centers based on the selected Plant.
* **`ResultsForm`**: Search engine interface. Allows filtering by Plant or TP List using Regex (`FilterComponents.bas`).
* **`frmRoutineVariantEditor`**: Advanced logic for handling Formula vs. Value preservation when cloning routines for variants.

### D. Utilities
* **`Utils.bas`**:
    * `FixDecimalSeparator`: Crucial for international compatibility (handling comma vs dot).
    * `RunProductBasedFormatting`: Applies dynamic conditional formatting (alternating colors per product) to tables.
* **`MassCopy.bas`**: Handles the "Mass Upload" feature, including fuzzy matching and placeholder creation for missing components.

---

## 4. Developer Guidelines (The "Rulebook")

**Important:** Any developer (human or AI) working on this project must adhere to the following rules to maintain stability and performance.

### 4.1. Data Access & Manipulation
1.  **ListObjects Only:** Never reference data by hardcoded ranges (e.g., `Range("A2:Z100")`). Always use `ListObjects("TableName")` and `ListColumns("ColumnName")`.
    * *Bad:* `Cells(i, 5).Value`
    * *Good:* `tbl.ListColumns("Quantity").DataBodyRange.Cells(i).Value`
2.  **Dynamic Column Indices:** Always retrieve column indices dynamically at the start of a sub. Columns in Excel tables may be moved by users; hardcoded indices (e.g., `col = 5`) will break the tool.
3.  **Array Processing:** For loops iterating over >100 rows, read the `DataBodyRange` into a Variant Array, process in memory, and write back. Do not read/write cell-by-cell inside a loop.

### 4.2. Coding Standards
1.  **Option Explicit:** Must be present at the top of every module.
2.  **Error Handling:**
    * Use `On Error GoTo ErrorHandler` for all non-trivial procedures.
    * Use `On Error Resume Next` **only** for specific, contained checks (e.g., checking if a Dictionary key exists or if an SAP object is active) and reset immediately with `On Error GoTo 0`.
3.  **Variable Naming:** Use CamelCase. Prefix variables to indicate type where helpful (e.g., `wsTarget` for Worksheet, `tblBOM` for ListObject).
4.  **JSON Handling:** Use the existing `JsonEscape` function in `PowerAutomateAPI` for any string being sent to the web API.
5. **Use Configuration Class** When using frequently referenced tables such as BOMDefinition or "1. BOM Definition", always import these terms from the config class and declare the variables as option explicit variables and then implement a Initialize function that loads the values via the configuration class.

### 4.3. SAP Scripting Rules
1.  **Session Check:** Always check `If Not IsObject(SAPSession)` before executing script lines.
2.  **Window Handling:** Use `isViewActiveWithinWindow` to handle pop-ups or optional screens gracefully.
3.  **Non-Blocking:** Ensure SAP scripts do not leave the Excel application in a frozen state if SAP crashes.

### 4.4. User Interaction
1.  **Status Bar:** Use `Application.StatusBar` to inform users of progress during long operations (Mass Upload, Database Refresh).
2.  **Decimal Separators:** All user inputs (InputBoxes) expecting numbers must run through `Utils.FixDecimalSeparator` to handle international locale settings.

### 4.5. Safety & Integrity
1.  **Formulas vs. Values:** When copying rows (especially for Variants), check `.HasFormula`. If a cell has a formula, copy the Formula; otherwise, copy the Value.
2.  **Table Integrity:** When adding rows, use `tbl.ListRows.Add`. Do not use `Range.Insert`.

---

## 5. Future Implementation Roadmap

* **Refactor:** Consolidate `CreateBOM` and `CreateRouting` logic to use a shared SAP connection class.
* **Feature:** Implement a "Cost Breakdown" module that visualizes Material vs. Labor costs using the data in `RoutineSummary`.
* **Security:** Move the hardcoded Power Automate URL in `PowerAutomateAPI` to the `Global Variables` sheet.