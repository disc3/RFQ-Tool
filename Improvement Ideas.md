# **RFQ Tool: Technical Audit and Modernization Strategy**

## **Executive Summary**

Application Overview  
The RFQ Tool is a sophisticated Excel-based application designed to bridge the gap between calculation logic (Excel) and ERP execution (SAP). It features advanced functionalities including SAP GUI scripting, API integration (Power Automate), and complex matrix calculations (ServoCalculation).  
Current State Assessment  
The tool currently functions as a "Shadow IT" solution. While powerful, it exhibits distinct signs of organic growth—inconsistent coding standards, duplicate logic, and potential security risks—which threaten its long-term maintainability and stability.  
Core Recommendation  
Immediate focus must be placed on modularization, security hardening (specifically API key management), and transitioning data processing from row-by-row iterations to array-based memory processing to drastically improve performance.

## **1\. Code Structure & Maintainability Analysis**

### **Current Status**

* **Strengths:** The utilization of ListObjects (Excel Tables) is excellent, providing robustness against row/column insertions compared to standard Range references.  
* **Weaknesses:**  
  * **Hardcoding:** Files such as PowerAutomateAPI.bas and DynamicRoutineForEachPlant.bas contain hardcoded sheet names, URLs, and specific logic.  
  * **Duplicate Logic:** Significant logic overlap exists between Search\_copy\_components.bas and MassCopy.bas.  
  * **God Modules:** Modules like Search\_copy\_components.bas violate separation of concerns by handling UI, data processing, and formatting simultaneously.

### **Recommendations**

| Feature | Current Implementation | Proposed Improvement |
| :---- | :---- | :---- |
| **Configuration** | Hardcoded in VBA (e.g., API URLs, Plant IDs). | Move static data to a hidden "Global Variables" or "Config" sheet. Create a Config class to read these values. |
| **SAP Logic** | Scattered across CreateBOM, CreateRouting, etc. | Create a SAPWrapper.cls class. Encapsulate connection logic, error handling, and transaction codes. |
| **Error Handling** | Inconsistent On Error Resume Next (e.g., PurchasingInfo.bas). | Implement a global error handler that logs errors to a text file or log sheet instead of suppressing them. |

## **2\. Performance Optimization**

### **Analysis**

The tool suffers from performance bottlenecks in modules that iterate through table rows individually. VBA interactions with the Excel grid are computationally expensive.

* **Problem Area:** UpdateRoutinesBySelectedPlant.bas and parts of RFQValidation.bas read/write cells individually.  
* **Gold Standard:** ServoCalculation.bas correctly loads the entire table into an array, processes in memory, and writes back in one operation.

### **Concrete Action Steps**

#### **1\. Refactor Loop Logic**

Rewrite UpdateSelectedRoutines to read the source table into a Variant Array. This can reduce execution time from minutes to seconds for large datasets.

**Concept Code:**

Dim dataIn As Variant, dataOut As Variant  
' Read entire data body to memory  
dataIn \= tblSource.DataBodyRange.Value2

' ... process logic in memory arrays ...

' Write back in one operation  
tblDestination.DataBodyRange.Value2 \= dataOut

#### **2\. Optimize Filtering**

FilterComponents.bas currently instantiates VBScript.RegExp inside a loop.

* **Fix:** Move CreateObject("VBScript.RegExp") outside the loop or utilize a Static variable to ensure initialization occurs only once per operation.

#### **3\. Global Screen Update Control**

Implement a SpeedBoost class or module to manage Application.ScreenUpdating, EnableEvents, and calculation modes. Ensure a failsafe (Class Terminate event) resets these settings if the code crashes.

## **3\. Stability & Error Handling**

### **Critical Findings**

* **Silent Failures:** MassCopy.bas utilizes On Error Resume Next broadly. If a column is renamed or a query fails, the code continues silently, risking data corruption.  
* **SAP Connection:** Modules like SAP\_CO09\_Exporter assume GetObject("SAPGUI") will always succeed. This causes crashes if SAP is closed or a modal dialog is active.

### **Recommendations**

1. **Robust SAP Hook:** Implement a connection check loop. If GetObject fails, prompt the user. Verify SAPSession.Info.Transaction ensures the user is on the home screen before script execution.  
2. **Pre-flight Validation:** Expand RFQValidation.cls to perform checks before heavy operations (e.g., ensuring target tables exist and headers match).

## **4\. Security (High Priority)**

### **Analysis**

**Exposed Credentials:** PowerAutomateAPI.bas contains hardcoded URLs with Signature (SAS Tokens) directly in the source code (e.g., ...\&sig=0\_CKd-MQZJrv5YaVY0...). Any user with access to the Excel file can trigger the Power Automate flows.

### **Recommendations**

1. **Externalize Secrets:** Move API URLs/Keys to the "Global Variables" sheet within a named range (e.g., API\_Config).  
2. **Access Control:** If possible, restrict the Power Automate flow to accept requests only from specific IPs or require a secondary header token that is not hardcoded.

## **5\. Usability & User Experience (UX)**

### **Analysis**

* **Disruptive Flow:** Forms like ResultsForm and AddComponentForm rely on MsgBox for validation errors, interrupting the user workflow.  
* **Lack of Feedback:** Long-running tasks (Mass Upload) lack progress indicators.

### **Recommendations**

1. **Visual Validation:** Replace error message boxes with visual cues (e.g., turning a textbox background red) and "Required" labels.  
2. **Progress Indicators:** Implement Application.StatusBar updates for loops in MassCopy and ServoCalculation.  
3. **Modeless Forms:** Set ShowModal \= False for ResultsForm to allow users to scroll and verify Excel data while the form remains open.

## **6\. Roadmap for Refactoring**

To implement changes without disrupting current operations, follow this phased approach:

### **Phase 1: Security & Configuration (Immediate)**

* \[ \] Move Power Automate URLs and Secrets to a hidden worksheet.  
* \[ \] Define all Column Names as Const strings in a shared module to prevent typo-induced bugs.

### **Phase 2: Performance (High Impact)**

* \[ \] Refactor UpdateSelectedRoutines and RefreshBOMData to use **Memory Arrays** instead of Range loops.  
* \[ \] Optimize Regex usage in filtering modules.

### **Phase 3: Architecture (Long Term)**

* \[ \] Consolidate duplicate logic from Search\_copy\_components and MassCopy into a single ComponentController module.  
* \[ \] Create a dedicated SAP\_Handler module to manage all SAP interactions safely.

**Conclusion:** Addressing **Array Processing (Performance)** and **Hardcoded Secrets (Security)** will provide the highest immediate ROI for this codebase.