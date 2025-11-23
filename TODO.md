# Project Improvements & TODOs

## Critical Issues

- [ ] **File Corruption**: The following files are corrupted or binary-encoded and cannot be edited:
  - `Modul1.bas`, `Module2.bas`
  - `Search_copy_components.bas`, `ServoCalculation.bas`, `CopyDataFromBM1Sheet.bas`
  - `RoutineForm.frm` (and other forms)
  - **Action**: Re-export these files using `ExportAllVBACode` from a working Excel instance.
- [x] **Analyzed (Manual Restoration Needed)**: `FilterComponents.bas`, `MassCopy.bas` (Code provided, but file write failed).

## Security Improvements

- [ ] **Hardcoded URLs**: The Power Automate URL is hardcoded in `Project FilesPowerAutomateAPI.cls`.
  - **Fix**: Externalize to "Global Variables" sheet.
- [ ] **SAP Session Security**: `CreateBOM`, `CreateRouting`, `AddComponentAllocation`, and `SAP_CO09_Exporter` all use `GetObject("SAPGUI")` without validating the connection.
  - **Fix**: Add checks to ensure the script attaches to the correct system/client.
- [ ] **Error Handling**: Widespread use of `On Error Resume Next` (e.g., `PurchasingInfo.bas`, `MassCopy.bas`).
  - **Fix**: Replace with structured error handling (`On Error GoTo`) to catch and log errors.
- [ ] **JSON Injection**: `JsonEscape` is manual.
  - **Fix**: Use a dedicated JSON library.

## Performance Optimization

- [ ] **Sheet References**: Avoid repeated `ThisWorkbook.Sheets(...)` calls in loops.
- [ ] **Loop Efficiency**:
  - `ApplyRowFormatting`: Optimize column iteration.
  - `VariantConfig.GenerateWorkCenterSummary`: Optimize nested loops for large datasets.
  - `FilterComponents.bas`: Already uses array-based processing (Good!).
- [ ] **Regex Object**: Instantiate `VBScript.RegExp` once (static/global) instead of per call (relevant for `FilterComponents.bas` and `MassCopy.bas`).

## Code Quality & Maintainability

- [ ] **Refactoring**: `ClearTables.bas` contains multiple subs (`ClearSelectedRoutinesTable`, `ClearSelectedComponentsTable`, etc.) with nearly identical logic.
  - **Improvement**: Create a generic `ClearTable(sheetName, tableName)` sub.
- [ ] **Duplicate Files**: Consolidate `PowerAutomateAPI.bas` and its `.cls` counterpart.
- [ ] **Magic Strings**: Centralize column names and SAP IDs in a `Config` module.
- [ ] **Modularization**: Break down large subs like `SendToPowerAutomate`.

## Functional Enhancements

- [ ] **User Feedback**: Replace modal `MsgBox` with `Application.StatusBar` for non-critical updates.
- [ ] **Logging**: Implement a central logging mechanism.
