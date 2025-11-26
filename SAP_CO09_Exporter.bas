Attribute VB_Name = "SAP_CO09_Exporter"
Sub TestAutomation()
    Dim SapGuiAuto As Object
    Dim application As Object
    Dim ws As Object
    
    Set ws = ThisWorkbook.ActiveSheet
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number <> 0 Then
        MsgBox "Could not get SAPGUI object. Is SAP GUI running?", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    Set application = SapGuiAuto.GetScriptingEngine
    If application Is Nothing Then
        MsgBox "Could not get SAP Scripting Engine. Is scripting enabled?", vbCritical
        Exit Sub
    End If
    If Not IsObject(Connection) Then
        Set Connection = application.Children(0)
    End If
    If Not IsObject(SAPSession) Then
        Set SAPSession = Connection.Children(0)
    End If
    If IsObject(WScript) Then
        WScript.ConnectObject SAPSession, "on"
        WScript.ConnectObject application, "on"
    End If
    SAPSession.findById("wnd[0]").maximize
    SAPSession.findById("wnd[0]/tbar[0]/okcd").Text = "/nco09"
    Err.Clear
    SAPSession.findById("wnd[0]").sendVKey 0
    On Error Resume Next
    
    SAPSession.findById("wnd[1]/tbar[0]/btn[0]").press
    If Err.Number = 0 Then
        SAPSession.findById("wnd[0]/tbar[0]/okcd").Text = "/nco09"
        SAPSession.findById("wnd[0]").sendVKey 0
    End If
    On Error GoTo 0
    SAPSession.findById("wnd[0]/usr/ctxtCAUFVD-MATNR").Text = "1119303"
    SAPSession.findById("wnd[0]/usr/ctxtCAUFVD-WERKS").Text = "5100"
    SAPSession.findById("wnd[0]/usr/ctxtAFPOD-BERID").Text = "5100"
    SAPSession.findById("wnd[0]/usr/ctxtCAUFVD-PRREG").Text = "A"
    SAPSession.findById("wnd[0]/usr/chkCAUFVD-PRMBD").Selected = True
    SAPSession.findById("wnd[0]").sendVKey 0
    ws.Cells(1, 1).Value = CDbl(SAPSession.findById("wnd[0]/usr/tbl/SAPAPO/SAPLATP4CTR_400/txt/SAPAPO/ATPDE-CATPQTY[6,0]").Text)


End Sub
'------------------------------------------------------------------------------
' Subroutine: RunCO09ForAllBOMRows
'
' Description:
'   Automates SAP CO09 transaction for "BOMDefinition" table.
'
' Logic:
'   - Checks if Plant (Werks) starts with "F" or "P".
'   - IF "F" or "P" (Non-HANA):
'       - Logic: Skips BERID/PRMBD fields.
'       - Read Path: MDEZ table (Row 5).
'   - ELSE (HANA):
'       - Logic: Fills BERID, Checks PRMBD.
'       - Read Path: SAPAPO table (Row 6).
'
' Safety:
'   - Only updates the Excel cell if the SAP read command is successful.
'------------------------------------------------------------------------------
Sub RunCO09ForAllBOMRows()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim matnr As String, werks As String, berid As String
    Dim freeStock As Double
    Dim colMaterial As Long, colPlant As Long, colFreeStock As Long
    
    ' SAP Objects
    Dim SapGuiAuto As Object
    Dim application As Object
    Dim Connection As Object
    Dim SAPSession As Object
    
    ' Logic Variables
    Dim isHanaCompany As Boolean
    Dim sapReadID As String
    Dim rawSAPValue As String
    Dim readSuccess As Boolean

    ' ---------------------------------------------------------
    ' 1. Excel Setup
    ' ---------------------------------------------------------
    Set ws = ThisWorkbook.Sheets("1. BOM Definition")
    On Error Resume Next
    Set tbl = ws.ListObjects("BOMDefinition")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "Table 'BOMDefinition' not found on sheet '1. BOM Definition'.", vbCritical
        Exit Sub
    End If
    
    ' Get column indices (Adjust header names if necessary)
    On Error Resume Next
    colMaterial = tbl.ListColumns("Material").Index
    colPlant = tbl.ListColumns("Plant").Index
    colFreeStock = tbl.ListColumns("Provisonal Free Stock").Index ' Kept spelling as per your snippet
    On Error GoTo 0

    If colMaterial = 0 Or colPlant = 0 Or colFreeStock = 0 Then
        MsgBox "Required columns ('Material', 'Plant', 'Provisonal Free Stock') not found.", vbCritical
        Exit Sub
    End If

    ' ---------------------------------------------------------
    ' 2. SAP Connection Setup
    ' ---------------------------------------------------------
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number <> 0 Then
        MsgBox "Could not get SAPGUI object. Is SAP GUI running?", vbCritical
        Exit Sub
    End If
    
    Set application = SapGuiAuto.GetScriptingEngine
    If application Is Nothing Then
        MsgBox "Could not get SAP Scripting Engine.", vbCritical
        Exit Sub
    End If
    
    Set Connection = application.Children(0)
    Set SAPSession = Connection.Children(0)
    On Error GoTo 0

    ' ---------------------------------------------------------
    ' 3. Row Iteration
    ' ---------------------------------------------------------
    For Each row In tbl.ListRows
        matnr = row.Range(1, colMaterial).Value
        werks = row.Range(1, colPlant).Value
        
        ' Logic: TP List Handling
        If werks = "TP List" Then werks = "5100"
        berid = werks
        
        ' Validation: Skip if data missing
        If matnr = "" Or werks = "" Then
            row.Range(1, colFreeStock).Value = "[Missing Data]"
            GoTo NextRow
        End If

        ' Logic: Determine Company Type (HANA vs Legacy)
        ' If Plant starts with F or P -> Non-HANA
        If UCase(Left(werks, 1)) = "F" Or UCase(Left(werks, 1)) = "P" Then
            isHanaCompany = False
        Else
            isHanaCompany = True
        End If

        ' -----------------------------------------------------
        ' 4. SAP GUI Interaction
        ' -----------------------------------------------------
        On Error Resume Next
        
        ' Reset Transaction (/nco09)
        SAPSession.findById("wnd[0]").maximize
        SAPSession.findById("wnd[0]/tbar[0]/okcd").Text = "/nco09"
        SAPSession.findById("wnd[0]").sendVKey 0
        
        ' Handle potential "Exit" popup from previous runs
        If Not SAPSession.findById("wnd[1]/tbar[0]/btn[0]") Is Nothing Then
            SAPSession.findById("wnd[1]/tbar[0]/btn[0]").press
            SAPSession.findById("wnd[0]/tbar[0]/okcd").Text = "/nco09"
            SAPSession.findById("wnd[0]").sendVKey 0
        End If
        
        ' Fill Common Fields
        SAPSession.findById("wnd[0]/usr/ctxtCAUFVD-MATNR").Text = matnr
        SAPSession.findById("wnd[0]/usr/ctxtCAUFVD-WERKS").Text = werks
        SAPSession.findById("wnd[0]/usr/ctxtCAUFVD-PRREG").Text = "A"
        
        ' Conditional Input: Only fill BERID/PRMBD if HANA
        If isHanaCompany Then
            SAPSession.findById("wnd[0]/usr/ctxtAFPOD-BERID").Text = berid
            SAPSession.findById("wnd[0]/usr/chkCAUFVD-PRMBD").Selected = True
            SAPSession.findById("wnd[0]/usr/ctxtCAUFVD-PRREG").Text = "ZA"
        Else
            SAPSession.findById("wnd[0]/usr/chkCAUFVD-PRMBD").Selected = False
        End If
        
        ' Execute
        SAPSession.findById("wnd[0]").sendVKey 0
        
        ' -----------------------------------------------------
        ' 5. Read Data (Error Safe)
        ' -----------------------------------------------------
        ' Determine the correct table ID based on company type
        If isHanaCompany Then
            ' HANA Table (SAPAPO), Row 6
            sapReadID = "wnd[0]/usr/tbl/SAPAPO/SAPLATP4CTR_400/txt/SAPAPO/ATPDE-CATPQTY[6,0]"
        Else
            ' Non-HANA Table (MDEZ), Row 5 - Fixed missing slashes here
            sapReadID = "wnd[0]/usr/tblSAPLATP4CTR_400/txtMDEZ-MNG04[5,0]"
        End If
        
        ' Attempt to read the value
        Err.Clear
        rawSAPValue = SAPSession.findById(sapReadID).Text ' Using .Text is safer than .Value for table cells
        
        ' Check if the command caused an error
        If Err.Number = 0 Then
            ' If no error, parse and write to Excel
            If IsNumeric(rawSAPValue) Then
                row.Range(1, colFreeStock).Value = CDbl(rawSAPValue)
            Else
                ' Handle cases where text is returned but not a number (optional)
                row.Range(1, colFreeStock).Value = 0
            End If
        Else
            ' If error occurred (e.g., ID not found/Table empty), do not write anything
            ' Optional: row.Range(1, colFreeStock).Value = "Error"
            Err.Clear
        End If
        
        On Error GoTo 0

NextRow:
    Next row
    
    MsgBox "CO09 automation completed.", vbInformation
End Sub
