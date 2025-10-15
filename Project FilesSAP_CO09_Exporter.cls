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
    ws.Cells(1, 1).value = CDbl(SAPSession.findById("wnd[0]/usr/tbl/SAPAPO/SAPLATP4CTR_400/txt/SAPAPO/ATPDE-CATPQTY[6,0]").Text)


End Sub
'------------------------------------------------------------------------------
' Subroutine: RunCO09ForAllBOMRows
'
' Description:
'   Automates the process of running SAP CO09 transaction for each row in the
'   "BOMDefinition" table on the "1. BOM Definition" worksheet. For each material
'   and plant combination, retrieves the provisional free stock from SAP and
'   writes it back to the corresponding row in the table.
'
'   - Connects to SAP GUI Scripting Engine.
'   - Iterates through all BOM rows, skipping those with missing material or plant.
'   - Handles SAP GUI errors and missing columns gracefully.
'   - Updates the "Provisional Free Stock" column with the retrieved value.
'
' Preconditions:
'   - SAP GUI must be running and scripting enabled.
'   - "BOMDefinition" table must exist with columns: "Material", "Plant", "Provisional Free Stock".
'
' Postconditions:
'   - Each BOM row will have its provisional free stock updated from SAP CO09.
'   - Displays a message box upon completion or error.
'
' Usage:
'   Call RunCO09ForAllBOMRows from VBA or assign to a button as needed.
'------------------------------------------------------------------------------
Sub RunCO09ForAllBOMRows()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim matnr As String, werks As String, berid As String
    Dim freeStock As Variant
    Dim colMaterial As Long, colPlant As Long, colFreeStock As Long
    Dim SapGuiAuto As Object
    Dim application As Object
    Dim i As Long

    ' Set worksheet and table
    Set ws = ThisWorkbook.Sheets("1. BOM Definition")
    Set tbl = ws.ListObjects("BOMDefinition")
    
    ' Get column indices
    On Error Resume Next
    colMaterial = tbl.ListColumns("Material").Index
    colPlant = tbl.ListColumns("Plant").Index
    colFreeStock = tbl.ListColumns("Provisonal Free Stock").Index
    On Error GoTo 0
    If colMaterial = 0 Or colPlant = 0 Or colFreeStock = 0 Then
        MsgBox "One or more required columns ('Material', 'Plant', 'Provisional Free Stock') not found in BOMDefinition table.", vbCritical
        Exit Sub
    End If

    ' Connect to SAP
    On Error Resume Next
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

    ' Loop through each row in BOMDefinition
    For Each row In tbl.ListRows
        matnr = row.Range(1, colMaterial).value
        werks = row.Range(1, colPlant).value
        If werks = "TP List" Then
            werks = "5100"
        End If
        berid = werks                            ' If BERID is always the same as plant
        
        ' Skip if any required value is missing
        If matnr = "" Or werks = "" Then
            row.Range(1, colFreeStock).value = "[Missing Material/Plant]"
            GoTo NextRow
        End If

        On Error Resume Next
        SAPSession.findById("wnd[0]").maximize
        SAPSession.findById("wnd[0]/tbar[0]/okcd").Text = "/nco09"
        SAPSession.findById("wnd[0]").sendVKey 0
        SAPSession.findById("wnd[1]/tbar[0]/btn[0]").press
        SAPSession.findById("wnd[0]/tbar[0]/okcd").Text = "/nco09"
        SAPSession.findById("wnd[0]").sendVKey 0
        
        
        SAPSession.findById("wnd[0]/usr/ctxtCAUFVD-MATNR").Text = matnr
        SAPSession.findById("wnd[0]/usr/ctxtCAUFVD-WERKS").Text = werks
        SAPSession.findById("wnd[0]/usr/ctxtAFPOD-BERID").Text = berid
        SAPSession.findById("wnd[0]/usr/ctxtCAUFVD-PRREG").Text = "A"
        SAPSession.findById("wnd[0]/usr/chkCAUFVD-PRMBD").Selected = True
        SAPSession.findById("wnd[0]").sendVKey 0
        
        ' Try to get the free stock value
        On Error Resume Next
        freeStock = CDbl(SAPSession.findById("wnd[0]/usr/tbl/SAPAPO/SAPLATP4CTR_400/txt/SAPAPO/ATPDE-CATPQTY[6,0]").Text)
        If Err.Number <> 0 Then
            freeStock = 0
        End If
        On Error GoTo 0
        
        row.Range(1, colFreeStock).value = freeStock
NextRow:
    Next row
    MsgBox "CO09 automation completed for all BOM rows.", vbInformation
End Sub

