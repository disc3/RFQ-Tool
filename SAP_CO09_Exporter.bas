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
'   Automates SAP CO09 transaction for "BOMDefinition".
'
' Logic Updates:
'   - isHanaCompany: Defined as TRUE unless plant starts with "F" or "P".
'   - If isHanaCompany is TRUE:
'       - Fills BERID, Selects PRMBD.
'       - Reads stock from SAPAPO table index [6,0].
'   - If isHanaCompany is FALSE (Plants F... or P...):
'       - Skips BERID and PRMBD.
'       - Reads stock from MDEZ table index [5,0].
'------------------------------------------------------------------------------
Sub RunCO09ForAllBOMRows()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim matnr As String, werks As String, berid As String
    Dim freeStock As Variant
    Dim colMaterial As Long, colPlant As Long, colFreeStock As Long
    
    ' SAP Objects
    Dim SapGuiAuto As Object
    Dim application As Object
    Dim Connection As Object
    Dim SAPSession As Object
    
    ' Logic Variables
    Dim isHanaCompany As Boolean
    Dim sapReadID As String

    ' Set worksheet and table
    Set ws = ThisWorkbook.Sheets("1. BOM Definition")
    On Error Resume Next
    Set tbl = ws.ListObjects("BOMDefinition")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "Table 'BOMDefinition' not found on sheet '1. BOM Definition'.", vbCritical
        Exit Sub
    End If
    
    ' Get column indices
    On Error Resume Next
    colMaterial = tbl.ListColumns("Material").Index
    colPlant = tbl.ListColumns("Plant").Index
    colFreeStock = tbl.ListColumns("Provisonal Free Stock").Index
    On Error GoTo 0

    If colMaterial = 0 Or colPlant = 0 Or colFreeStock = 0 Then
        MsgBox "Required columns ('Material', 'Plant', 'Provisonal Free Stock') not found.", vbCritical
        Exit Sub
    End If

    ' ---------------------------------------------------------
    ' SAP Connection Setup
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
    ' Row Iteration
    ' ---------------------------------------------------------
    For Each row In tbl.ListRows
        matnr = row.Range(1, colMaterial).Value
        werks = row.Range(1, colPlant).Value
        
        ' Handle Logic: TP List
        If werks = "TP List" Then werks = "5100"
        berid = werks
        
        ' Validation
        If matnr = "" Or werks = "" Then
            row.Range(1, colFreeStock).Value = "[Missing Data]"
            GoTo NextRow
        End If

        ' -----------------------------------------------------
        ' Determine Logic Type (HANA vs Legacy/MDEZ)
        ' -----------------------------------------------------
        ' If Plant starts with F or P, it is NOT a Hana Company
        If UCase(Left(werks, 1)) = "F" Or UCase(Left(werks, 1)) = "P" Then
            isHanaCompany = False
        Else
            isHanaCompany = True
        End If

        ' -----------------------------------------------------
        ' SAP GUI Interaction
        ' -----------------------------------------------------
        On Error Resume Next
        
        ' Reset Transaction
        SAPSession.findById("wnd[0]").maximize
        SAPSession.findById("wnd[0]/tbar[0]/okcd").Text = "/nco09"
        SAPSession.findById("wnd[0]").sendVKey 0
        
        ' Handle potential popup (wnd[1])
        If Not SAPSession.findById("wnd[1]/tbar[0]/btn[0]") Is Nothing Then
            SAPSession.findById("wnd[1]/tbar[0]/btn[0]").press
            SAPSession.findById("wnd[0]/tbar[0]/okcd").Text = "/nco09"
            SAPSession.findById("wnd[0]").sendVKey 0
        End If
        
        ' --- Fill Fields ---
        SAPSession.findById("wnd[0]/usr/ctxtCAUFVD-MATNR").Text = matnr
        SAPSession.findById("wnd[0]/usr/ctxtCAUFVD-WERKS").Text = werks
        SAPSession.findById("wnd[0]/usr/ctxtCAUFVD-PRREG").Text = "A"
        
        ' Only fill BERID and check PRMBD if it IS a Hana Company
        If isHanaCompany Then
            SAPSession.findById("wnd[0]/usr/ctxtAFPOD-BERID").Text = berid
            SAPSession.findById("wnd[0]/usr/chkCAUFVD-PRMBD").Selected = True
        End If
        
        ' Execute Search
        SAPSession.findById("wnd[0]").sendVKey 0
        
        ' --- Read Data ---
        ' Determine the ID path based on company type
        If isHanaCompany Then
            ' Standard HANA ID
            sapReadID = "wnd[0]/usr/tbl/SAPAPO/SAPLATP4CTR_400/txt/SAPAPO/ATPDE-CATPQTY[6,0]"
        Else
            ' Non-HANA / Legacy (F or P Plants) ID
            sapReadID = "wnd[0]/usr/tbl/MDEZ/SAPLATP4CTR_400/txt/MDEZ-MNG04[5,0]"
        End If
        
        freeStock = CDbl(SAPSession.findById(sapReadID).Text)
        
        If Err.Number <> 0 Then
            freeStock = 0
            Err.Clear
        End If
        On Error GoTo 0
        
        ' Write result
        row.Range(1, colFreeStock).Value = freeStock

NextRow:
    Next row
    
    MsgBox "CO09 automation completed.", vbInformation
End Sub

