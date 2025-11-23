Attribute VB_Name = "AddComponentAllocation"
' EDIT THIS LINE
' The program will start from this row.
' The first row should always be a header row (i.e. the value of column A should be 'H' in this row)
Const FIRST_ROW = 3

Dim objExcel
Dim objSheet, i, real_first_row
Dim SAPSession

Const COL_WITH_ENDMATERIAL_NUMBER = 2
Const COL_WITH_PLANT = 3
Const COL_WITH_BASE_QTY = 5
Const COL_WITH_DIVISION = 14
Const COL_WITH_COMPONENT_ALLOC = 16
Const COL_WITH_POS_NO = 8
Const COL_WITH_DIVISION_ERROR = 18
Const COL_WITH_DIVISION_ERROR_COUNTER = 19
Const COL_WITH_BOM_LOG = 20
Const ERR_MSG_PN_NOT_FOUND = "n/a"
Const ERR_MSG_NO_COLUMN_SPECIFICATION = "Please enter either 'I' or 'H' in the first column of every end product / component"
Const DIVISION_CABLE = "YES"




Const MAIN_WINDOW = "GuiMainWindow"
Const POPUP_WINDOW = "GuiModalWindow"

Sub MainComponentAllocation()

    If Not IsObject(SAPapplication) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set SAPapplication = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(SAPConnection) Then
       Set SAPConnection = SAPapplication.Children(0)
    End If
    If Not IsObject(SAPSession) Then
       Set SAPSession = SAPConnection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject SAPSession, "on"
       WScript.ConnectObject application, "on"
    End If
    
    Set objExcel = GetObject(, "Excel.Application")
    Set objSheet = objExcel.ActiveWorkbook.ActiveSheet
    
    On Error Resume Next

    real_first_row = application.InputBox(Prompt:="Please enter the first row for which you want to upload the data (has to have an 'H' in column A)", Title:="Pick the first row", Default:=FIRST_ROW)
    On Error GoTo 0
    
    If objSheet.Cells(real_first_row, 1) = "H" Then
        Call DoComponentAllocation
    Else
        MsgBox "Incorrect Row, re-try!"
        Set SAPSession = Nothing
        Set SAPConnection = Nothing
        Set SapGuiAuto = Nothing
    End If
End Sub
    
Private Sub DoComponentAllocation()

    row = real_first_row
    Call OpenTransaction("ca02")
    Do While (objSheet.Cells(row, 1) <> "")
        Call SetHeaderInformationForBom(row)
        Call AddComponentAllocations(row)
        ' save
        SAPSession.findById("wnd[0]/tbar[0]/btn[11]").press
        row = row + 1
        Do While (objSheet.Cells(row, 1) = "I")
            row = row + 1
        Loop
    Loop
    
    MsgBox ("Complete")
    Set SAPSession = Nothing
    Set SAPConnection = Nothing
    Set SapGuiAuto = Nothing
End Sub

' Returns true if the SAP-MessageBar shows either an error (type "E") or a warning (type "W").
Public Function isErrorOrWarningMsg()
    isErrorOrWarningMsg = (SAPSession.findById("wnd[0]/sbar").messagetype = "E" Or SAPSession.findById("wnd[0]/sbar").messagetype = "W")
End Function

Public Function isErrorMsg()
    isErrorMsg = (SAPSession.findById("wnd[0]/sbar").messagetype = "E")
End Function

' Checks if a given view (which technical name must be the input parameter) is active.
' If the view is not active, the SAPSession.findById expression will lead to an error -> fixed with its surrounding error handler.
Public Function isViewActiveWithinWindow(id)
    isViewActiveWithinWindow = False
    On Error Resume Next
    isViewActiveWithinWindow = Not IsObject(SAPSession.findById(id))
    On Error GoTo 0
End Function

Private Sub OpenTransaction(transactionCode)
    SAPSession.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & transactionCode
    SAPSession.findById("wnd[0]").sendVKey 0
End Sub

Private Sub SetErrorMsg(row, column, errMsg)
    objSheet.Cells(row, column) = errMsg
End Sub

Private Sub SetHeaderInformationForBom(currentRow)
    
    SAPSession.findById("wnd[0]/usr/ctxtRC27M-MATNR").Text = objSheet.Cells(currentRow, COL_WITH_ENDMATERIAL_NUMBER).Text
    SAPSession.findById("wnd[0]/usr/ctxtRC27M-WERKS").Text = objSheet.Cells(currentRow, COL_WITH_PLANT).Text
    
    SAPSession.findById("wnd[0]").sendVKey 0
    If isErrorMsg() Then
        Call SetErrorMsg(currentRow, 20, SAPSession.findById("wnd[0]/sbar").Text)
        Call SetErrorMsg(currentRow, 20, "Aborted")
        Exit Sub
    End If
    SAPSession.findById("wnd[0]").sendVKey 0
    
    
    ' Go-To Component allocation screen
    SAPSession.findById("wnd[0]/tbar[1]/btn[7]").press
    
End Sub

Private Sub AddComponentAllocations(currentRow)
    tmpRow = currentRow
    
    If objSheet.Cells(tmpRow, 1).Text = "H" Then
        tmpRow = tmpRow + 1
    End If
    
    Do While (objSheet.Cells(tmpRow, 1).Text = "I")
        If objSheet.Cells(tmpRow, COL_WITH_COMPONENT_ALLOC) <> "" Then
            ' Choose position number of item
            SAPSession.findById("wnd[0]/tbar[1]/btn[42]").press
            SAPSession.findById("wnd[1]/usr/txtSEARCH_BY-POSNR").Text = objSheet.Cells(tmpRow, COL_WITH_POS_NO).Text
            SAPSession.findById("wnd[1]").sendVKey 0
            ' Do component allocation
            SAPSession.findById("wnd[0]/tbar[1]/btn[5]").press
            SAPSession.findById("wnd[1]/usr/txtRCM01-VORNR").Text = objSheet.Cells(tmpRow, COL_WITH_COMPONENT_ALLOC).Text
            SAPSession.findById("wnd[1]").sendVKey 0
        End If
        tmpRow = tmpRow + 1
    Loop
    'msgbox("old value: " & currentRow & "; new value: " & tmpRow)
End Sub

