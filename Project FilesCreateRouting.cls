Attribute VB_Name = "CreateRouting"
' EDIT THIS LINE
' The program will start from this row.
' The first row should always be a header row (i.e. the value of column A should be 'H' in this row
Const FIRST_ROW = 4


Rem ADDED BY EXCEL *************************************

Dim objExcel
Dim objSheet, intRow, i, real_first_row
Dim SAPSession

Const COL_WITH_ENDMATERIAL_NUMBER = 2
Const COL_WITH_PLANT = 3

Const COL_WITH_USAGE = 4
Const COL_WITH_STATUS = 5
Const COL_WITH_PLANER_GROUP = 6
Const COL_WITH_OP_INDEX = 8
Const COL_WITH_WORKCENTER = 10
Const COL_WITH_CONTROL_KEY = 12
Const COL_WITH_DESCRIPTION = 14
Const COL_WITH_BASE_QTY = 15
Const COL_WITH_SETUP_TIME = 17
Const COL_WITH_MACHINE_TIME = 19
Const COL_WITH_PERSONAL_TIME = 21




Const COL_WITH_ROUTING_LOG = 22
Const ERR_MSG_PN_NOT_FOUND = "n/a"
Const HEADER_SPECIFICIATION = "H"
Const ITEM_SPECIFICATION = "O"





Const MAIN_WINDOW = "GuiMainWindow"
Const POPUP_WINDOW = "GuiModalWindow"
    
Sub MainCreateRouting()

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
        Call DoCreateRouting
    Else
        MsgBox "Incorrect Row, re-try!"
        Set SAPSession = Nothing
        Set SAPConnection = Nothing
        Set SapGuiAuto = Nothing
    End If
End Sub
    
Private Sub DoCreateRouting()

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

    row = FIRST_ROW
    Call OpenTransaction("ca01")
    Do While (objSheet.Cells(row, 1) <> "")
        Call SetHeaderInformationForRouting(row)
        SAPSession.findById("wnd[0]").sendVKey 7
        Call CopyPositionDataFromExcelToRouting(row)
        SAPSession.findById("wnd[0]/tbar[0]/btn[11]").press
        row = row + 1
        Do While (objSheet.Cells(row, 1) = ITEM_SPECIFICATION)
            row = row + 1
        Loop
    Loop
    MsgBox "Complete", vbSystemModal
End Sub




' Returns true if the SAP-MessageBar shows either an error (type "E") or a warning (type "W").
Public Function isErrorOrWarningMsg()
    isErrorOrWarningMsg = (SAPSession.findById("wnd[0]/sbar").messagetype = "E" Or SAPSession.findById("wnd[0]/sbar").messagetype = "W")
End Function

' Returns true if the SAP-MessageBar shows either an error (type "E")
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

Private Sub SetHeaderInformationForRouting(currentRow)
    
    SAPSession.findById("wnd[0]/usr/ctxtRC27M-MATNR").Text = objSheet.Cells(currentRow, COL_WITH_ENDMATERIAL_NUMBER).Text
    SAPSession.findById("wnd[0]/usr/ctxtRC27M-WERKS").Text = objSheet.Cells(currentRow, COL_WITH_PLANT).Text
    SAPSession.findById("wnd[0]").sendVKey 0
    If isErrorMsg() Then
        Call SetErrorMsg(currentRow, COL_WITH_ROUTING_LOG, SAPSession.findById("wnd[0]/sbar").Text)
        Call SetErrorMsg(currentRow, COL_WITH_ROUTING_LOG, "Aborted")
        Exit Sub
    End If
    SAPSession.findById("wnd[0]").sendVKey 0
    
    
    ' check if there is a separate popup window
    If SAPSession.ActiveWindow.Type = POPUP_WINDOW Then
        SAPSession.findById("wnd[1]").sendVKey 0
    End If
    ' Check if there are multiple alternatives
    If isViewActiveWithinWindow("wnd[0]/usr/tblSAPLCSDITCALT") Then
        Call SetErrorMsg(currentRow, COL_WITH_ROUTING_LOG, "Aborted! There are multiple Routing alternatives for this article!")
        Exit Sub
    End If
    ' Set header data
    SAPSession.findById("wnd[0]/usr/subGENERALVW:SAPLCPDA:1211/ctxtPLKOD-VERWE").Text = objSheet.Cells(currentRow, COL_WITH_USAGE).Text
    SAPSession.findById("wnd[0]/usr/subGENERALVW:SAPLCPDA:1211/ctxtPLKOD-STATU").Text = objSheet.Cells(currentRow, COL_WITH_STATUS).Text
    SAPSession.findById("wnd[0]/usr/subGENERALVW:SAPLCPDA:1211/ctxtPLKOD-VAGRP").Text = objSheet.Cells(currentRow, COL_WITH_PLANER_GROUP).Text
End Sub

Private Sub CopyPositionDataFromExcelToRouting(currentRow)
    tmpRow = currentRow
    
    If objSheet.Cells(tmpRow, 1).Text = HEADER_SPECIFICIATION Then
        tmpRow = tmpRow + 1
    End If
    absoluteRow = 0
    Do While (objSheet.Cells(tmpRow, 1).Text = ITEM_SPECIFICATION)
        SAPSession.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VORNR[0," & absoluteRow & "]").Text = objSheet.Cells(tmpRow, COL_WITH_OP_INDEX).Text
        SAPSession.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-ARBPL[2," & absoluteRow & "]").Text = objSheet.Cells(tmpRow, COL_WITH_WORKCENTER).Text
        SAPSession.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-STEUS[4," & absoluteRow & "]").Text = objSheet.Cells(tmpRow, COL_WITH_CONTROL_KEY).Text
        SAPSession.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6," & absoluteRow & "]").Text = objSheet.Cells(tmpRow, COL_WITH_DESCRIPTION).Text
        SAPSession.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[16," & absoluteRow & "]").Text = objSheet.Cells(tmpRow, COL_WITH_SETUP_TIME).Text
        SAPSession.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[19," & absoluteRow & "]").Text = objSheet.Cells(tmpRow, COL_WITH_MACHINE_TIME).Text
        SAPSession.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW03[22," & absoluteRow & "]").Text = objSheet.Cells(tmpRow, COL_WITH_PERSONAL_TIME).Text
                
        tmpRow = tmpRow + 1
        absoluteRow = absoluteRow + 1
        ' scroll window to current position after every 20 components
        If CInt(SAPSession.findById("wnd[0]/usr/txtRC27X-ENTRIES").Text) Mod 20 = 0 Then
            SAPSession.findById("wnd[0]").sendVKey 0
            SAPSession.findById("wnd[0]/usr/txtRC27X-ENTRY_ACT").Text = SAPSession.findById("wnd[0]/usr/txtRC27X-ENTRIES").Text
            SAPSession.findById("wnd[0]").sendVKey 0
            absoluteRow = 1
        End If
    Loop
    
End Sub



