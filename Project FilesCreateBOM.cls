Attribute VB_Name = "CreateBOM"
' EDIT THIS LINE
' The program will start from this row.
' The first row should always be a header row (i.e. the value of column A should be 'H' in this row)
Const FIRST_ROW = 3

Dim objExcel
Dim objSheet, i, real_first_row
Dim SAPSession

Const COL_ENDMATERIAL_NUMBER = 2
Const COL_PLANT = 3
Const COL_BASE_QTY = 5
Const COL_POSITION_NO = 8
Const COL_POSITION_MATERIAL_NO = 9
Const COL_POSITION_MATERIAL_QTY = 11
Const COL_COMPONENT_SCRAP = 13
Const COL_DIVISION = 14
Const COL_POSITION_INDIVIDUAL_QTY = 15
Const COL_FIXED_QTY = 17
Const COL_COSTING_RELEVANCY = 18
Const COL_TEXT_ONE = 19
Const COL_TEXT_TWO = 20

Const ERR_MSG_PN_NOT_FOUND = "n/a"
Const ERR_MSG_NO_COLUMN_SPECIFICATION = "Please enter either 'I' or 'H' in the first column of every end product / component"
Const DIVISION_CABLE = "YES"




Const MAIN_WINDOW = "GuiMainWindow"
Const POPUP_WINDOW = "GuiModalWindow"
  
Sub MainCreateBOM()

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
        Call DoBomCreate
    Else
        MsgBox "Incorrect Row, re-try!"
        Set SAPSession = Nothing
        Set SAPConnection = Nothing
        Set SapGuiAuto = Nothing
    End If
End Sub


Private Sub DoBomCreate()

    row = real_first_row
    Call OpenTransaction("cs01")
    Do While (objSheet.Cells(row, 1) <> "")
        ' Actions for header row ("H" in column 1)
        Call SetHeaderInformationForBom(row)
        rowInWindow = 0
        row = row + 1

        ' Actions for each item (for all components)
        Do While (objSheet.Cells(row, 1) = "I")
            

            Call CopyPositionDataFromExcelToBom(row)
            
            If objSheet.Cells(row, COL_DIVISION) = DIVISION_CABLE Then
                Call AddClassificationForPosition(row, rowInWindow)
            End If
            
            'TODO
            If objSheet.Cells(row, COL_COMPONENT_SCRAP) <> "" Or objSheet.Cells(row, COL_FIXED_QTY) <> "" Then
                Call AddDetailScreenBasicData(row, rowInWindow)
            End If
            
            If objSheet.Cells(2, COL_COSTING_RELEVANCY) <> "" Then
                Call AddDetailScreenStatusLongtextData(row, rowInWindow)
            End If
            row = row + 1
            rowInWindow = rowInWindow + 1
            SAPSession.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").verticalScrollbar.Position = rowInWindow
        Loop
        SAPSession.findById("wnd[0]/tbar[0]/btn[11]").press
    Loop
    
    Set SAPSession = Nothing
    Set SAPConnection = Nothing
    Set SapGuiAuto = Nothing

    MsgBox ("Complete")
End Sub


' Returns true if the SAP-MessageBar shows either an error (type "E") or a warning (type "W").
Public Function isErrorOrWarningMsg()
    isErrorOrWarningMsg = (SAPSession.findById("wnd[0]/sbar").messagetype = "E" Or SAPSession.findById("wnd[0]/sbar").messagetype = "W")
End Function

Public Function isErrorMsg()
    isErrorMsg = (SAPSession.findById("wnd[0]/sbar").messagetype = "E")
End Function

' Checks if a given view (which technical name must be the input parameter) is active.
' If the view is not active, the SAPsession.findById expression will lead to an error -> fixed with its surrounding error handler.
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
    
    SAPSession.findById("wnd[0]/usr/ctxtRC29N-MATNR").Text = objSheet.Cells(currentRow, COL_ENDMATERIAL_NUMBER).Text
    SAPSession.findById("wnd[0]/usr/ctxtRC29N-WERKS").Text = objSheet.Cells(currentRow, COL_PLANT).Text
    SAPSession.findById("wnd[0]/usr/ctxtRC29N-STLAN").Text = "3"
    SAPSession.findById("wnd[0]").sendVKey 0
    If isErrorMsg() Then
        Call SetErrorMsg(currentRow, 20, SAPSession.findById("wnd[0]/sbar").Text)
        Call SetErrorMsg(currentRow, 20, "Aborted")
        Exit Sub
    End If
    SAPSession.findById("wnd[0]").sendVKey 0
    
    
    ' check if there is a separate popup window
    If SAPSession.ActiveWindow.Type = POPUP_WINDOW Then
        SAPSession.findById("wnd[1]").sendVKey 0
    End If
    ' Check if there are multiple alternatives
    If isViewActiveWithinWindow("wnd[0]/usr/tblSAPLCSDITCALT") Then
        Call SetErrorMsg(currentRow, COL_BOM_LOG, "Aborted! There are multiple BOM alternatives for this article!")
        Exit Sub
    End If
    ' Set qty in BOM header
    SAPSession.findById("wnd[0]/tbar[1]/btn[6]").press
    SAPSession.findById("wnd[0]/usr/tabsTS_HEAD/tabpKHPT/ssubSUBPAGE:SAPLCSDI:1110/txtRC29K-BMENG").Text = objSheet.Cells(currentRow, COL_BASE_QTY)
    SAPSession.findById("wnd[0]/tbar[1]/btn[5]").press
End Sub

' Deletes all component rows from the (previous) BOM
Private Sub DeletePreviousPositionRowsFromBom()
    SAPSession.findById("wnd[0]/tbar[1]/btn[27]").press
    SAPSession.findById("wnd[0]/tbar[1]/btn[14]").press
    SAPSession.findById("wnd[0]").sendVKey 0
    SAPSession.findById("wnd[1]/usr/btnSPOP-OPTION1").press
End Sub

Private Sub CopyPositionDataFromExcelToBom(currentRow)
    SAPSession.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-POSNR[0,0]").Text = objSheet.Cells(currentRow, COL_POSITION_NO).Text
    SAPSession.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-IDNRK[2,0]").Text = objSheet.Cells(currentRow, COL_POSITION_MATERIAL_NO).Text
    SAPSession.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-MENGE[4,0]").Text = Round(objSheet.Cells(currentRow, COL_POSITION_MATERIAL_QTY), 2)
    SAPSession.findById("wnd[0]").sendVKey 0
End Sub

Private Sub AddClassificationForPosition(currentRow, absRow)
    'msgbox("position " & absoluteRow)
    Length = Round(objSheet.Cells(currentRow, COL_POSITION_INDIVIDUAL_QTY), 2)
    SAPSession.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").getAbsoluteRow(absRow).Selected = True
    SAPSession.findById("wnd[0]/mbar/menu[3]/menu[4]").Select
    SAPSession.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB2").Select
    SAPSession.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB2/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,3]").Text = Length
    SAPSession.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB2/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,6]").Text = "RG"
    'msgbox("reached")
    SAPSession.findById("wnd[0]").sendVKey 0
    SAPSession.findById("wnd[0]/tbar[1]/btn[8]").press
    SAPSession.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").getAbsoluteRow(absRow).Selected = False
End Sub

Private Sub AddDetailScreenBasicData(currentRow, absRow)
    SAPSession.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").getAbsoluteRow(absRow).Selected = True
    SAPSession.findById("wnd[0]").sendVKey 7
    On Error Resume Next
        SAPSession.findById("wnd[0]/usr/tabsTS_ITEM/tabpPHPT").Select
    On Error GoTo 0
    SAPSession.findById("wnd[0]/usr/tabsTS_ITEM/tabpPHPT/ssubSUBPAGE:SAPLCSDI:0830/txtRC29P-AUSCH").Text = objSheet.Cells(currentRow, COL_COMPONENT_SCRAP).Text
    SAPSession.findById("wnd[0]/usr/tabsTS_ITEM/tabpPHPT/ssubSUBPAGE:SAPLCSDI:0830/chkRC29P-FMENG").Selected = (UCase(Trim(objSheet.Cells(currentRow, COL_FIXED_QTY).Text)) = "X")
    SAPSession.findById("wnd[0]/tbar[0]/btn[3]").press
    SAPSession.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").getAbsoluteRow(absRow).Selected = False
End Sub

Private Sub AddDetailScreenStatusLongtextData(currentRow, absRow)
    SAPSession.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").getAbsoluteRow(absRow).Selected = True
    SAPSession.findById("wnd[0]").sendVKey 7
    On Error Resume Next
        SAPSession.findById("wnd[0]/usr/tabsTS_ITEM/tabpPDAT").Select
    On Error GoTo 0
    SAPSession.findById("wnd[0]/usr/tabsTS_ITEM/tabpPDAT/ssubSUBPAGE:SAPLCSDI:0840/ctxtRC29P-SANKA").Text = objSheet.Cells(currentRow, COL_COSTING_RELEVANCY).Text
    SAPSession.findById("wnd[0]/usr/tabsTS_ITEM/tabpPDAT/ssubSUBPAGE:SAPLCSDI:0840/txtRC29P-POTX1").Text = objSheet.Cells(currentRow, COL_TEXT_ONE).Text
    SAPSession.findById("wnd[0]/usr/tabsTS_ITEM/tabpPDAT/ssubSUBPAGE:SAPLCSDI:0840/txtRC29P-POTX2").Text = objSheet.Cells(currentRow, COL_TEXT_TWO).Text
    SAPSession.findById("wnd[0]/tbar[0]/btn[3]").press
    SAPSession.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").getAbsoluteRow(absRow).Selected = False
End Sub

