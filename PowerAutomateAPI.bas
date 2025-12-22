Attribute VB_Name = "PowerAutomateAPI"
Public Sub SendFullyValidatedRFQ()
    Dim wsValidation As Worksheet
    Dim tblProjectData As ListObject
    Dim wsProjectData As Worksheet
    Dim DateValue As String
    Dim AdditionalData As Object
    Set AdditionalData = CreateObject("Scripting.Dictionary")
    Dim RFQNumber As String
    Dim AOV As Double
    Dim NumberOfLAPPComponents As Long, NumberOfThirdPartyComponents As Long
    Dim ValueOfLAPPComponents As Double, ValueOfThirdPartyComponents As Double, OverstockPercentageOfMargin As Double
    Dim OverstockValue As Double
    Dim validationMessage As String
    Dim commentsText As String   ' <<< NEW

    ' Set references
    Set wsValidation = ThisWorkbook.Sheets("3. Clarification Validation")
    Set wsProjectData = ThisWorkbook.Sheets("0. ProjectData")
    Set tblProjectData = wsProjectData.ListObjects("ProjectData")

    ' Get the validation message from cell J7
    validationMessage = wsValidation.Range("J7").value

    ' Ensure RFQ is valid before proceeding
    If validationMessage <> "All Products verified!" Then
        MsgBox "The RFQ is not valid. Please ensure all Products are verified before proceeding. After validating the components, click this button again to send to the Funnel file", vbExclamation
        Call ValidateAllComponentsAndProducts
        Exit Sub
    End If

    ' Retrieve relevant data from the ProjectData table
    On Error Resume Next
    RFQNumber = tblProjectData.ListRows(1).Range(tblProjectData.ListColumns("RFQ Number (CRM Opportunity)").Index).value
    AOV = tblProjectData.ListRows(1).Range(tblProjectData.ListColumns("Offered volume p.a. in EUR").Index).value
    NumberOfLAPPComponents = tblProjectData.ListRows(1).Range(tblProjectData.ListColumns("Number of Lapp components").Index).value
    NumberOfThirdPartyComponents = tblProjectData.ListRows(1).Range(tblProjectData.ListColumns("Number of third party components").Index).value
    ValueOfLAPPComponents = tblProjectData.ListRows(1).Range(tblProjectData.ListColumns("Value of Lapp components").Index).value
    ValueOfThirdPartyComponents = tblProjectData.ListRows(1).Range(tblProjectData.ListColumns("Value of third party components").Index).value
    OverstockValue = tblProjectData.ListRows(1).Range(tblProjectData.ListColumns("Overstock value").Index).value
    OverstockPercentageOfMargin = tblProjectData.ListRows(1).Range(tblProjectData.ListColumns("Overstock pct of margin").Index).value
    On Error GoTo 0

    ' Validate the required data exists
    If IsEmpty(RFQNumber) Or IsEmpty(AOV) Then
        MsgBox "Required data is missing in the ProjectData table. Please ensure all fields are filled correctly.", vbCritical
        Exit Sub
    End If

    ' Refresh & store Comments by Product before sending
    UpdateCommentsByProductInProjectData commentsText   ' <<< NEW

    ' Update the status and add the current UTC time
    tblProjectData.ListRows(1).Range(tblProjectData.ListColumns("Status").Index).value = "Full RFQ Validated"
    DateValue = GetISOTimestamp() ' Current UTC time
    tblProjectData.ListRows(1).Range(tblProjectData.ListColumns("RFQ valuation completion time").Index).value = DateValue

    ' Prepare additional data for Power Automate
    AdditionalData.Add "AOV", AOV
    AdditionalData.Add "NumberOfLAPPComponents", NumberOfLAPPComponents
    AdditionalData.Add "NumberOfThirdPartyComponents", NumberOfThirdPartyComponents
    AdditionalData.Add "ValueOfLAPPComponents", ValueOfLAPPComponents
    AdditionalData.Add "ValueOfThirdPartyComponents", ValueOfThirdPartyComponents
    AdditionalData.Add "OverstockValue", OverstockValue
    AdditionalData.Add "OverstockPctOfMargin", OverstockPercentageOfMargin
    AdditionalData.Add "CommentsByProduct", commentsText   ' <<< NEW

    ' Send data to Power Automate
    Call SendToPowerAutomate("Full RFQ Validated", AdditionalData)

    ' Provide feedback to the user
    MsgBox "RFQ validated successfully and data sent to Funnel File.", vbInformation
End Sub


Private Sub CheckAndSendRFQ()
    Dim wsValidation As Worksheet
    Dim validationMessage As String

    ' Set reference to the Validation sheet
    Set wsValidation = ThisWorkbook.Sheets("3. Clarification Validation")

    ' Get the value of cell J7
    validationMessage = wsValidation.Range("J7").value

    ' Check if the RFQ is valid
    If validationMessage = "All Products verified!" Then
        ' If valid, call the function to send data to Power Automate
        Call SendFullyValidatedRFQ
    Else
        ' If not valid, show an error message
        MsgBox "The RFQ is not valid. Please ensure all Products are verified before sending.", vbExclamation
        wsValidation.Activate
    End If
End Sub
Private Sub SendToPowerAutomate(status As String, Optional AdditionalData As Object)
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim tbl As ListObject
    Dim RFQID As String, DateValue As String, RFQNumber As String
    Dim wsOutput As Worksheet
    Dim key As Variant
    Dim userID As String
    Dim engineer As String
    Dim purchaser As String
    Dim comment As String
    Dim PlantInternalID As String
    
    userID = application.UserName
    url = ActiveWorkbook.Sheets("Global Variables").Cells(7, 2).Text
    
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    Set tbl = ThisWorkbook.Sheets("0. ProjectData").ListObjects("ProjectData")
    
    ' Get values with null handling
    RFQNumber = SafeString(tbl.ListRows(1).Range(tbl.ListColumns("RFQ Number (CRM Opportunity)").Index).value)
    engineer = SafeString(tbl.ListRows(1).Range(tbl.ListColumns("Calculation Engineer").Index).value)
    purchaser = SafeString(tbl.ListRows(1).Range(tbl.ListColumns("Purchasing Responsible").Index).value)
    comment = SafeString(tbl.ListRows(1).Range(tbl.ListColumns("Comment from plant (optional)").Index).value)
    PlantInternalID = SafeString(tbl.ListRows(1).Range(tbl.ListColumns("Internal ID").Index).value)
    
    DateValue = GetISOTimestamp()
    
    ' Build JSON using array (cleaner and safer)
    Dim jsonParts() As String
    ReDim jsonParts(0 To 7)
    
    jsonParts(0) = """RFQID"": """ & JsonEscape(RFQNumber) & """"
    jsonParts(1) = """PlantInternalID"": """ & JsonEscape(PlantInternalID) & """"
    jsonParts(2) = """Status"": """ & JsonEscape(status) & """"
    jsonParts(3) = """Date"": """ & DateValue & """"
    jsonParts(4) = """CalcEngineer"": """ & JsonEscape(engineer) & """"
    jsonParts(5) = """PurResponsible"": """ & JsonEscape(purchaser) & """"
    jsonParts(6) = """CommentFromPlant"": """ & JsonEscape(comment) & """"
    jsonParts(7) = """UserID"": """ & JsonEscape(userID) & """"
    
    requestBody = "{" & Join(jsonParts, ", ")
    
    ' Handle CommentsByProduct
    Dim hasCBP As Boolean: hasCBP = False
    Dim cbp As String
    
    If Not AdditionalData Is Nothing Then
        For Each key In AdditionalData.Keys
            If CStr(key) = "CommentsByProduct" Then
                hasCBP = True
                cbp = SafeString(AdditionalData(key))
                requestBody = requestBody & ", ""CommentsByProduct"": """ & JsonEscape(cbp) & """"
                Exit For
            End If
        Next key
    End If
    
    ' If not in AdditionalData, check ProjectData table
    If Not hasCBP Then
        Dim colCmt As Long
        On Error Resume Next
        colCmt = tbl.ListColumns("Comments by Product").Index
        On Error GoTo 0
        If colCmt > 0 And tbl.ListRows.Count > 0 Then
            cbp = SafeString(tbl.ListRows(1).Range(colCmt).value)
            If Len(cbp) > 0 Then
                requestBody = requestBody & ", ""CommentsByProduct"": """ & JsonEscape(cbp) & """"
            End If
        End If
    End If
    
    ' Add remaining AdditionalData (skip CommentsByProduct since we already handled it)
    If Not AdditionalData Is Nothing Then
        For Each key In AdditionalData.Keys
            If CStr(key) <> "CommentsByProduct" Then
                If IsNumeric(AdditionalData(key)) Then
                    requestBody = requestBody & ", """ & JsonEscape(CStr(key)) & """: " & FormatNumericForJson(AdditionalData(key))
                Else
                    requestBody = requestBody & ", """ & JsonEscape(CStr(key)) & """: """ & JsonEscape(SafeString(AdditionalData(key))) & """"
                End If
            End If
        Next key
    End If
    
    requestBody = requestBody & "}"
    
    ' DEBUG: Log the actual JSON being sent
    Debug.Print "=== JSON PAYLOAD ==="
    Debug.Print requestBody
    Debug.Print "===================="
    
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json; charset=utf-8"
        .send requestBody
    End With
    
    ' Enhanced error reporting
    If http.status = 200 Or http.status = 202 Then
        Set wsOutput = ThisWorkbook.Sheets("4. Sales Calculation (Internal)")
        wsOutput.Range("N1").value = "RFQ data sent to Funnel File at " & DateValue
    Else
        MsgBox "Error: " & http.status & " - " & http.statusText & vbCrLf & vbCrLf & _
               "Response: " & http.responseText & vbCrLf & vbCrLf & _
               "Payload sent: " & Left(requestBody, 500), vbCritical
        Debug.Print "ERROR RESPONSE: " & http.responseText
    End If
    
    Set http = Nothing
End Sub

' NEW: Safe string conversion
Private Function SafeString(ByVal value As Variant) As String
    If IsEmpty(value) Or IsNull(value) Then
        SafeString = ""
    Else
        SafeString = CStr(value)
    End If
End Function

' UPDATED: Better numeric formatting
Private Function FormatNumericForJson(ByVal value As Variant) As String
    Dim result As String
    result = CStr(value)
    ' Replace comma with period regardless of locale
    result = Replace(result, ",", ".")
    FormatNumericForJson = result
End Function

' UPDATED: Enhanced JsonEscape
Private Function JsonEscape(str As String) As String
    If Len(str) = 0 Then
        JsonEscape = ""
        Exit Function
    End If
    
    Dim temp As String
    temp = str
    
    ' Order matters - backslash must be first
    temp = Replace(temp, "\", "\\")
    temp = Replace(temp, """", "\""")
    temp = Replace(temp, vbCrLf, "\n")  ' Simplified to \n
    temp = Replace(temp, vbCr, "\n")
    temp = Replace(temp, vbLf, "\n")
    temp = Replace(temp, vbTab, "\t")
    ' Remove any control characters (ASCII < 32) except newline/tab
    Dim i As Long
    Dim cleanStr As String
    For i = 1 To Len(temp)
        Dim ch As String
        ch = Mid(temp, i, 1)
        If Asc(ch) >= 32 Or ch = vbLf Or ch = vbTab Then
            cleanStr = cleanStr & ch
        End If
    Next i
    
    JsonEscape = cleanStr
End Function

Public Sub AddFirstProduct()
    Dim tbl As ListObject
    Dim ws As Worksheet
    Dim DateValue As String

    ' Define the ProjectData table
    Set ws = ThisWorkbook.Sheets("0. ProjectData")
    Set tbl = ws.ListObjects("ProjectData")

    ' Update the status and add the current UTC time
    tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = "RFQ Calculation Started"
    DateValue = GetISOTimestamp()
    tbl.ListRows(1).Range(tbl.ListColumns("RFQ calculation start time").Index).value = DateValue
    ' Send to Power Automate
    Call SendToPowerAutomate("RFQ Calculation Started")
End Sub

Public Sub BOMRoutingCreated()
    Dim tbl As ListObject
    Dim ws As Worksheet
    Dim DateValue As String

    ' Define the ProjectData table
    Set ws = ThisWorkbook.Sheets("0. ProjectData")
    Set tbl = ws.ListObjects("ProjectData")
    Call ValidateAllComponentsAndProducts
    
    ' Update the status and add the current UTC time
    tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = "BOM&Routing Created"
    DateValue = GetISOTimestamp()
    tbl.ListRows(1).Range(tbl.ListColumns("RFQ BOM and Routings completion time").Index).value = DateValue

    ' Send to Power Automate
    Call SendToPowerAutomate("BOM&Routing Created")
End Sub

Public Sub StartCustomerClarification(roundNumber As Integer)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim wsData As Worksheet
    Dim DateValue As String

 
    Dim currentTime As String
    Dim statusMessage As String
    
    ' Define the ProjectData table
    Set wsData = ThisWorkbook.Sheets("0. ProjectData")
    Set tbl = wsData.ListObjects("ProjectData")
    
    ' Set the worksheet reference
    Set ws = ThisWorkbook.Sheets("3. Clarification Validation")
    
    ' Get the current time
    currentTime = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    
    ' Determine the status message and update the "Start" cell based on the round number
    Select Case roundNumber
        Case 1
            ws.Range("E8").value = currentTime
            statusMessage = "StartCustomerClarification Round 1"
            tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = "StartCustomerClarification Round 1"
        Case 2
            ws.Range("E11").value = currentTime
            statusMessage = "StartCustomerClarification Round 2"
            tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = "StartCustomerClarification Round 2"
        Case 3
            ws.Range("E14").value = currentTime
            statusMessage = "StartCustomerClarification Round 3"
            tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = "StartCustomerClarification Round 3"
    End Select
    
    ' Send the status message to Power Automate
    Call SendToPowerAutomate(statusMessage)
End Sub

Public Sub EndCustomerClarification(roundNumber As Integer)
    Dim ws As Worksheet
    Dim currentTime As String
    Dim statusMessage As String
    Dim tbl As ListObject
    Dim wsData As Worksheet
    
    ' Set the worksheet reference
    Set ws = ThisWorkbook.Sheets("3. Clarification Validation")
    
    ' Define the ProjectData table
    Set wsData = ThisWorkbook.Sheets("0. ProjectData")
    Set tbl = wsData.ListObjects("ProjectData")
    
    ' Get the current time
    currentTime = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    
    ' Determine the status message and update the "End" cell based on the round number
    Select Case roundNumber
        Case 1
            ws.Range("G8").value = currentTime
            statusMessage = "EndCustomerClarification Round 1"
            tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = "EndCustomerClarification Round 1"
        Case 2
            ws.Range("G11").value = currentTime
            statusMessage = "EndCustomerClarification Round 2"
            tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = "EndCustomerClarification Round 2"
        Case 3
            ws.Range("G14").value = currentTime
            statusMessage = "EndCustomerClarification Round 3"
            tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = "EndCustomerClarification Round 3"
    End Select
    
    ' Send the status message to Power Automate
    Call SendToPowerAutomate(statusMessage)
End Sub

Sub StartCustomerRound1()
    Call StartCustomerClarification(1)
End Sub

Sub EndCustomerRound1()
    Call EndCustomerClarification(1)
End Sub

Sub StartCustomerRound2()
    Call StartCustomerClarification(2)
End Sub

Sub EndCustomerRound2()
    Call EndCustomerClarification(2)
End Sub

Sub StartCustomerRound3()
    Call StartCustomerClarification(3)
End Sub

Sub EndCustomerRound3()
    Call EndCustomerClarification(3)
End Sub

Public Sub StartPurchasingClarification(roundNumber As Integer)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim wsData As Worksheet
    Dim currentTime As String
    Dim statusMessage As String
    
    ' Define the ProjectData table
    Set wsData = ThisWorkbook.Sheets("0. ProjectData")
    Set tbl = wsData.ListObjects("ProjectData")
    
    ' Set the worksheet reference
    Set ws = ThisWorkbook.Sheets("3. Clarification Validation")
    
    ' Get the current time
    currentTime = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    
    ' Determine the status message and update the "Start" cell based on the round number
    Select Case roundNumber
        Case 1
            ws.Range("E17").value = currentTime
            statusMessage = "StartPurchasingClarification Round 1"
            tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = "StartPurchasingClarification Round 1"
        Case 2
            ws.Range("E20").value = currentTime
            statusMessage = "StartPurchasingClarification Round 2"
            tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = "StartPurchasingClarification Round 2"
        Case 3
            ws.Range("E23").value = currentTime
            statusMessage = "StartPurchasingClarification Round 3"
            tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = "StartPurchasingClarification Round 3"
    End Select
    
    ' Send the status message to Power Automate
    Call SendToPowerAutomate(statusMessage)
End Sub

Public Sub EndPurchasingClarification(roundNumber As Integer)
    Dim ws As Worksheet
    Dim currentTime As String
    Dim statusMessage As String
    Dim tbl As ListObject
    Dim wsData As Worksheet
    
    ' Set the worksheet reference
    Set ws = ThisWorkbook.Sheets("3. Clarification Validation")
    
    ' Define the ProjectData table
    Set wsData = ThisWorkbook.Sheets("0. ProjectData")
    Set tbl = wsData.ListObjects("ProjectData")
    
    ' Get the current time
    currentTime = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    
    ' Determine the status message and update the "End" cell based on the round number
    Select Case roundNumber
        Case 1
            ws.Range("G17").value = currentTime
            statusMessage = "EndPurchasingClarification Round 1"
            tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = "EndPurchasingClarification Round 1"
        Case 2
            ws.Range("G20").value = currentTime
            statusMessage = "EndPurchasingClarification Round 2"
            tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = "EndPurchasingClarification Round 2"
        Case 3
            ws.Range("G23").value = currentTime
            statusMessage = "EndPurchasingClarification Round 3"
            tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = "EndPurchasingClarification Round 3"
    End Select
    
    ' Send the status message to Power Automate
    Call SendToPowerAutomate(statusMessage)
End Sub

Public Sub PauseDueToPriorityChange()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim wsData As Worksheet
    Dim currentTime As String
    Dim statusMessage As String

    ' Define the ProjectData table
    Set wsData = ThisWorkbook.Sheets("0. ProjectData")
    Set tbl = wsData.ListObjects("ProjectData")

    ' Set the worksheet reference
    Set ws = ThisWorkbook.Sheets("3. Clarification Validation")

    ' Get the current time
    currentTime = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")

    ' Set timestamp and status message
    ws.Range("E27").value = currentTime ' You can change the target cell as needed
    statusMessage = "Paused due to priority change"
    tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = statusMessage

    ' Send the status message to Power Automate
    Call SendToPowerAutomate(statusMessage)
End Sub

Public Sub ResumeAfterPriorityChange()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim wsData As Worksheet
    Dim currentTime As String
    Dim statusMessage As String

    ' Define the ProjectData table
    Set wsData = ThisWorkbook.Sheets("0. ProjectData")
    Set tbl = wsData.ListObjects("ProjectData")

    ' Set the worksheet reference
    Set ws = ThisWorkbook.Sheets("3. Clarification Validation")

    ' Get the current time
    currentTime = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")

    ' Set timestamp and status message
    ws.Range("G27").value = currentTime ' You can change the target cell as needed
    statusMessage = "Resumed after priority change"
    tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = statusMessage

    ' Send the status message to Power Automate
    Call SendToPowerAutomate(statusMessage)
End Sub

Public Sub RejectRFQ()
    Dim wsValidation As Worksheet
    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim rejectionReason As String
    Dim utcTime As String
    Dim localTimeDisplay As String
    Dim AdditionalData As Object

    ' Set references
    Set wsValidation = ThisWorkbook.Sheets("3. Clarification Validation")
    Set wsData = ThisWorkbook.Sheets("0. ProjectData")
    Set tbl = wsData.ListObjects("ProjectData")
    Set AdditionalData = CreateObject("Scripting.Dictionary")

    ' Get the rejection reason from cell L4
    rejectionReason = Trim(wsValidation.Range("L4").value)

    ' Check if reason is empty
    If rejectionReason = "" Then
        MsgBox "Please provide a reason for rejection in cell L4 before submitting.", vbExclamation
        wsValidation.Range("L4").Select
        Exit Sub
    End If

    ' Get current UTC and local time
    utcTime = Format(Now, "yyyy-mm-dd\Thh:nn:ss")
    localTimeDisplay = Format(Now, "yyyy-mm-dd hh:mm") ' display time (for G2)

    ' Update cell O2 with local time in display format
    wsValidation.Range("O2").value = localTimeDisplay

    ' Update status in ProjectData
    tbl.ListRows(1).Range(tbl.ListColumns("Status").Index).value = "rejected"

    ' Prepare additional data
    AdditionalData.Add "RejectionReason", rejectionReason

    ' Send to Power Automate
    Call SendToPowerAutomate("rejected", AdditionalData)

    ' Optional feedback
    MsgBox "RFQ has been marked as rejected and data sent to the RFQ Funnel.", vbInformation
End Sub



Sub StartPurchasingRound1()
    Call StartPurchasingClarification(1)
End Sub

Sub EndPurchasingRound1()
    Call EndPurchasingClarification(1)
End Sub

Sub StartPurchasingRound2()
    Call StartPurchasingClarification(2)
End Sub

Sub EndPurchasingRound2()
    Call EndPurchasingClarification(2)
End Sub

Sub StartPurchasingRound3()
    Call StartPurchasingClarification(3)
End Sub

Sub EndPurchasingRound3()
    Call EndPurchasingClarification(3)
End Sub

'================ Helpers: collect & store "Comments by Product" ================

' Try a few likely sheet names for the Sales Calculation sheet.
Private Function GetSalesCalcSheet() As Worksheet
    Dim names As Variant, n As Variant, ws As Worksheet
    names = Array("4. Sales Calculation (Internal)", "4.1. Sales Calculation (Internal)", "4.1 Sales Calculation (Internal)", "4.1")
    For Each n In names
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(n))
        On Error GoTo 0
        If Not ws Is Nothing Then Set GetSalesCalcSheet = ws: Exit Function
    Next n
    ' If not found, return Nothing
End Function

' Returns a multi-line string from Columns:
'   B = Product Number
'   F = Comment
' starting at row 15. Only non-empty comments are included:
'   <Product Number> - <Comment>
Public Function BuildCommentsByProduct() As String
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim pn As String, cmt As String
    Dim buf As String
    
    Set ws = GetSalesCalcSheet()
    If ws Is Nothing Then Exit Function
    
    Const FIRST_ROW As Long = 15
    lastRow = application.WorksheetFunction.Max( _
                    ws.Cells(ws.Rows.Count, "B").End(xlUp).row, _
                    ws.Cells(ws.Rows.Count, "F").End(xlUp).row)
    
    For r = FIRST_ROW To lastRow
        cmt = Trim$(CStr(ws.Cells(r, "F").value))
        If Len(cmt) > 0 Then
            pn = CStr(ws.Cells(r, "B").value)
            If buf <> "" Then buf = buf & vbCrLf
            buf = buf & pn & " - " & cmt
        End If
    Next r
    
    BuildCommentsByProduct = buf
End Function

' Writes the multi-line text into 0. ProjectData ? ProjectData[Comments by Product].
' Creates the column if missing.
Public Sub UpdateCommentsByProductInProjectData(Optional ByRef outText As String)
    Dim ws As Worksheet, tbl As ListObject
    Dim colIdx As Long
    Dim txt As String
    
    txt = BuildCommentsByProduct()
    outText = txt
    
    Set ws = ThisWorkbook.Sheets("0. ProjectData")
    On Error Resume Next
    Set tbl = ws.ListObjects("ProjectData")
    On Error GoTo 0
    If tbl Is Nothing Then Exit Sub
    
    On Error Resume Next
    colIdx = tbl.ListColumns("Comments by Product").Index
    On Error GoTo 0
    If colIdx = 0 Then
        colIdx = tbl.ListColumns.Count + 1
        tbl.ListColumns.Add Position:=colIdx
        tbl.ListColumns(colIdx).name = "Comments by Product"
    End If
    
    If tbl.ListRows.Count = 0 Then tbl.ListRows.Add
    tbl.ListRows(1).Range(tbl.ListColumns(colIdx).Index).value = txt
End Sub

Public Sub MakeCommentsByProduct()
    Dim txt As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim colIdx As Long
    Dim Target As Range
    
    ' Build "PN - Comment" lines (reads Sales sheet: Col B & F from row 15)
    txt = BuildCommentsByProduct()
    
    ' Write to 0. ProjectData ? ProjectData[Comments by Product]
    Set ws = ThisWorkbook.Sheets("0. ProjectData")
    Set tbl = ws.ListObjects("ProjectData")
    
    On Error Resume Next
    colIdx = tbl.ListColumns("Comments by Product").Index
    On Error GoTo 0
    If colIdx = 0 Then
        colIdx = tbl.ListColumns.Count + 1
        tbl.ListColumns.Add Position:=colIdx
        tbl.ListColumns(colIdx).name = "Comments by Product"
    End If
    
    If tbl.ListRows.Count = 0 Then tbl.ListRows.Add
    Set Target = tbl.DataBodyRange.Cells(1, colIdx)
    Target.value = txt
    
    ' Bring user to the updated cell
    ws.Activate
    Target.Select
End Sub

' Add this at the top of your module, after other Private Functions
Private Function GetISOTimestamp() As String
    ' Returns ISO 8601 timestamp with colons (not locale-dependent dots/commas)
    ' Format: 2025-12-08T15:07:22
    Dim dt As Date
    dt = Now
    GetISOTimestamp = Format(dt, "yyyy-mm-dd") & "T" & _
                      Right("00" & Hour(dt), 2) & ":" & _
                      Right("00" & Minute(dt), 2) & ":" & _
                      Right("00" & Second(dt), 2)
End Function


Public Sub SendGeneralStatusUpdate()
    Call SendToPowerAutomate("GeneralStatusUpdate")
End Sub
