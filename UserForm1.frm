VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    ' =========================================================================
    ' ===          Excel File Copier and Data Processor                     ===
    ' ===   This macro is triggered by a button on a UserForm. It copies    ===
    ' ===   the current workbook 'n' times based on user input, modifies   ===
    ' ===   a specific cell in each copy, and prepares a JSON string.       ===
    ' =========================================================================

    ' --- Variable Declaration ---
    Dim n As Long ' The number of copies to create, from the TextBox
    Dim i As Long ' Loop counter
    
    Dim originalWorkbook As Workbook
    Dim newWorkbook As Workbook
    Dim targetSheet As Worksheet
    
    Dim originalPath As String      ' e.g., "https://yourcompany.sharepoint.com/sites/YourSite/Shared Documents/"
    Dim originalFullName As String  ' Full path and name of the original file
    Dim baseName As String          ' File name without extension
    Dim fileExt As String           ' File extension (e.g., ".xlsm")
    Dim newFullName As String       ' Full path and name for the new copy
    
    Dim originalC2Value As String
    Dim newC2Value As String
    Dim jsonString As String
    
    ' --- Settings for Performance ---
    ' Turn off screen updating to make the macro run faster and smoother
    application.ScreenUpdating = False
    ' Disable events to prevent other macros from firing unexpectedly
    application.EnableEvents = False

    ' --- Input Validation ---
    ' Check if the input in TextBox1 is a valid number.
    If Not IsNumeric(Me.TextBox1.Value) Then
        MsgBox "Error: Please enter a valid number in the textbox.", vbCritical, "Invalid Input"
        GoTo Cleanup ' Jump to the cleanup section
    End If
    
    ' Convert the text input to a number.
    n = CLng(Me.TextBox1.Value)
    
    ' Check if the number is positive.
    If n <= 0 Then
        MsgBox "Error: The number of copies must be greater than zero.", vbCritical, "Invalid Number"
        GoTo Cleanup ' Jump to the cleanup section
    End If

    ' --- Main Process ---
    On Error GoTo ErrorHandler ' If any error occurs from now on, jump to the error handler
    
    ' Set a reference to the workbook containing this code.
    Set originalWorkbook = ThisWorkbook
    
    ' --- Get Original File Information ---
    originalFullName = originalWorkbook.FullName
    originalPath = originalWorkbook.Path & application.PathSeparator ' PathSeparator adds '\' or '/' as needed.
    
    ' Separate the file name from its extension.
    baseName = Left(originalWorkbook.name, InStrRev(originalWorkbook.name, ".") - 1)
    fileExt = Mid(originalWorkbook.name, InStrRev(originalWorkbook.name, "."))
    
    ' Get the original value from cell C2.
    On Error Resume Next ' Temporarily ignore errors
    Set targetSheet = originalWorkbook.Sheets("0. ProjectData")
    On Error GoTo ErrorHandler ' Restore the main error handler
    
    If targetSheet Is Nothing Then
        MsgBox "Error: Sheet '0. ProjectData' could not be found in the original file.", vbCritical, "Sheet Not Found"
        GoTo Cleanup
    End If
    originalC2Value = targetSheet.Range("C2").Value
    Set targetSheet = Nothing ' Clear the object reference
    
    ' --- The Loop to Create Copies ---
    For i = 1 To n
        ' Construct the new file name with the suffix "_i"
        newFullName = originalPath & baseName & "_" & i & fileExt
        
        ' 1. Create the copy. SaveCopyAs is reliable for SharePoint locations.
        originalWorkbook.SaveCopyAs fileName:=newFullName
        
        ' 2. Open the newly created workbook to modify it.
        Set newWorkbook = Workbooks.Open(newFullName)
        
        ' 3. Find the sheet "0. ProjectData" in the new copy.
        On Error Resume Next ' Temporarily ignore errors
        Set targetSheet = newWorkbook.Sheets("0. ProjectData")
        On Error GoTo ErrorHandler ' Restore the main error handler
        
        If Not targetSheet Is Nothing Then
            ' 4. Create the new value for cell C2 and update the cell.
            newC2Value = originalC2Value & "_" & i
            targetSheet.Range("C2").Value = newC2Value
            
            ' 5. Build the JSON string. The double quotes ("") are used to include a quote inside a string.
            jsonString = "{""RFQID"": """ & newC2Value & """}"
            
            ' 6. Display the JSON string in a message box as requested.
            ' In a real scenario, you would replace this with your API call.
            MsgBox "JSON for iteration " & i & " (File: " & newWorkbook.name & "):" & vbCrLf & jsonString, vbInformation, "JSON Output"
            
        Else
            ' If the sheet was not found in the new copy, show a warning.
            MsgBox "Warning: Sheet '0. ProjectData' not found in copy: " & newWorkbook.name, vbExclamation, "Sheet Missing"
        End If
        
        ' Save changes and close the new workbook.
        newWorkbook.Close SaveChanges:=True
        Set newWorkbook = Nothing ' Release the workbook object
        Set targetSheet = Nothing ' Release the sheet object
    Next i
    
    ' --- Final Success Message ---
    MsgBox "Process completed successfully." & vbCrLf & _
           "A total of " & n & " file(s) have been created." & vbCrLf & vbCrLf & _
           "You may now want to delete this original template file.", vbInformation, "Success"
    
' --- Cleanup and Exit Point ---
Cleanup:
    ' Restore application settings to normal
    application.ScreenUpdating = True
    application.EnableEvents = True
    ' Release any object variables that might still be held
    Set newWorkbook = Nothing
    Set originalWorkbook = Nothing
    Set targetSheet = Nothing
    Exit Sub ' Exit the subroutine gracefully

' --- Error Handling Section ---
ErrorHandler:
    ' Display a descriptive error message.
    MsgBox "An unexpected error occurred." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Description: " & Err.description, vbCritical, "Macro Error"
    ' Jump to the cleanup section to ensure settings are restored.
    GoTo Cleanup

End Sub


