VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadedDataSelectionForm 
   Caption         =   "Select component"
   ClientHeight    =   5340
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   17685
   OleObjectBlob   =   "LoadedDataSelectionForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "LoadedDataSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_selectedRowIndex As Long

' Property to get and set the selectedRowIndex
Public Property Get selectedRowIndex() As Long
    selectedRowIndex = m_selectedRowIndex
End Property

Public Property Let selectedRowIndex(Value As Long)
    m_selectedRowIndex = Value
End Property

Public Sub InitializeForm(matches As Range)
    Dim rowData As Range
    Dim i As Long, j As Long
    Dim tblGetExtendedMMD As ListObject
    Dim headerCell As Range
    Dim matchEntireRow As Range

    ' Set the worksheet and table
    Set tblGetExtendedMMD = ThisWorkbook.Sheets("ExtendedMMD").ListObjects("GetExtendedMMD")

    ' Clear existing items and headers in the ListView
    Me.lvwMatchingRows.ListItems.Clear
    Me.lvwMatchingRows.columnHeaders.Clear

    ' Set up the ListView
    With Me.lvwMatchingRows
        .View = lvwReport ' Set view to Report to show columns
        .Gridlines = True ' Enable gridlines
        .FullRowSelect = True ' Select the full row when clicked
        .HideColumnHeaders = False ' Show headers
    End With

    ' Add column headers based on the GetExtendedMMD table headers
    For Each headerCell In tblGetExtendedMMD.headerRowRange
        Me.lvwMatchingRows.columnHeaders.Add , , headerCell.Value, 100 ' Adjust column width as needed
    Next headerCell

    ' Populate ListView with data and associate each item with a row index
    Dim area As Range
    For Each area In matches.Areas
        For Each rowData In area.Rows
            ' Retrieve the entire row based on the match found in one cell
            Dim relativeRowIndex As Long
            relativeRowIndex = rowData.row - tblGetExtendedMMD.headerRowRange.row
            
            If relativeRowIndex >= 1 And _
               relativeRowIndex <= tblGetExtendedMMD.DataBodyRange.Rows.Count Then

                Set matchEntireRow = tblGetExtendedMMD.DataBodyRange.Rows(relativeRowIndex)

                ' Add the first cell of the row as the main item
                Dim listItem As listItem
                Set listItem = Me.lvwMatchingRows.ListItems.Add(, , matchEntireRow.Cells(1, 1).Value)
                listItem.Tag = matchEntireRow.row  ' Store the actual row index in the Tag

                ' Loop through remaining columns to add them as subitems
                For j = 2 To matchEntireRow.Columns.Count
                    listItem.SubItems(j - 1) = matchEntireRow.Cells(1, j).Value
                Next j
            End If
        Next rowData
    Next area

    ' If only one match is found, automatically select it
    If Me.lvwMatchingRows.ListItems.Count = 1 Then
        Me.lvwMatchingRows.ListItems(1).Selected = True ' Select the first and only item
        Me.selectedRowIndex = Me.lvwMatchingRows.ListItems(1).Tag ' Automatically set selectedRowIndex
        Me.Hide ' Close the form automatically
    Else
        ' Default to no selection if multiple items are shown
        selectedRowIndex = -1
        m_selectedRowIndex = -1
    End If
End Sub

Private Sub cmdOK_Click()
    If Me.lvwMatchingRows.selectedItem Is Nothing Then
        MsgBox "Please select a row.", vbExclamation
    Else
        ' Get the actual row index stored in the Tag
        m_selectedRowIndex = Me.lvwMatchingRows.selectedItem.Tag
        Me.Hide
    End If
End Sub

Private Sub cmdCancel_Click()
    Set selectedRow = Nothing  ' No selection
    m_selectedRowIndex = -1
    Me.Hide
End Sub

Public Sub ProcessSelectedMaterial(selectedMaterial As String)
    Dim tblGetExtendedMMD As ListObject
    ' Set the worksheet and table
    Set tblGetExtendedMMD = ThisWorkbook.Sheets("ExtendedMMD").ListObjects("GetExtendedMMD")
    
    Debug.Print "Processing material:", selectedMaterial  ' Debug line

    ' Assuming tblExtendedMMD is the table with your material data
    Dim foundRow As Range
    Set foundRow = tblGetExtendedMMD.ListColumns("Material").DataBodyRange.Find(selectedMaterial, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundRow Is Nothing Then
        Debug.Print "Material found in ExtendedMMD:", foundRow.Address
        ' Continue with processing, e.g., copying data from foundRow
    Else
        Debug.Print "No match found in ExtendedMMD for Material:", selectedMaterial
        MsgBox "No match found in LoadedData for Material: " & selectedMaterial, vbExclamation
    End If
End Sub


