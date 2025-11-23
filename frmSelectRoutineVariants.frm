VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectRoutineVariants 
   Caption         =   "Routine setting for variants"
   ClientHeight    =   5340
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   14625
   OleObjectBlob   =   "frmSelectRoutineVariants.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmSelectRoutineVariants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private backing variables
Private pVariantNames() As String
Private pBaseProduct As String
Private pNumVariants As Long

' -------------------------------
' Property for VariantNames (Array)
' -------------------------------
Public Property Let VariantNames(ByRef Value() As String)
    pVariantNames = Value
End Property

Public Property Get VariantNames() As String()
    VariantNames = pVariantNames
End Property

' -------------------------------
' Property for BaseProduct
' -------------------------------
Public Property Let baseProduct(ByVal Value As String)
    pBaseProduct = Value
End Property

Public Property Get baseProduct() As String
    baseProduct = pBaseProduct
End Property

' -------------------------------
' Property for NumVariants
' -------------------------------
Public Property Let NumVariants(ByVal Value As Long)
    pNumVariants = Value
End Property

Public Property Get NumVariants() As Long
    NumVariants = pNumVariants
End Property

' -----------------------------------------------------------
' InitializeForm: Initializes the ListView for base operations
' -----------------------------------------------------------
Public Sub InitializeForm()
    Dim wsRoutines As Worksheet
    Dim tblRoutines As ListObject
    Dim routineRow As ListRow
    Dim listItem As listItem

    ' Validate BaseProduct
    If pBaseProduct = "" Then
        MsgBox "Error: Base Product not passed to the form.", vbCritical
        Unload Me
        Exit Sub
    End If

    Debug.Print "Base Product in InitializeForm: "; pBaseProduct

    ' Set worksheet and table
    Set wsRoutines = ThisWorkbook.Sheets("2. Routines")
    Set tblRoutines = wsRoutines.ListObjects("SelectedRoutines")

    ' Configure ListView
    With Me.OperationsListView
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .columnHeaders.Clear
        .columnHeaders.Add , , "Macrophase", 120
        .columnHeaders.Add , , "Microphase", 120
        .columnHeaders.Add , , "Material", 100
        .columnHeaders.Add , , "Machine", 100
        .columnHeaders.Add , , "Number of Operations", 100
    End With

    ' Populate ListView with filtered operations for the base Product
    For Each routineRow In tblRoutines.ListRows
        If CStr(routineRow.Range(tblRoutines.ListColumns("Product Number").Index).Value) = CStr(pBaseProduct) Then
            Set listItem = Me.OperationsListView.ListItems.Add(, , routineRow.Range(tblRoutines.ListColumns("Macrophase").Index).Value)
            listItem.SubItems(1) = routineRow.Range(tblRoutines.ListColumns("Microphase").Index).Value
            listItem.SubItems(2) = routineRow.Range(tblRoutines.ListColumns("Material").Index).Value
            listItem.SubItems(3) = routineRow.Range(tblRoutines.ListColumns("Machine").Index).Value
            listItem.SubItems(4) = routineRow.Range(tblRoutines.ListColumns("Number of operations").Index).Value
        End If
    Next routineRow

    ' If no operations found, notify the user and close the form
    If Me.OperationsListView.ListItems.Count = 0 Then
        MsgBox "No operations found for the selected base Product: " & pBaseProduct, vbInformation
        Unload Me
    End If
End Sub

' -----------------------------------------------------------
' btnSelectRoutine_Click: Triggers variant creation process
' -----------------------------------------------------------
Private Sub btnSelectRoutine_Click()
    Dim selectedMacrophase As String, selectedMicrophase As String
    Dim selectedMaterial As String, selectedMachine As String
    Dim selectedOperations As Double
    Dim wsRoutines As Worksheet
    Dim tblRoutines As ListObject
    Dim routineRow As ListRow
    Dim baseOperationQuantity As Double
    Dim variantQuantities() As Double
    Dim i As Long

    ' Ensure a routine is selected
    If Me.OperationsListView.selectedItem Is Nothing Then
        MsgBox "Please select a routine to create variants.", vbExclamation
        Exit Sub
    End If

    ' Retrieve selected routine details
    selectedMacrophase = Me.OperationsListView.selectedItem.Text
    selectedMicrophase = Me.OperationsListView.selectedItem.SubItems(1)
    selectedMaterial = Me.OperationsListView.selectedItem.SubItems(2)
    selectedMachine = Me.OperationsListView.selectedItem.SubItems(3)
    selectedOperations = CDbl(Me.OperationsListView.selectedItem.SubItems(4))

    ' Debugging: Print selected details
    Debug.Print "Selected Macrophase: " & selectedMacrophase
    Debug.Print "Selected Microphase: " & selectedMicrophase
    Debug.Print "Selected Material: " & selectedMaterial
    Debug.Print "Selected Machine: " & selectedMachine
    Debug.Print "Selected Operations: " & selectedOperations

    ' Locate the base operation in the SelectedRoutines table
    Set wsRoutines = ThisWorkbook.Sheets("2. Routines")
    Set tblRoutines = wsRoutines.ListObjects("SelectedRoutines")

    For Each routineRow In tblRoutines.ListRows
        If CStr(routineRow.Range(tblRoutines.ListColumns("Product Number").Index).Value) = CStr(Me.baseProduct) _
           And CStr(routineRow.Range(tblRoutines.ListColumns("Macrophase").Index).Value) = selectedMacrophase _
           And CStr(routineRow.Range(tblRoutines.ListColumns("Microphase").Index).Value) = selectedMicrophase _
           And CStr(routineRow.Range(tblRoutines.ListColumns("Material").Index).Value) = selectedMaterial _
           And CStr(routineRow.Range(tblRoutines.ListColumns("Machine").Index).Value) = selectedMachine _
           And CDbl(routineRow.Range(tblRoutines.ListColumns("Number of operations").Index).Value) = selectedOperations Then
            baseOperationQuantity = routineRow.Range(tblRoutines.ListColumns("Number of operations").Index).Value
            Exit For
        End If
    Next routineRow

    ' Debugging: Confirm if routine was found
    If routineRow Is Nothing Then
        MsgBox "The selected routine was not found.", vbExclamation
        Debug.Print "Routine not found!"
        Exit Sub
    End If

    Debug.Print "Base Operation Quantity: " & baseOperationQuantity

    ' Prompt for variant quantities
    ReDim variantQuantities(1 To NumVariants)
    For i = 1 To NumVariants
        Do
            variantQuantities(i) = application.InputBox("Enter quantity for Variant " & i & ":", "Quantity Input", Type:=1)
            If IsNumeric(variantQuantities(i)) And variantQuantities(i) > 0 Then Exit Do
            MsgBox "Please enter a valid positive number.", vbExclamation
        Loop
    Next i

    ' Call the variant creation procedure
    Call CreateRoutineVariants(tblRoutines, routineRow, Me.baseProduct, Me.NumVariants, variantQuantities, baseOperationQuantity)

    ' Notify user and close form
    MsgBox "Variants created successfully!", vbInformation
    Me.Hide
End Sub

Private Sub CreateRoutineVariants(tblRoutines As ListObject, routineRow As ListRow, _
                                  baseProduct As String, NumVariants As Long, _
                                  variantQuantities() As Double, baseOperationQuantity As Double)
    Dim newRoutineRow As ListRow
    Dim i As Long, j As Long
    Dim sourceCell As Range, targetCell As Range
    Dim variantName As String
    Dim allVariants() As String
    allVariants = Me.VariantNames


    ' Create variants by duplicating the selected operation
    For i = 1 To NumVariants
        variantName = allVariants(i)
        
        ' Row 1: Negative quantity
        Set newRoutineRow = tblRoutines.ListRows.Add
        For j = 1 To tblRoutines.ListColumns.Count
            Set sourceCell = routineRow.Range(j)
            Set targetCell = newRoutineRow.Range(j)

            If tblRoutines.ListColumns(j).name = "Product Number" Then
                targetCell.Value = variantName
            ElseIf tblRoutines.ListColumns(j).name = "Number of operations" Then
                targetCell.Value = -baseOperationQuantity
            ElseIf tblRoutines.ListColumns(j).name = "Variant of" Then
                targetCell.Value = baseProduct
            ElseIf Not sourceCell.HasFormula Then
                targetCell.Value = sourceCell.Value
           End If
        Next j

        ' Row 2: Positive quantity
        Set newRoutineRow = tblRoutines.ListRows.Add
        For j = 1 To tblRoutines.ListColumns.Count
            Set sourceCell = routineRow.Range(j)
            Set targetCell = newRoutineRow.Range(j)

            If tblRoutines.ListColumns(j).name = "Product Number" Then
                targetCell.Value = variantName
            ElseIf tblRoutines.ListColumns(j).name = "Number of operations" Then
                targetCell.Value = variantQuantities(i)
            ElseIf tblRoutines.ListColumns(j).name = "Variant of" Then
                targetCell.Value = baseProduct
            ElseIf Not sourceCell.HasFormula Then
                targetCell.Value = sourceCell.Value
            End If
        Next j
    Next i
End Sub


