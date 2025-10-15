VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ServoCableSelectForm 
   Caption         =   "Select cable"
   ClientHeight    =   3480
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7365
   OleObjectBlob   =   "Project FilesServoCableSelectForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "ServoCableSelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- UserForm: ServoCableSelectForm ---
Option Explicit

Public SelectedCableMaterial As String

Private Sub Label2_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim selectedProduct As String

    Me.SelectedCableMaterial = ""

    ' ? Corrected source
    Set ws = ThisWorkbook.Sheets("1. BOM Definition")
    Set tbl = ws.ListObjects("BOMDefinition")

    selectedProduct = Trim(ws.Range("F11").value)

    If selectedProduct = "" Then
        MsgBox "No product selected in cell F11.", vbExclamation
        Unload Me
        Exit Sub
    End If

    Me.cmbCableComponent.Clear

    ' ? Corrected field name: Product Number instead of Article Number
    For Each row In tbl.ListRows
        If Trim(row.Range(tbl.ListColumns("Product Number").Index).value) = selectedProduct Then
            Me.cmbCableComponent.AddItem row.Range(tbl.ListColumns("Material").Index).value
        End If
    Next row

    If Me.cmbCableComponent.ListCount > 0 Then
        Me.cmbCableComponent.ListIndex = 0
    Else
        MsgBox "No components found for the selected product.", vbExclamation
        Unload Me
    End If
End Sub

Private Sub cmbCableComponent_Change()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim selectedProduct As String
    Dim selectedMaterial As String

    Set ws = ThisWorkbook.Sheets("1. BOM Definition")
    Set tbl = ws.ListObjects("BOMDefinition")

    selectedProduct = Trim(ws.Range("F11").value)
    selectedMaterial = Me.cmbCableComponent.value

    Dim desc As String, dia As String
    desc = ""
    dia = ""

    For Each row In tbl.ListRows
        If Trim(row.Range(tbl.ListColumns("Product Number").Index).value) = selectedProduct And _
           Trim(row.Range(tbl.ListColumns("Material").Index).value) = selectedMaterial Then
            desc = row.Range(tbl.ListColumns("Material description").Index).value
            dia = row.Range(tbl.ListColumns("Cable diameter in mm").Index).value
            Exit For
        End If
    Next row

    If desc <> "" Or dia <> "" Then
        Me.lblInfo.Caption = "Description: " & desc & vbCrLf & "Cable Diameter: " & dia
    Else
        Me.lblInfo.Caption = "Component details not found."
    End If
End Sub


Private Sub btnOK_Click()
    If cmbCableComponent.ListIndex = -1 Then
        MsgBox "Please select a cable component.", vbExclamation
    Else
        Debug.Print "Selected cable: " & cmbCableComponent.value
        SelectedCableMaterial = cmbCableComponent.value
        Me.Hide
    End If
End Sub

Private Sub btnCancel_Click()
    SelectedCableMaterial = ""
    Unload Me
End Sub

