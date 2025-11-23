VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================
' MODULE:       Configuration (Class Module)
' DESCRIPTION:  Reads settings from a dedicated Excel Table to eliminate hardcoding.
' USAGE:        Dim Config As New Configuration
'               myValue = Config.Item("MySettingKey")
' ==============================================================================
Option Explicit

Private pSettings As Object ' Scripting.Dictionary to hold key-value pairs

Private Sub Class_Initialize()
    Set pSettings = CreateObject("Scripting.Dictionary")
    pSettings.CompareMode = vbTextCompare ' Case-insensitive keys
    LoadSettings
End Sub

' Loads data from the "ConfigTable" on the "Global Variables" sheet
Private Sub LoadSettings()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim r As Long
    Dim key As String, val As Variant
    
    ' 1. Safely attempt to find the table
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    Set tbl = ws.ListObjects("tbl_config")
    On Error GoTo 0
    
    ' 2. Graceful exit if setup is missing
    If tbl Is Nothing Then
        Debug.Print "CRITICAL: 'ConfigTable' not found on 'Global Variables' sheet."
        Exit Sub
    End If
    
    ' 3. Read table into Dictionary (assuming Col 1 = Key, Col 2 = Value)
    If Not tbl.DataBodyRange Is Nothing Then
        Dim data As Variant
        data = tbl.DataBodyRange.Value2 ' Faster than looping range objects
        
        For r = 1 To UBound(data, 1)
            key = Trim(CStr(data(r, 1)))
            val = data(r, 2)
            
            If Len(key) > 0 Then
                If pSettings.exists(key) Then
                    pSettings(key) = val ' Overwrite if duplicate
                Else
                    pSettings.Add key, val
                End If
            End If
        Next r
    End If
End Sub

' Public accessor for settings
Public Property Get Item(key As String) As Variant
    If pSettings.exists(key) Then
        Item = pSettings(key)
    Else
        ' Log missing key to Immediate Window for debugging
        Debug.Print "WARNING: Config key '" & key & "' not found."
        Item = "" ' Return empty string to prevent crash
    End If
End Property

