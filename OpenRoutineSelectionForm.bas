Attribute VB_Name = "OpenRoutineSelectionForm"
Sub OpenRoutineForm()
    Dim myRoutineForm As RoutineForm
    Set myRoutineForm = New RoutineForm
    myRoutineForm.SetupForm ' This runs SetupForm even without preselection
    myRoutineForm.Show
End Sub

