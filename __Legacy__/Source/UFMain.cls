Option Explicit

' Requires MRuntime

' OPTIONAL: Add if required.
Private Sub UserForm_Activate()
    Call MRuntime.Execute(Me.Name, "Activate")
End Sub

' TODO: Add user form specific macros.
