Option Explicit

' Requires MRuntime

' OPTIONAL: Add if required.
Private Sub Worksheet_Change( _
    ByVal Target As Range _
)
    Call MRuntime.Execute(Me.CodeName, "Change", _
        vEventHandling:=WorksheetChange, vEventParameter:=Target)
End Sub

' OPTIONAL: Add if required.
Private Sub Worksheet_SelectionChange( _
    ByVal Target As Range _
)
    Call MRuntime.Execute(Me.CodeName, "SelectionChange", _
        vEventHandling:=WorksheetSelectionChange, vEventParameter:=Target)
End Sub

' OPTIONAL: Add sheet specific macros.
