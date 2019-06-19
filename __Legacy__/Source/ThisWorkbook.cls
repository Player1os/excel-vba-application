Option Explicit

' Requires MRuntime

' OPTIONAL: Add if required.
Private Sub Workbook_Activate()
    Call MRuntime.Execute(Me.CodeName, "Activate")
End Sub

' OPTIONAL: Add if required.
Private Sub Workbook_Deactivate()
    Call MRuntime.Execute(Me.CodeName, "Deactivate")
End Sub

' OPTIONAL: Remove vIsIgnored:=True if implemented.
Private Sub Workbook_Open()
    Call MRuntime.Execute(Me.CodeName, "Open", vEventHandling:=EEventHandling.vWorkbookOpen, _
        vIsIgnored:=True)
End Sub

' OPTIONAL: Remove vIsIgnored:=True if implemented.
Private Sub Workbook_BeforeClose( _
    ByRef Cancel As Boolean _
)
    Call MRuntime.Execute(Me.CodeName, "BeforeClose", vEventHandling:=EEventHandling.vWorkbookBeforeClose, _
        vEventParameter:=Cancel, vIsIgnored:=True)
End Sub

Public Sub Deploy()
    Call MRuntime.Execute(Me.CodeName, "Deploy", vIsDebugRequired:=True)
End Sub

Public Sub ExportProject()
    Call MRuntime.Execute(Me.CodeName, "ExportProject", vIsDebugRequired:=True)
End Sub

Public Sub OpenConfig()
    Call MRuntime.Execute(Me.CodeName, "OpenConfig", vIsDebugRequired:=True)
End Sub

Public Sub UnprotectProject()
    Call MRuntime.Execute(Me.CodeName, "UnprotectProject", vIsDebugRequired:=True)
End Sub

' OPTIONAL: Add global macros.
