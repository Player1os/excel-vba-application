Option Explicit
Option Private Module

Public Const vProjectName As String = "" ' TODO: Change.
Public Const vProjectPassword As String = "" ' TODO: Change.

Public Const vDefaultErrorMessage As String = "An unknown error had occurred, please contact the administrator." ' OPTIONAL: Change.

Public Sub InitializeState( _
    ByRef vConfigWorkbook As Workbook _
)
    ' OPTIONAL: Implement.
End Sub

Public Sub TerminateState()
    ' OPTIONAL: Implement.
End Sub

Public Function IsErrorMessageEnabled() As Boolean
    IsErrorMessageEnabled = True ' OPTIONAL: Change.
End Function

Public Function Execute( _
    ByRef vFullActionName As String, _
    ByRef vEventParameter As Variant _
) As Boolean
    Execute = True
    Select Case vFullActionName
        ' TODO: Add actions.
        Case Else
            Execute = False
    End Select
End Function
