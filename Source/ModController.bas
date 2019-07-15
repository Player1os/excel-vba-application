Attribute VB_Name = "ModController"
Option Explicit
Option Private Module

Public Sub Navigate( _
    ByRef vPath As String, _
    ByRef vParameters As Dictionary _
)
    ' Declare local variables.
    Dim vInnerHtml As String
    Dim vParameterKey As Variant

    ' Display the path and parameters on the loaded html page.
    vInnerHtml = "<h1>Path</h1>"
    vInnerHtml = vInnerHtml & "<p>" & vPath & "</p>"
    vInnerHtml = vInnerHtml & "<h1>Parameters</h1>"
    If vParameters.Count = 0 Then
        vInnerHtml = vInnerHtml & "<p><i>None</i></p>"
    Else
        vInnerHtml = vInnerHtml & "<ul>"
        For Each vParameterKey In vParameters.Keys()
            vInnerHtml = vInnerHtml & "<li><b>" & CStr(vParameterKey) & "</b>: " & vParameters(vParameterKey) & "</li>"
        Next
        vInnerHtml = vInnerHtml & "</ul>"
    End If
    Call ThisUserForm.SetInnerHtml(vInnerHtml)

    ' Output the current timestamp to a file.
    With Runtime.FileSystemObject()
        With .CreateTextFile(.BuildPath(ThisWorkbook.Path, "Output.log"))
            Call .WriteLine(CStr(Now))
            Call .Close
        End With
    End With

    ' If the excel application instance is running in the background, trigger an early exit.
    If Not Application.Visible Then
        Call ThisUserForm.Hide
    End If
End Sub

'''''''''''''''''''''''
'                     '
' Procedure Template: '
'                     '
'''''''''''''''''''''''

' Public [Sub | Function] ProcedureName()
'     ' Declare local variables.
'     ' TODO: Implement.

'     ' Setup error handling.
'     On Error GoTo HandleError:

'     ' Allocate resources.
'     ' TODO: Implement.

'     ' Implement the application logic.
'     ' TODO: Implement.

' Terminate:
'     ' Reset error handling.
'     On Error GoTo 0

'     ' Release all allocated resources if needed.
'     ' TODO: Implement.

'     ' Re-raise the error if needed.
'     Call MRuntime.ReRaiseError

'     ' Exit the procedure.
'     Exit [Sub | Function]

' HandleError:
'     ' Store the error for further handling.
'     Call MRuntime.StoreError

'     ' TODO: Verify whether the error should be re-raised.

'     ' Resume to procedure termination.
'     Resume Terminate:
' End [Sub | Function]
