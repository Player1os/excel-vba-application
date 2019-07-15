Attribute VB_Name = "ModController"
Option Explicit
Option Private Module

' Requires ThisUserForm
' Requires Runtime

Public Sub Navigate( _
    ByRef vPath As String, _
    ByRef vParameters As Dictionary _
)
    ' Declare local variables.
    Dim vInnerHtml As String
    Dim vParameterKey As Variant

    ' Check whether the application is running in background mode.
    If Runtime.IsBackgroundModeEnabled() Then
        ' Output the current timestamp to a file.
        With Runtime.FileSystemObject()
            With .CreateTextFile(.BuildPath(ThisWorkbook.Path, "Output.log"))
                Call .WriteLine("[[Time]]")
                Call .WriteLine(CStr(Now))
                Call .WriteLine("[[Path]]")
                Call .WriteLine(vPath)
                Call .WriteLine("[[Parameters]]")
                If vParameters.Count = 0 Then
                    Call .WriteLine("None")
                Else
                    For Each vParameterKey In vParameters.Keys()
                        Call .WriteLine(" - " & CStr(vParameterKey) & " => " & CStr(vParameters(vParameterKey)))
                    Next
                End If
                Call .Close
            End With
        End With
    Else
        ' Display the path and parameters on the loaded html page.
        vInnerHtml = "<h1>Time</h1>"
        vInnerHtml = vInnerHtml & CStr(Now)
        vInnerHtml = vInnerHtml & "<h1>Path</h1>"
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
