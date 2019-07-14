Attribute VB_Name = "ModController"
Option Explicit
Option Private Module

Public Sub Navigate( _
    ByRef vPath As String _
)
    ThisUserForm.ThisWebBrowser.Document.body.InnerHTML = "<p>" & vPath & "</p>" _
        & "<ul><li><a href=""#Alpha#1"">Alpha</a></li><li><a href=""#Beta#2"">Beta</a></li><li><a href=""#Gamma#3"">Gamma</a></li></ul>" _
        & "<input type=""text"" />"

    ' Call ThisUserForm.Hide
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
