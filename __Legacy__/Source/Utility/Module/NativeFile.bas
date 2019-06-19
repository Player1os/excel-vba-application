Option Explicit
Option Private Module

' Requires MRuntime

Public Function ReadData( _
    ByRef vFilePath As String _
) As String
    ' Declare local variables.
    Dim vFileHandle As Integer

    ' Setup error handling.
    On Error GoTo HandleError:

    ' Determine the next file number available for use.
    vFileHandle = VBA.FileSystem.FreeFile

    ' Open the text file.
    Open vFilePath For Input As #vFileHandle

    ' Load file text from the file.
    ReadData = Input(VBA.LOF(vFileHandle), vFileHandle)

Terminate:
    ' Reset error handling.
    On Error GoTo 0

    ' Close the file handle.
    Close #vFileHandle

    ' Re-raise the error if needed.
    Call MRuntime.ReRaiseError

    ' Exit the procedure.
    Exit Function

HandleError:
    ' Store the error for further handling.
    Call MRuntime.StoreError

    ' Resume to procedure termination.
    Resume Terminate:
End Function

Public Sub WriteData( _
    ByRef vFilePath As String, _
    ByRef vData As String _
)
    ' Declare local variables.
    Dim vFileHandle As Integer

    ' Setup error handling.
    On Error GoTo HandleError:

    ' Determine the next file number available for use.
    vFileHandle = VBA.FileSystem.FreeFile

    ' Open the text file.
    Open vFilePath For Output As #vFileHandle

    ' Write file text to the file.
    Print #vFileHandle, vData;

Terminate:
    ' Reset error handling.
    On Error GoTo 0

    ' Close the file handle.
    Close #vFileHandle

    ' Re-raise the error if needed.
    Call MRuntime.ReRaiseError

    ' Exit the procedure.
    Exit Sub

HandleError:
    ' Store the error for further handling.
    Call MRuntime.StoreError

    ' Resume to procedure termination.
    Resume Terminate:
End Sub
