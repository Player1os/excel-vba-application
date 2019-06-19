Option Explicit
Option Private Module

' Requires MRuntime

Private Const vPermissionDeniedErrorNumber As Long = 70
Private Const vIllegalFileNameCharacters As String = "\/:*?""<>|"

Public Function DesktopPath() As String
    DesktopPath = CStr(CreateObject("WScript.Shell").SpecialFolders("Desktop"))
End Function

Public Function IsValidFileName( _
    ByRef vFileName As String _
) As Boolean
    ' Declare local variables.
    Dim vIllegalCharacterPosition As Long
    Dim vIllegalCharacterCount As Long

    ' Set the initial result.
    IsValidFileName = False

    If vFileName <> VBA.vbNullString Then
        ' A non empty file name is assumed to be valid.
        IsValidFileName = True

        ' Search for any invalid characters.
        vIllegalCharacterCount = VBA.Len(vIllegalFileNameCharacters)
        For vIllegalCharacterPosition = 1 To vIllegalCharacterCount
            If IsValidFileName Then
                If VBA.InStr(1, vFileName, VBA.Mid(vIllegalFileNameCharacters, vIllegalCharacterPosition, 1)) <> 0 Then
                    IsValidFileName = False
                End If
            End If
        Next
    End If
End Function

Public Function ReplaceInvalidFileNameCharacters( _
    ByRef vFileName As String, _
    Optional ByVal vReplacementCharacter As String = "_" _
) As String
    ' Declare local variables.
    Dim vIllegalCharacterPosition As Long
    Dim vIllegalCharacterCount As Long

    ' Set the initial result.
    ReplaceInvalidFileNameCharacters = vFileName

    ' Check whether the submitted filename is empty.
    If ReplaceInvalidFileNameCharacters = VBA.vbNullString Then
        Exit Function
    End If

    ' Search for any invalid characters.
    vIllegalCharacterCount = VBA.Len(vIllegalFileNameCharacters)
    For vIllegalCharacterPosition = 1 To vIllegalCharacterCount
        ReplaceInvalidFileNameCharacters = VBA.Replace(ReplaceInvalidFileNameCharacters, _
            VBA.Mid(vIllegalFileNameCharacters, vIllegalCharacterPosition, 1), vReplacementCharacter)
    Next
End Function

Public Function DirectoryContents( _
    ByRef vDirectoryPath As String _
) As Collection
    ' Declare local variables.
    Dim vFileName As String

    ' Initialize the result collection.
    Set DirectoryContents = New Collection

    ' Initial query to the dir function.
    vFileName = VBA.Dir(vDirectoryPath & "\*.*")

    ' Sequentially add the contents to the result.
    Do While vFileName <> VBA.vbNullString
        Call DirectoryContents.Add(vFileName)
        vFileName = VBA.Dir()
    Loop
End Function

Public Function IsOpen( _
    vFileName As String _
) As Boolean
    ' Declare local variables.
    Dim vFileHandle As Integer

    ' Setup error handling.
    On Error GoTo HandleError:

    ' Attempt to acquite a file handle and open a file.
    IsOpen = False
    vFileHandle = VBA.FileSystem.FreeFile
    Open vFileName For Input Lock Read As #vFileHandle
    IsOpen = True

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
    If VBA.Err.Number <> vPermissionDeniedErrorNumber Then
        Call MRuntime.StoreError
    End If

    ' Resume to procedure termination.
    Resume Terminate:
End Function

Public Function Exists( _
    ByRef vFilePath As String _
) As Boolean
    Exists = VBA.Dir(vFilePath) <> VBA.vbNullString
End Function

Public Function ReadData( _
    ByRef vFilePath As String, _
    Optional ByVal vCharset As String = "UTF-8" _
) As String
    ' Declare local variables.
    Dim vAdodbStream As New ADODB.Stream

    ' Setup error handling.
    On Error GoTo HandleError:

    ' Load all data from the stream.
    With vAdodbStream
        ' Set the charset and open the stream.
        .Charset = vCharset
        Call .Open

        ' Load the specified file and read its data from the stream.
        Call .LoadFromFile(vFilePath)
        ReadData = .ReadText()
    End With

Terminate:
    ' Reset error handling.
    On Error GoTo 0

    ' Close the stream.
    With vAdodbStream
        If .State <> adStateClosed Then
            Call .Close
        End If
    End With

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
    ByRef vData As String, _
    Optional ByVal vCharset As String = "UTF-8" _
)
    ' Declare local variables.
    Dim vAdodbStream As New ADODB.Stream

    ' Setup error handling.
    On Error GoTo HandleError:

    ' Store all data to the stream.
    With vAdodbStream
        ' Set the charset and open the stream.
        .Charset = vCharset
        Call .Open

        ' Write the data to the stream and save it to the specified file.
        Call .WriteText(vData)
        Call .SaveToFile(vFilePath, adSaveCreateOverWrite)
    End With

Terminate:
    ' Reset error handling.
    On Error GoTo 0

    ' Close the stream.
    With vAdodbStream
        If .State <> adStateClosed Then
            Call .Close
        End If
    End With

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

Public Sub AppendData( _
    ByRef vFilePath As String, _
    ByRef vData As String, _
    Optional ByVal vCharset As String = "UTF-8" _
)
    ' Declare local variables.
    Dim vOriginalData As String

    ' Extract contents from the existing file.
    vOriginalData = VBA.vbNullString
    If Exists(vFilePath) Then
        vOriginalData = ReadData(vFilePath, vCharset)
    End If

    ' Concatenate and write the new data to the same file.
    Call WriteData(vFilePath, vOriginalData & vData, vCharset)
End Sub
