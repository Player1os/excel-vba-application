Private Sub pActionExportProject()
    ' Declare local variables.
    Dim vComponent As VBIDE.VBComponent
    Dim vSuffix As String
    Dim vComponentPath As String
    Dim vDirectoryPath As String
    Dim vFileNameValue As Variant

    ' Make sure the project is unprotected.
    If ThisWorkbook.VBProject.Protection = vbext_pp_locked Then
        Call VBA.MsgBox("The VBA Project must be unprotected.", vbExclamation)
        Exit Sub
    End If

    ' Remove all previous export files.
    vDirectoryPath = ThisWorkbook.Path & "\Export"
    For Each vFileNameValue In MUtilityFile.DirectoryContents(vDirectoryPath)
        If CStr(vFileNameValue) <> ".gitignore" Then
            Call VBA.Kill(vDirectoryPath & "\" & CStr(vFileNameValue))
        End If
    Next

    ' Iterate through each component in the project.
    For Each vComponent In ThisWorkbook.VBProject.VBComponents
        ' Disable the suffix by default.
        vSuffix = VBA.vbNullString

        ' Determine the file suffix to use.
        Select Case vComponent.Type
            Case VBIDE.vbext_ct_ClassModule
                vSuffix = "cls"
            Case VBIDE.vbext_ct_MSForm
                vSuffix = "frm"
            Case VBIDE.vbext_ct_StdModule
                If vComponent.name <> "MRuntime" Then
                    vSuffix = "bas"
                End If
        End Select

        ' Verify that an exportable component was encountered.
        If vSuffix <> VBA.vbNullString Then
            On Error Resume Next

            ' Attempt to export the file.
            vComponentPath = vDirectoryPath & "\" & vComponent.name & "." & vSuffix
            Call vComponent.Export(vComponentPath)

            ' Report failure if it had occured.
            If VBA.Err.Number <> 0 Then
                Call VBA.MsgBox("Failed to export '" & vComponentPath & "'.")
                Call VBA.Err.Clear
            End If

            On Error GoTo 0
        End If
    Next
End Sub

Private Sub pActionOpenConfig()
    Call Application.Workbooks.Open(ConfigFilePath())
End Sub

Private Sub pImportProject()
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    Set cmpComponents = wkbTarget.VBProject.VBComponents

    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files

        If (objFSO.GetExtensionName(objFile.name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If

    Next objFile

    MsgBox "Import is ready"
End Sub
