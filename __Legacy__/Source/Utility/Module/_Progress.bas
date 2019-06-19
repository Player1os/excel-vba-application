Option Explicit

' TODO: Finalize and combine with auto resizing.

Private Const vInvalidProgressSecond As Byte = 60

Private vLastProgressSecond As Byte

Public Sub Initialize()
    ' Reset the last progress second.
    vLastProgressSecond = vInvalidProgressSecond

    Call Show
End Sub

Public Sub Update( _
    ByRef vDescription As String _
)
    ' ByVal vCurrentValue As Long, _
    ' ByVal vTotalValue As Long _

    ' Declare local variables.
    Dim vCurrentSecond As Byte

    ' Verify that a second has elapsed since the last progress update.
    vCurrentSecond = VBA.Second(VBA.Now())
    If vCurrentSecond <> vLastProgressSecond Then
        vLastProgressSecond = vCurrentSecond

        ' Update the progress indicators if a handler is defined.
        'If pProcedureExists(vIsUpdateProgressCallbackName) Then
        '    Call Application.Run(vIsUpdateProgressCallbackName, vDescription, vCurrentValue, vTotalValue)
        'End If
        Me.Label.Caption = vDescription

        ' Prompt event execution to refresh any visualizations that may display the progress.
        Call Me.Repaint()
        Call VBA.DoEvents()
    End If
End Sub

Public Sub Terminate()
    Call Hide()
End Sub
