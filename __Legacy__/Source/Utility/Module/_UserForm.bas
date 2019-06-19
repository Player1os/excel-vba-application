' Variable font size
' http://stackoverflow.com/questions/256823/finding-if-a-textbox-label-caption-fits-in-the-control
' https://support.microsoft.com/en-us/kb/76388

Function TextWidth( _
    vText As String, _
    Optional vFont As NewFont _
) As Single
    Dim theFont As New NewFont
    Dim notSeenTBox As Control

    On Error Resume Next ' Trap for vFont = Nothing
    theFont = vFont ' Try assign

    If VBA.Err.Number Then ' Can't use vFont because it's not instantiated / set.
        theFont.Name = "Tahoma"
        theFont.Size = 8
        theFont.Bold = False
        theFont.Italic = False
    End If
    On Error GoTo ErrHandler

    ' Make a TextBox, fiddle with autosize et al, retrive control width
    Set notSeenTBox = UserForms(0).Controls.Add("Forms.TextBox.1", "notSeen1", False)
    notSeenTBox.MultiLine = False
    notSeenTBox.AutoSize = True ' The trick
    notSeenTBox.Font.Name = theFont.Name
    notSeenTBox.SpecialEffect = 0
    notSeenTBox.Width = 0 ' Otherwise we get an offset (a ""feature"" from MS)
    notSeenTBox.Text = vText
    TextWidth = notSeenTBox.Width

    ' Done with it, to scrap I say
    Call UserForms(0).Controls.Remove("notSeen1")
    Exit Function

ErrHandler:
    TextWidth = -1
    Call VBA.MsgBox("TextWidth failed: " + VBA.Err.Description)
End Function
