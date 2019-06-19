Option Explicit
Option Private Module

Private Const vAccCharacters As String = "áÁäÄčČďĎéÉěĚíÍľĽĺĹňŇóÓöÖőŐôÔŕŔřŘšŠťŤůŮúÚüÜűŰýÝžŽ"
Private Const vRegCharacters As String = "aAaAcCdDeEeEiIlLlLnNoOoOoOoOrRrRsStTuUuUuUuUyYzZ"

Public Function StripAccent( _
    ByRef vAccentedText As String _
) As String
    Dim vCharacterCount As Long
    Dim vCharacterPosition As Long

    vCharacterCount = VBA.Len(vAccCharacters)

    StripAccent = vAccentedText

    For vCharacterPosition = 1 To vCharacterCount
        StripAccent = VBA.Replace(StripAccent, _
            VBA.Mid(vAccCharacters, vCharacterPosition, 1), _
            VBA.Mid(vRegCharacters, vCharacterPosition, 1))
    Next
End Function

Public Function EvaluateTemplate( _
    ByRef vTemplate As String, _
    ByRef vValues As Dictionary _
) As String
    ' Declare local variables.
    Dim vValue As Variant

    ' Initialize the result with the body template.
    EvaluateTemplate = vTemplate

    ' Iterate through each of the submitted row value keys.
    For Each vValue In vValues
        ' Replace each template variable with the corresponding row value.
        EvaluateTemplate = VBA.Replace(EvaluateTemplate, _
            "{{" & CStr(vValue) & "}}", CStr(vValues(vValue)))
    Next
End Function

Public Function AsciiToUnicode( _
    ByRef vAsciiValue As String _
) As String
    Dim vCharacterPosition As Long
    Dim vCharacterCount As Long
    Dim vCharacter As String

    AsciiToUnicode = VBA.vbNullString

    vCharacterPosition = 1
    vCharacterCount = VBA.Len(strInput)
    Do While vCharacterPosition <= vCharacterCount
        vCharacter = VBA.Mid(vAsciiValue, vCharacterPosition, 1)

        If vCharacter = "\" Then
            AsciiToUnicode = AsciiToUnicode & VBA.ChrW(CLng("&H" & VBA.Mid(vAsciiValue, vCharacterPosition + 1, 4)))
            vCharacterPosition = vCharacterPosition + 5
        Else
            AsciiToUnicode = AsciiToUnicode & vCharacter
            vCharacterPosition = vCharacterPosition + 1
        End If
    Loop
End Function
