Option Explicit
Option Private Module

Public Function DateTimeString( _
    ByVal vDate As Date, _
    Optional ByVal vDateDelimiter As String = "/", _
    Optional ByVal vTimeDelimiter As String = ":", _
    Optional ByVal vPartDelimiter As String = " " _
) As String
    DateTimeString = VBA.Format(vDate, _
        "yyyy" & vDateDelimiter & "mm" & vDateDelimiter & "dd" _
        & vPartDelimiter _
        & "Hh" & vTimeDelimiter & "Nn" & vTimeDelimiter & "Ss")
End Function
