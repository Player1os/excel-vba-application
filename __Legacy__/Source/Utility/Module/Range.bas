Option Explicit
Option Private Module

' Requires MRuntime

Public Function IsEmpty( _
    ByRef vRange As Range _
) As Boolean
    ' Declare local variables.
    Dim vCellRange As Range

    ' Set the initial result.
    IsEmpty = True

    ' Search for a non empty cell value within the range.
    For Each vCellRange In vRange
        If IsEmpty Then
            If CStr(vCellRange.Value2) <> VBA.vbNullString Then
                IsEmpty = False
                Exit Function
            End If
        End If
    Next
End Function

Public Function GetLastColumnRange( _
    ByRef vColumnRange As Range _
) As Range
    ' Verify that the given range spans a single row.
    If vColumnRange.Columns.Count <> 1 Then
        Call MRuntime.RaiseError("MUtilityRange.GetLastColumnRange", "The submitted range '" _
            & vColumnRange.Address() & "' must have exactly one column.")
    End If

    ' Set the initial result.
    Set GetLastColumnRange = vColumnRange

    ' Verify if the current row range is empty.
    If IsEmpty(GetLastColumnRange) Then
        Exit Function
    End If

    ' Verify if the next row range is empty.
    If IsEmpty(GetLastColumnRange.Offset(0, 1)) Then
        Exit Function
    End If

    ' Skip to the bottom.
    Set GetLastColumnRange = GetLastColumnRange.Offset(0, _
        GetLastColumnRange.End(xlToRight).Column - GetLastColumnRange.Column)

    ' Iterate downward until an empty row is encountered.
    Do While Not IsEmpty(GetLastColumnRange.Offset(0, 1))
        Set GetLastColumnRange = GetLastColumnRange.Offset(0, 1)
    Loop
End Function

Public Function GetLastRowRange( _
    ByRef vRowRange As Range _
) As Range
    ' Verify that the given range spans a single row.
    If vRowRange.Rows.Count <> 1 Then
        Call MRuntime.RaiseError("MUtilityRange.GetLastRowRange", "The submitted range '" _
            & vRowRange.Address() & "' must have exactly one row.")
    End If

    ' Set the initial result.
    Set GetLastRowRange = vRowRange

    ' Verify if the current row range is empty.
    If IsEmpty(GetLastRowRange) Then
        Exit Function
    End If

    ' Verify if the next row range is empty.
    If IsEmpty(GetLastRowRange.Offset(1)) Then
        Exit Function
    End If

    ' Skip to the bottom.
    Set GetLastRowRange = GetLastRowRange.Offset(GetLastRowRange.End(xlDown).Row - GetLastRowRange.Row)

    ' Iterate downward until an empty row is encountered.
    Do While Not IsEmpty(GetLastRowRange.Offset(1))
        Set GetLastRowRange = GetLastRowRange.Offset(1)
    Loop
End Function
