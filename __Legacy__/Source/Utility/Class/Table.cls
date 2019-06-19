Option Explicit

' Requires CUtilitySet
' Requires MUtilityRange

Private vHeaderRange As Range

Public Function UseHeaderRange( _
    ByRef vNewHeaderRange As Range _
) As CUtilityTable
    ' Declare local variables.
    Dim vCellRange As Range
    Dim vHeaderNamesSet As CUtilitySet

    ' Verify that the header range consists of a single row.
    If vNewHeaderRange.Rows.Count <> 1 Then
        Call MRuntime.RaiseError("CUtilityTable.UseHeaderRange", "The header range '" & vNewHeaderRange.Address() & "' must have exactly one row.")
    End If

    ' Verify that the header range contains unique non empty values in each cell.
    Set vHeaderNamesSet = New CUtilitySet
    For Each vCellRange In vNewHeaderRange
        ' Verify that the current cell value is non-empty.
        If vCellRange.Value2 = VBA.vbNullString Then
            Call MRuntime.RaiseError("CUtilityTable.UseHeaderRange", "The header cell '" & vCellRange.Address() & "' contains an empty value.")
        End If

        ' Verify that the current cell value is unique within the range.
        If vHeaderNamesSet.Exists(CStr(vCellRange.Value2)) Then
            Call MRuntime.RaiseError("CUtilityTable.UseHeaderRange", "The header cell '" & vCellRange.Address() & "' contains a non-unique value.")
        End If

        ' Add to the set of encountered cell values.
        Call vHeaderNamesSet.Add(CStr(vCellRange.Value2))
    Next

    ' Store the new header range.
    Set vHeaderRange = vNewHeaderRange

    ' Return the current instance for chaining.
    Set UseHeaderRange = Me
End Function

Public Function GetHeaderIndex( _
    ByRef vHeaderName As String _
) As Long
    ' Declare local variables.
    Dim vCellRange As Range

    ' Set an invalid value as the initial result value.
    GetHeaderIndex = -1

    ' Search for the given name in the header range.
    For Each vCellRange In vHeaderRange
        If GetHeaderIndex = -1 Then
            If vCellRange.Value2 = vHeaderName Then
                GetHeaderIndex = vCellRange.Column - vHeaderRange.Column + 1
            End If
        End If
    Next

    ' Verify that the name was found.
    If GetHeaderIndex = -1 Then
        Call MRuntime.RaiseError("CUtilityTable.GetHeaderIndex", "The name '" & vHeaderName & "' does not exist in the header range.")
    End If
End Function

Public Function HasHeader( _
    ByRef vHeaderName As String _
) As Boolean
    ' Declare local variables.
    Dim vCellRange As Range

    ' Set the initial result value.
    HasHeader = False

    ' Search for the given name in the header range.
    For Each vCellRange In vHeaderRange
        If Not HasHeader Then
            If vCellRange.Value2 = vHeaderName Then
                HasHeader = True
            End If
        End If
    Next
End Function

Private Function VerifyHeader( _
    ByRef vHeaderName As String _
) As CUtilityTable
    ' Delegate the computation to a semantically identical method.
    Call GetHeaderIndex(vHeaderName)

    ' Return the current instance for chaining.
    Set VerifyHeader = Me
End Function

Public Function GetHeaderNames() As Collection
    ' Declare local variables.
    Dim vCellRange As Range

    ' Intialize the result.
    Set GetHeaderNames = New Collection

    ' Collect the header range cell values.
    For Each vCellRange In vHeaderRange
        Call GetHeaderNames.Add(CStr(vCellRange.Value2))
    Next
End Function

Public Function GetHeaderIndexToHeaderNameMap() As Dictionary
    ' Declare local variables.
    Dim vHeaderNames As Collection
    Dim vHeaderNamesCount As Long
    Dim vHeaderIndex As Long

    ' Retrieve the header names.
    Set vHeaderNames = GetHeaderNames()
    vHeaderNamesCount = vHeaderNames.Count

    ' Intialize the result.
    Set GetHeaderIndexToHeaderNameMap = New Dictionary

    ' Collect the header range cell values and map them to the cells' indices.
    For vHeaderIndex = 1 To vHeaderNamesCount
        Call GetHeaderIndexToHeaderNameMap.Add(vHeaderIndex, vHeaderNames(vHeaderIndex))
    Next
End Function

Public Function GetRowCount() As Long
    ' Declare local variables.
    Dim vLastRowRange As Range

    ' Set the initial result value.
    Set vLastRowRange = MUtilityRange.GetLastRowRange(vHeaderRange.Offset(1))

    ' Verify that the table has any rows and compute the row count.
    If MUtilityRange.IsEmpty(vLastRowRange) Then
        GetRowCount = 0
    Else
        GetRowCount = vLastRowRange.Row - vHeaderRange.Row
    End If
End Function

Public Function HasRowIndex( _
    ByVal vRowIndex As Long _
) As Boolean
    ' Verify whether the given row index lies within the table's bounds.
    HasRowIndex = (1 <= vRowIndex) And (vRowIndex <= GetRowCount())
End Function

Private Sub VerifyRowIndex( _
    ByVal vRowIndex As Long _
)
    ' Delegate the computation to a semantically identical method.
    If Not HasRowIndex(vRowIndex) Then
        Call MRuntime.RaiseError("CUtilityTable.VerifyRowIndex", "The row index '" & CStr(vRowIndex) & "' is beyond the table's bounds.")
    End If
End Sub

Public Function GetValues() As Collection
    ' Declare local variables.
    Dim vHeaderNames As Collection
    Dim vHeaderNamesCount As Long
    Dim vRowRange As Range
    Dim vRowValues As Dictionary
    Dim vHeaderIndex As Long

    ' Retrieve the header names and set the initaial row range.
    Set vHeaderNames = GetHeaderNames()
    vHeaderNamesCount = vHeaderNames.Count
    Set vRowRange = vHeaderRange.Offset(1)

    ' Intialize the result.
    Set GetValues = New Collection

    ' Iterate through all non-empty row ranges.
    Do While Not MUtilityRange.IsEmpty(vRowRange)
        ' Create a new row values instance.
        Set vRowValues = New Dictionary

        ' Collect the new row values.
        For vHeaderIndex = 1 To vHeaderNamesCount
            Call vRowValues.Add(vHeaderNames(vHeaderIndex), CStr(vRowRange(1, vHeaderIndex).Value2))
        Next

        ' Add the row values to the result.
        Call GetValues.Add(vRowValues)

        ' Continue to the next row range.
        Set vRowRange = vRowRange.Offset(1)
    Loop
End Function

Public Function SetValues( _
    ByRef vValues As Collection _
) As CUtilityTable
    ' Declare local variables.
    Dim vHeaderIndexToHeaderNameMap As Dictionary
    Dim vHeaderNamesCount As Long
    Dim vRowRange As Range
    Dim vRowValues As Dictionary
    Dim vHeaderIndex As Long

    ' Retrieve the header index to header name map and set the initaial row range.
    Set vHeaderIndexToHeaderNameMap = GetHeaderIndexToHeaderNameMap()
    vHeaderNamesCount = vHeaderIndexToHeaderNameMap.Count
    Set vRowRange = vHeaderRange.Offset(1)

    ' Iterate through all the given values.
    For Each vRowValues In vValues
        ' Fill the new cell values.
        For vHeaderIndex = 1 To vHeaderNamesCount
            vRowRange(1, vHeaderIndex).Value2 = vRowValues(vHeaderIndexToHeaderNameMap(vHeaderIndex))
        Next

        ' Continue to the next row range.
        Set vRowRange = vRowRange.Offset(1)
    Next

    ' Iterate though all the remaining non-empty row ranges and clear them.
    Do While Not MUtilityRange.IsEmpty(vRowRange)
        Call vRowRange.Clear
        Set vRowRange = vRowRange.Offset(1)
    Loop

    ' Return the current instance for chaining.
    Set SetValues = Me
End Function

Public Function GetHeaderValues( _
    ByRef vHeaderName As String _
) As Collection
    ' Declare local variables.
    Dim vHeaderIndex As Long
    Dim vRowRange As Range

    ' Retrieve the header index and set the initaial row range.
    vHeaderIndex = GetHeaderIndex(vHeaderName)
    Set vRowRange = vHeaderRange.Offset(1)

    ' Intialize the result.
    Set GetHeaderValues = New Collection

    ' Iterate through all non-empty row ranges.
    Do While Not MUtilityRange.IsEmpty(vRowRange)
        ' Collect the current header value.
        Call GetHeaderValues.Add(CStr(vRowRange(1, vHeaderIndex).Value2))

        ' Continue to the next row range.
        Set vRowRange = vRowRange.Offset(1)
    Loop
End Function

Public Function SetHeaderValues( _
    ByRef vHeaderName As String, _
    ByRef vHeaderValues As Collection _
) As CUtilityTable
    ' Declare local variables.
    Dim vHeaderIndex As Long
    Dim vRowRange As Range
    Dim vHeaderValuesCount As Long
    Dim vRowIndex As Long

    ' Retrieve the header index and set the initaial row range.
    vHeaderIndex = GetHeaderIndex(vHeaderName)
    Set vRowRange = vHeaderRange.Offset(1)
    vHeaderValuesCount = vHeaderValues.Count

    ' Iterate through all the given header values.
    For vRowIndex = 1 To vHeaderValuesCount
        ' Fill the new cell value.
        vRowRange(1, vHeaderIndex).Value2 = vHeaderValues(vRowIndex)

        ' Continue to the next row range.
        Set vRowRange = vRowRange.Offset(1)
    Next

    ' Iterate though all the remaining non-empty row ranges and clear them.
    Do While Not MUtilityRange.IsEmpty(vRowRange)
        Call vRowRange.Clear
        Set vRowRange = vRowRange.Offset(1)
    Loop

    ' Return the current instance for chaining.
    Set SetHeaderValues = Me
End Function

Public Function GetRowValues( _
    ByVal vRowIndex As Long _
) As Dictionary
    ' Declare local variables.
    Dim vHeaderNames As Collection
    Dim vHeaderNamesCount As Long
    Dim vRowRange As Range
    Dim vHeaderIndex As Long

    ' Verify the given row index.
    Call VerifyRowIndex(vRowIndex)

    ' Retrieve the header names and set the row range to the given row index.
    Set vHeaderNames = GetHeaderNames()
    vHeaderNamesCount = vHeaderNames.Count
    Set vRowRange = vHeaderRange.Offset(vRowIndex)

    ' Intialize the result.
    Set GetRowValues = New Dictionary

    ' Collect the result row values.
    For vHeaderIndex = 1 To vHeaderNamesCount
        Call GetRowValues.Add(vHeaderNames(vHeaderIndex), CStr(vRowRange(1, vHeaderIndex).Value2))
    Next
End Function

Public Function SetRowValues( _
    ByVal vRowIndex As Long, _
    ByRef vRowValues As Dictionary _
) As CUtilityTable
    ' Declare local variables.
    Dim vHeaderIndexToHeaderNameMap As Dictionary
    Dim vHeaderNamesCount As Long
    Dim vRowRange As Range
    Dim vHeaderIndex As Long

    ' Verify the given row index.
    Call VerifyRowIndex(vRowIndex)

    ' Retrieve the header index to header name map and set the row range to the given row index.
    Set vHeaderIndexToHeaderNameMap = GetHeaderIndexToHeaderNameMap()
    vHeaderNamesCount = vHeaderIndexToHeaderNameMap.Count
    Set vRowRange = vHeaderRange.Offset(vRowIndex)

    ' Fill the cell row values.
    For vHeaderIndex = 1 To vHeaderNamesCount
        vRowRange(1, vHeaderIndex).Value2 = vRowValues(vHeaderIndexToHeaderNameMap(vHeaderIndex))
    Next

    ' Return the current instance for chaining.
    Set SetRowValues = Me
End Function

Public Function AddRowValues( _
    ByRef vRowValues As Dictionary _
) As CUtilityTable
    ' Declare local variables.
    Dim vHeaderIndexToHeaderNameMap As Dictionary
    Dim vHeaderNamesCount As Long
    Dim vRowRange As Range
    Dim vHeaderIndex As Long

    ' Retrieve the header index to header name map and set the row range to the given row index.
    Set vHeaderIndexToHeaderNameMap = GetHeaderIndexToHeaderNameMap()
    vHeaderNamesCount = vHeaderIndexToHeaderNameMap.Count
    Set vRowRange = vHeaderRange.Offset(GetRowCount() + 1)

    ' Fill the cell row values.
    For vHeaderIndex = 1 To vHeaderNamesCount
        vRowRange(1, vHeaderIndex).Value2 = vRowValues(vHeaderIndexToHeaderNameMap(vHeaderIndex))
    Next

    ' Return the current instance for chaining.
    Set AddRowValues = Me
End Function

Public Function GetValue( _
    ByRef vHeaderName As String, _
    ByVal vRowIndex As Long _
) As String
    ' Verify the given row index.
    Call VerifyRowIndex(vRowIndex)

    ' Collect the specified cell value.
    GetValue = CStr(vHeaderRange.Offset(vRowIndex)(1, GetHeaderIndex(vHeaderName)).Value2)
End Function

Public Function SetValue( _
    ByRef vHeaderName As String, _
    ByVal vRowIndex As Long, _
    ByRef vValue As String _
) As CUtilityTable
    ' Verify the given row index.
    Call VerifyRowIndex(vRowIndex)

    ' Fill the specified cell value.
    vHeaderRange.Offset(vRowIndex)(1, GetHeaderIndex(vHeaderName)).Value2 = vValue

    ' Return the current instance for chaining.
    Set SetValue = Me
End Function

Public Function GetValueToRowValuesMap( _
    ByRef vFromHeaderName As String _
) As Dictionary
    ' Declare local variables.
    Dim vValues As Collection
    Dim vRowValues As Dictionary

    ' Verify the given header name.
    Call VerifyHeader(vFromHeaderName)

    ' Retrieve the table's values.
    Set vValues = GetValues()

    ' Intialize the result.
    Set GetValueToRowValuesMap = New Dictionary

    ' Create a mapping from each header value to the corresponding row values.
    For Each vRowValues In vValues
        Call GetValueToRowValuesMap.Add(vRowValues(vFromHeaderName), vRowValues)
    Next
End Function

Public Function GetValueToValueMap( _
    ByRef vFromHeaderName As String, _
    ByRef vToHeaderName As String _
) As Dictionary
    ' Declare local variables.
    Dim vFromHeaderValues As Collection
    Dim vToHeaderValues As Collection
    Dim vRowCount As Long
    Dim vRowIndex As Long

    ' Retrieve header values for both the from and the to header.
    Set vFromHeaderValues = GetHeaderValues(vFromHeaderName)
    Set vToHeaderValues = GetHeaderValues(vToHeaderName)
    vRowCount = vFromHeaderValues.Count

    ' Intialize the result.
    Set GetValueToValueMap = New Dictionary

    ' Create a mapping from each from header value to the to header value.
    For vRowIndex = 1 To vRowCount
        Call GetValueToValueMap.Add(vFromHeaderValues(vRowIndex), vToHeaderValues(vRowIndex))
    Next
End Function
