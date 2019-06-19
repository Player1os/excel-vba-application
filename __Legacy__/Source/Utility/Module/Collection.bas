Option Explicit
Option Private Module

Public Function Join( _
    ByRef vCollection As Collection, _
    ByRef vDelimiter As String _
) As String
    ' Declare local variables.
    Dim vIndex As Long

    ' Check if the collection is empty.
    If vCollection.Count() = 0 Then
        Exit Function
    End If

    ' Set the initial value for the result.
    Join = vCollection(1)

    ' Concatenate the individual values and the delimiter.
    For vIndex = 2 To vCollection.Count()
        Join = Join & vDelimiter & vCollection(vIndex)
    Next
End Function

Public Function Split( _
    ByVal vValue As String, _
    ByRef vDelimiter As String _
) As Collection
    ' Declare local variables.
    Dim vDelimiterLength As Long
    Dim vDelimiterPosition As Long

    ' Store the delimiter string length.
    vDelimiterLength = VBA.Len(vDelimiter)

    ' Initialize the result.
    Set Split = New Collection

    ' Gradually chip away at the string value until the delimiter is not found.
    Do
        ' Find the delimiter.
        vDelimiterPosition = VBA.InStr(1, vValue, vDelimiter)

        If vDelimiterPosition = 0 Then
            ' Add the remaining string to the result, if not found.
            Call Split.Add(vValue)
            vValue = VBA.vbNullString

            ' Exit the loop.
            Exit Do
        Else
            ' Add up to the delimiter position to the result, if found.
            Call Split.Add(VBA.Left(vValue, vDelimiterPosition - 1))
            vValue = VBA.Mid(vValue, vDelimiterPosition + vDelimiterLength)
        End If
    Loop
End Function

Public Function ToArray( _
    ByRef vCollection As Collection _
) As Variant()
    ' Declare local variables.
    Dim vItemCount As Long
    Dim vArray() As Variant
    Dim vPosition As Long

    ' Initialize the array.
    vItemCount = vCollection.Count
    ReDim vArray(1 To vItemCount)

    ' Copy the collection values to the array.
    For vPosition = 1 To vItemCount
        vArray(vPosition) = vCollection(vPosition)
    Next

    ' Copy array to the result.
    ToArray = vArray
End Function

Public Function FromArray( _
    ByRef vArray() As Variant _
) As Collection
    ' Declare local variables.
    Dim vValue As Variant

    ' Initialize the result.
    Set FromArray = New Collection

    ' Copy the array values to the collection.
    For Each vValue In vArray
        Call FromArray.Add(vValue)
    Next
End Function

Public Sub SortArray( _
    ByRef vArray As Variant, _
    ByRef vComparisonFunctionName As String, _
    Optional ByVal vLow As Long = -1, _
    Optional ByVal vHigh As Long = -1 _
)
    ' Declare local variables.
    Dim vTempLow As Long
    Dim vTempHigh As Long
    Dim vPivot As Variant
    Dim vTempValue As Variant

    ' Verify that the array isn't already presorted.
    If vLow >= vHigh Then
        Exit Sub
    End If

    ' Override the default array boundaries.
    If vLow = -1 Then
        vLow = LBound(vArray)
    End If
    If vHigh = -1 Then
        vHigh = UBound(vArray)
    End If

    ' Set the pivot to the middle of the array.
    If VBA.IsObject(vArray(vLow)) Then
        Set vPivot = vArray((vLow + vHigh) \ 2)
    Else
        vPivot = vArray((vLow + vHigh) \ 2)
    End If

    ' Initialize the swap pointers.
    vTempLow = vLow - 1
    vTempHigh = vHigh + 1

    ' Ensure all values to the left of the pivot are smaller than the pivot
    ' and all values to the right of the pivot are larger than the pivot.
    Do
        ' Search for a value smaller than the pivot.
        Do
            vTempLow = vTempLow + 1
        Loop While (vTempLow < vHigh) _
            And Application.Run(vComparisonFunctionName, vArray(vTempLow), vPivot)

        ' Search for a value larger than the pivot.
        Do
            vTempHigh = vTempHigh - 1
        Loop While (vTempHigh > vLow) _
            And Application.Run(vComparisonFunctionName, vPivot, vArray(vTempHigh))

        ' Check whether the swap pointers have met.
        If vTempLow >= vTempHigh Then
            Exit Do
        End If

        ' Swap the unconforming array values.
        If VBA.IsObject(vPivot) Then
            Set vTempValue = vArray(vTempLow)
            Set vArray(vTempLow) = vArray(vTempHigh)
            Set vArray(vTempHigh) = vTempValue
        Else
            vTempValue = vArray(vTempLow)
            vArray(vTempLow) = vArray(vTempHigh)
            vArray(vTempHigh) = vTempValue
        End If
    Loop

    ' Call recursiveley on the unsorted left and right halves of the array.
    Call SortArray(vArray, vComparisonFunctionName, vLow, vTempHigh)
    Call SortArray(vArray, vComparisonFunctionName, vTempHigh + 1, vHigh)
End Sub

Private Function pDefaultComparisonFunction( _
    ByRef vValue1 As Variant, _
    ByRef vValue2 As Variant _
) As Boolean
    ' Compare the values based on their string representations.
    pDefaultComparisonFunction = False
    If VBA.StrComp(CStr(vValue1), CStr(vValue2)) = -1 Then
        ' Return true only if the first value is smaller than the second value.
        pDefaultComparisonFunction = True
    End If
End Function

Public Function Sort( _
    ByRef vCollection As Collection, _
    Optional ByRef vComparisonFunctionName As String = "pDefaultComparisonFunction" _
) As Collection
    ' Declare local variables.
    Dim vArray() As Variant

    ' Convert the collection to an array and sort it in place.
    vArray = ToArray(vCollection)
    Call SortArray(vArray, vComparisonFunctionName)

    ' Set the result to the sorted array after it is converted back to a collection.
    Set Sort = FromArray(vArray)
End Function
