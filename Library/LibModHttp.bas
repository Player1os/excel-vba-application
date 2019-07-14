Attribute VB_Name = "LibModHttp"
Option Explicit
Option Private Module

' Requires MUtilityCollection

Private Const vNameValueDelimiter As String = ":"

Private Function ParseResponseHeaders( _
    ByRef vRawHeaders As String _
) As Dictionary
    ' Declare local variables.
    Dim vNameValueDelimiterLength As Long
    Dim vHeaderKey As Variant
    Dim vHeader As String
    Dim vNameValueDelimiterPosition As Long
    Dim vHeaderName As String
    Dim vHeaderValue As String

    ' Determine the name value delimiter length.
    vNameValueDelimiterLength = VBA.Len(vNameValueDelimiter)

    ' Initialize the result.
    Set ParseResponseHeaders = New Dictionary

    ' Verify that the response contains any headers.
    If vRawHeaders <> VBA.vbNullString Then
        ' Split the raw headers into a collection of headers.
        For Each vHeaderKey In MUtilityCollection.Split(vRawHeaders, VBA.vbCrLf)
            ' Determine the name value delimiter position.
            vHeader = CStr(vHeaderKey)
            vNameValueDelimiterPosition = VBA.InStr(1, vHeader, vNameValueDelimiter)

            ' Verify that the delimiter was found.
            If vNameValueDelimiterPosition <> 0 Then
                ' Parse the header name and value.
                vHeaderName = VBA.Left(vHeader, vNameValueDelimiterPosition - 1)
                vHeaderValue = VBA.Mid(vHeader, vNameValueDelimiterPosition + vNameValueDelimiterLength)

                ' Ensure the header name has its first character in uppercase and the rest in lowercase.
                vHeaderName = VBA.UCase(VBA.Left(vHeaderName, 1)) & VBA.LCase(VBA.Mid(vHeaderName, 2))

                ' Check whether the header name was already encountered.
                If ParseResponseHeaders.Exists(vHeaderName) Then
                    ' Merge with the previous header value.
                    vHeaderValue = ParseResponseHeaders.Item(vHeaderName) & vbCrLf & vHeaderValue

                    ' Remove the previous header from the result.
                    Call ParseResponseHeaders.Remove(vHeaderName)
                End If

                ' Add a new header to the result.
                Call ParseResponseHeaders.Add(vHeaderName, vHeaderValue)
            End If
        Next
    End If

    ' Convert merged header values into collections.
    For Each vHeaderKey In ParseResponseHeaders
        vHeaderName = CStr(vHeaderKey)
        vHeaderValue = ParseResponseHeaders.Item(vHeaderName)

        ' Verify that the current item is a merged header.
        If VBA.InStr(1, vHeaderValue, vbCrLf) <> 0 Then
            ' Remove the header from the result.
            Call ParseResponseHeaders.Remove(vHeaderName)

            ' Re add a parsed version of the same header back into the result.
            Call ParseResponseHeaders.Add(vHeaderName, MUtilityCollection.Split(vHeaderValue, vbCrLf))
        End If
    Next
End Function

' TODO: Refactor to use MSXML reference.
Public Function Send( _
    ByRef vUrl As String, _
    Optional ByRef vMethod As String = "GET", _
    Optional ByRef vHeaders As Dictionary = Nothing, _
    Optional ByRef vBody As String = VBA.vbNullString _
) As Dictionary
    ' Declare local variables.
    Dim vXmlHttp As Object
    Dim vHeaderName As Variant

    ' Initialize the result.
    Set Send = New Dictionary

    ' Initialize the HTTP client.
    Set vXmlHttp = CreateObject("MSXML2.serverXMLHTTP")

    ' Open a connection to the supplied url.
    Call vXmlHttp.Open(vMethod, vUrl, False)

    ' Fill request headers if any were supplied.
    If Not (vHeaders Is Nothing) Then
        For Each vHeaderName In vHeaders
            Call vXmlHttp.SetRequestHeader(CStr(vHeaderName), vHeaders(CStr(vHeaderName)))
        Next
    End If

    ' Send the request with the supplied payload.
    If vBody = VBA.vbNullString Then
        Call vXmlHttp.Send
    Else
        Call vXmlHttp.Send(vBody)
    End If

    ' Fill the result status.
    Call Send.Add("Status", vXmlHttp.Status)
    Call Send.Add("StatusText", vXmlHttp.StatusText)

    ' Fill the result headers.
    Call Send.Add("Headers", ParseResponseHeaders(vXmlHttp.GetAllResponseHeaders()))

    ' Fill the result payload.
    Call Send.Add("Body", vXmlHttp.ResponseText)
End Function
