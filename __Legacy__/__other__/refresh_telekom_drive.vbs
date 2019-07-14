Option Explicit

'''''''''''''''

Const vNameValueDelimiter = ":"

Function pHttpParseResponseHeaders( _
    ByRef vRawHeaders _
)
    ' Declare local variables.
    Dim vNameValueDelimiterLength
    Dim vHeaderKey
    Dim vHeader
    Dim vNameValueDelimiterPosition
    Dim vHeaderName
    Dim vHeaderValue

    ' Determine the name value delimiter length.
    vNameValueDelimiterLength = Len(vNameValueDelimiter)

    ' Initialize the result.
    Set pHttpParseResponseHeaders = CreateObject("Scripting.Dictionary")

    ' Verify that the response contains any headers.
    If vRawHeaders <> vbNullString Then
        ' Split the raw headers into a collection of headers.
        For Each vHeaderKey In MUtilityCollection.Split(vRawHeaders, vbCrLf)
            ' Determine the name value delimiter position.
            vHeader = CStr(vHeaderKey)
            vNameValueDelimiterPosition = InStr(1, vHeader, vNameValueDelimiter)

            ' Verify that the delimiter was found.
            If vNameValueDelimiterPosition <> 0 Then
                ' Parse the header name and value.
                vHeaderName = Left(vHeader, vNameValueDelimiterPosition - 1)
                vHeaderValue = Mid(vHeader, vNameValueDelimiterPosition + vNameValueDelimiterLength)

                ' Ensure the header name has its first character in uppercase and the rest in lowercase.
                vHeaderName = UCase(Left(vHeaderName, 1)) & LCase(Mid(vHeaderName, 2))

                ' Check whether the header name was already encountered.
                If pHttpParseResponseHeaders.Exists(vHeaderName) Then
                    ' Merge with the previous header value.
                    vHeaderValue = pHttpParseResponseHeaders.Item(vHeaderName) & vbCrLf & vHeaderValue

                    ' Remove the previous header from the result.
                    Call pHttpParseResponseHeaders.Remove(vHeaderName)
                End If

                ' Add a new header to the result.
                Call pHttpParseResponseHeaders.Add(vHeaderName, vHeaderValue)
            End If
        Next
    End If

    ' Convert merged header values into collections.
    For Each vHeaderKey In pHttpParseResponseHeaders
        vHeaderName = CStr(vHeaderKey)
        vHeaderValue = pHttpParseResponseHeaders.Item(vHeaderName)

        ' Verify that the current item is a merged header.
        If InStr(1, vHeaderValue, vbCrLf) <> 0 Then
            ' Remove the header from the result.
            Call pHttpParseResponseHeaders.Remove(vHeaderName)

            ' Re add a parsed version of the same header back into the result.
            Call pHttpParseResponseHeaders.Add(vHeaderName, MUtilityCollection.Split(vHeaderValue, vbCrLf))
        End If
    Next
End Function

Public Function pHttpSend( _
    ByRef vUrl, _
    ByRef vMethod, _
    ByRef vHeaders, _
    ByRef vBody _
)
    ' Declare local variables.
    Dim vXmlHttp
    Dim vHeaderName

    ' Initialize the result.
    Set pHttpSend = CreateObject("Scripting.Dictionary")

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

    ' pHttpSend the request with the supplied payload.
    If vBody = vbNullString Then
        Call vXmlHttp.pHttpSend
    Else
        Call vXmlHttp.pHttpSend(vBody)
    End If

    ' Fill the result status.
    Call pHttpSend.Add("Status", vXmlHttp.Status)
    Call pHttpSend.Add("StatusText", vXmlHttp.StatusText)

    ' Fill the result headers.
    Call pHttpSend.Add("Headers", pHttpParseResponseHeaders(vXmlHttp.GetAllResponseHeaders()))

    ' Fill the result payload.
    Call pHttpSend.Add("Body", vXmlHttp.ResponseText)
End Function

'''''''''''''''

Call pHttpSend("http://google.com/", "GET", Nothing, vbNullString)


' curl -k "https://st.telekomdrive.sk/" \
	' -H "Connection: keep-alive" \
	' -H 'Upgrade-Insecure-Requests: 1' \
	' -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36" ^
	' -H "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3" ^
	' -H "Accept-Encoding: gzip, deflate, br" ^
	' -H "Accept-Language: en-US,en;q=0.9,sk;q=0.8,cs;q=0.7" ^
	' --compressed

' :::: Set-Cookie: occd85af575f=8usuldj0j156osre789s11oen3; path=/; HttpOnly

' curl "https://st.telekomdrive.sk/" ^
	' -H "Connection: keep-alive" ^
	' -H "Cache-Control: max-age=0" ^
	' -H "Origin: https://st.telekomdrive.sk" ^
	' -H "Upgrade-Insecure-Requests: 1" ^
	' -H "Content-Type: application/x-www-form-urlencoded" ^
	' -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36" ^
	' -H "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3" ^
	' -H "Referer: https://st.telekomdrive.sk/" ^
	' -H "Accept-Encoding: gzip, deflate, br" ^
	' -H "Accept-Language: en-US,en;q=0.9,sk;q=0.8,cs;q=0.7" ^
	' -H "Cookie: occd85af575f=8usuldj0j156osre789s11oen3" ^
	' --data "user=osama.hassanein^&password=Tele456tele^%^24MagOraKys^&remember_login=1^&timezone-offset=2^&timezone=Europe^%^2FBerlin^&requesttoken=6S4gQPzHlVd8OF^%^2F2kM9zzYp2K^%^2FQy3b" ^
	' --compressed

' :::: Set-Cookie: occd85af575f=5tj7hiqljl1i9br7u38ct16q43; path=/; HttpOnly
' :::: Set-Cookie: oc_username=40649; expires=Thu, 13-Jun-2019 15:15:48 GMT; Max-Age=3600
' :::: Set-Cookie: oc_token=OaBS%2FRGXnzXD7a1FTLUgbv1aVdkDIO7c; expires=Thu, 13-Jun-2019 15:15:48 GMT; Max-Age=3600; httponly
' :::: Set-Cookie: oc_remember_login=1; expires=Thu, 13-Jun-2019 15:15:48 GMT; Max-Age=3600

' curl 'https://st.telekomdrive.sk/index.php/apps/files/ajax/list.php?dir=%2F&sort=name&sortdirection=asc' \
	' -H 'Accept: */*' \
	' -H 'requesttoken: 6S4gQPzHlVd8OF/2kM9zzYp2K/Qy3b' \
	' -H 'Referer: https://st.telekomdrive.sk/index.php/apps/files/?dir=%2F' \
	' -H 'X-Requested-With: XMLHttpRequest' \
	' -H 'User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36' \
	' -H 'OCS-APIREQUEST: true' \
	' --compressed

' :::: {"data":{"directory":"\/","files":[{"id":119044,"parentId":117317,"date":"June 13, 2019 at 1:48:31 PM GMT+2","mtime":1560426511000,"icon":"\/core\/img\/filetypes\/folder.svg","name":"MARS OIC KPI shared","permissions":31,"mimetype":"httpd\/unix-directory","size":0,"type":"dir","etag":"5d02380f27795"},{"id":119051,"parentId":117317,"date":"June 13, 2019 at 4:19:27 PM GMT+2","mtime":1560435567000,"icon":"\/core\/img\/filetypes\/folder.svg","name":"Shared","permissions":31,"mimetype":"httpd\/unix-directory","size":0,"type":"dir","etag":"5d025b6f3e372"}],"permissions":31},"status":"success"}

' curl "https://st.telekomdrive.sk/index.php/core/ajax/share.php" ^
	' -H "requesttoken: 6S4gQPzHlVd8OF/2kM9zzYp2K/Qy3b" ^
	' -H "Origin: https://st.telekomdrive.sk" ^
	' -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36" ^
	' -H "OCS-APIREQUEST: true" ^
	' -H "Content-Type: application/x-www-form-urlencoded; charset=UTF-8" ^
	' -H "Accept: */*" ^
	' -H "Referer: https://st.telekomdrive.sk/index.php/apps/files/?dir=^%^2F" ^
	' -H "X-Requested-With: XMLHttpRequest" ^
	' --data "action=share^&itemType=folder^&itemSource=119044^&shareType=3^&shareWith=^&permissions=7^&itemSourceName=MARS+OIC+KPI+shared^&expirationDate=2019-6-27+00^%^3A00^%^3A00" --compressed

' :::: {"data":{"token":"NJo7SbyWTQTGq8P"},"status":"success"}


