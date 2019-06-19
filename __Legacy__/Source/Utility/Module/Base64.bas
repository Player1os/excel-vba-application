' A Base64 Encoder/Decoder.
'
' This module is used to encode and decode data in Base64 format as described in RFC 1521.
'
' Home page: www.source-code.biz.
' Copyright 2007: Christian d'Heureuse, Inventec Informatik AG, Switzerland.
'
' This module is multi-licensed and may be used under the terms
' of any of the following licenses:
'
'  EPL, Eclipse Public License, V1.0 or later, http://www.eclipse.org/legal
'  LGPL, GNU Lesser General Public License, V2.1 or later, http://www.gnu.org/licenses/lgpl.html
'  GPL, GNU General Public License, V2 or later, http://www.gnu.org/licenses/gpl.html
'  AGPL, GNU Affero General Public License V3 or later, http://www.gnu.org/licenses/agpl.html
'  AL, Apache License, V2.0 or later, http://www.apache.org/licenses
'  BSD, BSD License, http://www.opensource.org/licenses/bsd-license.php
'  MIT, MIT License, http://www.opensource.org/licenses/MIT
'
' Please contact the author if you need another license.
' This module is provided "as is", without warranties of any kind.

Option Explicit
Option Private Module

Private InitDone As Boolean
Private Map1(0 To 63) As Byte
Private Map2(0 To 127) As Byte

' Encodes a string into Base64 format.
' No blanks or line breaks are inserted.
' Parameters:
'        S  a String to be encoded.
' Returns:  a String with the Base64 encoded data.
Public Function EncodeString( _
    ByVal s As String _
) As String
    EncodeString = Encode(ConvertStringToBytes(s))
End Function

' Encodes a byte array into Base64 format.
' No blanks or line breaks are inserted.
' Parameters:
'   InData  an array containing the data bytes to be encoded.
'    InLen  Optional: number of bytes to process in InData (if not provided the whole array is processed).
' Returns:  a string with the Base64 encoded data.
Public Function Encode( _
    ByRef InData() As Byte, _
    Optional ByVal InLen As Long = -1 _
) As String
    If Not InitDone Then
        Call Init
    End If
    If InLen = -1 Then
        InLen = UBound(InData) - LBound(InData) + 1
    End If
    If InLen = 0 Then
        Encode = VBA.vbNullString
        Exit Function
    End If

    ' Declare local variables.
    Dim ODataLen As Long
    Dim OLen As Long
    Dim Out() As Byte
    Dim ip0 As Long
    Dim ip As Long
    Dim op As Long
    Dim i0 As Byte
    Dim i1 As Byte
    Dim i2 As Byte
    Dim o0 As Byte
    Dim o1 As Byte
    Dim o2 As Byte
    Dim o3 As Byte

    ' Output length without padding
    ODataLen = (InLen * 4 + 2) \ 3
    ' Output length including padding
    OLen = ((InLen + 2) \ 3) * 4
    ReDim Out(0 To OLen - 1) As Byte
    ip0 = LBound(InData)

    Do While ip < InLen
        i0 = InData(ip0 + ip)
        ip = ip + 1

        If ip < InLen Then
            i1 = InData(ip0 + ip)
            ip = ip + 1
        Else
            i1 = 0
        End If

        If ip < InLen Then
            i2 = InData(ip0 + ip)
            ip = ip + 1
        Else
            i2 = 0
        End If

        o0 = i0 \ 4
        o1 = ((i0 And 3) * &H10) Or (i1 \ &H10)
        o2 = ((i1 And &HF) * 4) Or (i2 \ &H40)
        o3 = i2 And &H3F

        Out(op) = Map1(o0)
        op = op + 1
        Out(op) = Map1(o1)
        op = op + 1
        Out(op) = VBA.IIf(op < ODataLen, Map1(o2), Asc("="))
        op = op + 1
        Out(op) = VBA.IIf(op < ODataLen, Map1(o3), Asc("="))
        op = op + 1
    Loop

    Encode = ConvertBytesToString(Out)
End Function

' Decodes a string from Base64 format.
' Parameters:
'        s  a Base64 String to be decoded.
' Returns:  a String containing the decoded data.
Public Function DecodeString( _
    ByVal s As String _
) As String
    If s = VBA.vbNullString Then
        DecodeString = VBA.vbNullString
        Exit Function
    End If

    DecodeString = ConvertBytesToString(Decode(s))
End Function

' Decodes a byte array from Base64 format.
' Parameters
'        s  a Base64 String to be decoded.
' Returns:  an array containing the decoded data bytes.
Public Function Decode( _
    ByVal s As String _
) As Byte()
    If Not InitDone Then
        Call Init
    End If

    ' Declare local variables.
    Dim IBuf() As Byte
    Dim ILen As Long
    Dim OLen As Long
    Dim Out() As Byte
    Dim ip As Long
    Dim op As Long
    Dim i0 As Byte
    Dim i1 As Byte
    Dim i2 As Byte
    Dim i3 As Byte
    Dim b0 As Byte
    Dim b1 As Byte
    Dim b2 As Byte
    Dim b3 As Byte
    Dim o0 As Byte
    Dim o1 As Byte
    Dim o2 As Byte

    IBuf = ConvertStringToBytes(s)
    ILen = UBound(IBuf) + 1

    If (ILen Mod 4) <> 0 Then
        Call VBA.Err.Raise(VBA.vbObjectError, , "Length of Base64 encoded input string is not a multiple of 4.")
    End If

    Do While ILen > 0
        If IBuf(ILen - 1) <> Asc("=") Then
            Exit Do
        End If

        ILen = ILen - 1
    Loop

    OLen = (ILen * 3) \ 4
    ReDim Out(0 To OLen - 1) As Byte

    Do While ip < ILen
        i0 = IBuf(ip)
        ip = ip + 1
        i1 = IBuf(ip)
        ip = ip + 1

        If ip < ILen Then
            i2 = IBuf(ip)
            ip = ip + 1
        Else
            i2 = Asc("A")
        End If

        If ip < ILen Then
            i3 = IBuf(ip)
            ip = ip + 1
        Else
            i3 = Asc("A")
        End If

        If _
            (i0 > 127) _
            Or (i1 > 127) _
            Or (i2 > 127) _
            Or (i3 > 127) _
        Then
            Call VBA.Err.Raise(VBA.vbObjectError, , "Illegal character in Base64 encoded data.")
        End If

        b0 = Map2(i0)
        b1 = Map2(i1)
        b2 = Map2(i2)
        b3 = Map2(i3)

        If _
            (b0 > 63) _
            Or (b1 > 63) _
            Or (b2 > 63) _
            Or (b3 > 63) _
        Then
            Call VBA.Err.Raise(VBA.vbObjectError, , "Illegal character in Base64 encoded data.")
        End If

        o0 = (b0 * 4) Or (b1 \ &H10)
        o1 = ((b1 And &HF) * &H10) Or (b2 \ 4)
        o2 = ((b2 And 3) * &H40) Or b3

        Out(op) = o0
        op = op + 1
        If op < OLen Then
            Out(op) = o1
            op = op + 1
        End If
        If op < OLen Then
            Out(op) = o2
            op = op + 1
        End If
    Loop

    Decode = Out
End Function

Private Sub Init()
    ' Declare local variables.
    Dim c As Integer
    Dim i As Integer

    ' Set Map1
    i = 0
    For c = Asc("A") To VBA.Asc("Z")
        Map1(i) = c
        i = i + 1
    Next
    For c = Asc("a") To VBA.Asc("z")
        Map1(i) = c
        i = i + 1
    Next
    For c = Asc("0") To VBA.Asc("9")
        Map1(i) = c
        i = i + 1
    Next
    Map1(i) = VBA.Asc("+")
    i = i + 1
    Map1(i) = VBA.Asc("/")
    i = i + 1

    ' Set Map2
    For i = 0 To 127
        Map2(i) = 255
    Next
    For i = 0 To 63
        Map2(Map1(i)) = i
    Next

    InitDone = True
End Sub

Private Function ConvertStringToBytes( _
    ByVal s As String _
) As Byte()
    ' Declare local variables.
    Dim b1() As Byte
    Dim l As Long
    Dim b2() As Byte
    Dim p As Long
    Dim c As Long

    b1 = s
    l = (UBound(b1) + 1) \ 2
    If l = 0 Then
        ConvertStringToBytes = b1
        Exit Function
    End If

    ReDim b2(0 To l - 1) As Byte
    For p = 0 To l - 1
        c = b1(2 * p) + 256 * CLng(b1(2 * p + 1))

        If c >= 256 Then
            c = VBA.Asc("?")
        End If

        b2(p) = c
    Next
    ConvertStringToBytes = b2
End Function

Private Function ConvertBytesToString( _
    ByRef b() As Byte _
) As String
    ' Declare local variables.
    Dim l As Long
    Dim b2() As Byte
    Dim p0 As Long
    Dim p As Long

    l = UBound(b) - LBound(b) + 1
    ReDim b2(0 To (2 * l) - 1) As Byte
    p0 = LBound(b)

    For p = 0 To l - 1
        b2(2 * p) = b(p0 + p)
    Next

    ConvertBytesToString = b2
End Function
