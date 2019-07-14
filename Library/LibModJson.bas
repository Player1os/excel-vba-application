Attribute VB_Name = "LibModJson"
''
' VBA-JSON v2.2.3
' (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
'
' JSON Converter for VBA
'
' Errors:
' 10001 - JSON parse error
'
' @class JsonConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Based originally on vba-json (with extensive changes)
' BSD license included below
'
' JSONLib, http://code.google.com/p/vba-json/
'
' Copyright (c) 2013, Ryo Yokoyama
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit
Option Private Module

' === VBA-UTC Headers
#If Mac Then

#If VBA7 Then

' 64-bit Mac (2016)
Private Declare PtrSafe Function utc_popen Lib "libc.dylib" Alias "popen" ( _
    ByVal utc_Command As String, _
    ByVal utc_Mode As String _
) As LongPtr
Private Declare PtrSafe Function utc_pclose Lib "libc.dylib" Alias "pclose" ( _
    ByVal utc_File As Long _
) As LongPtr
Private Declare PtrSafe Function utc_fread Lib "libc.dylib" Alias "fread" ( _
    ByVal utc_Buffer As String, _
    ByVal utc_Size As LongPtr, _
    ByVal utc_Number As LongPtr, _
    ByVal utc_File As LongPtr _
) As LongPtr
Private Declare PtrSafe Function utc_feof Lib "libc.dylib" Alias "feof" ( _
    ByVal utc_File As LongPtr _
) As LongPtr

#Else

' 32-bit Mac
Private Declare Function utc_popen Lib "libc.dylib" Alias "popen" ( _
    ByVal utc_Command As String, _
    ByVal utc_Mode As String _
) As Long
Private Declare Function utc_pclose Lib "libc.dylib" Alias "pclose" ( _
    ByVal utc_File As Long _
) As Long
Private Declare Function utc_fread Lib "libc.dylib" Alias "fread" ( _
    ByVal utc_Buffer As String, _
    ByVal utc_Size As Long, _
    ByVal utc_Number As Long, _
    ByVal utc_File As Long _
) As Long
Private Declare Function utc_feof Lib "libc.dylib" Alias "feof" ( _
    ByVal utc_File As Long _
) As Long

#End If

#ElseIf VBA7 Then

' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
Private Declare PtrSafe Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" ( _
    utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION _
) As Long
Private Declare PtrSafe Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" ( _
    utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, _
    utc_lpUniversalTime As utc_SYSTEMTIME, _
    utc_lpLocalTime As utc_SYSTEMTIME _
) As Long
Private Declare PtrSafe Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" ( _
    utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, _
    utc_lpLocalTime As utc_SYSTEMTIME, _
    utc_lpUniversalTime As utc_SYSTEMTIME _
) As Long

#Else

Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" ( _
    utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION _
) As Long
Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" ( _
    utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, _
    utc_lpUniversalTime As utc_SYSTEMTIME, _
    utc_lpLocalTime As utc_SYSTEMTIME _
) As Long
Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" ( _
    utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, _
    utc_lpLocalTime As utc_SYSTEMTIME, _
    utc_lpUniversalTime As utc_SYSTEMTIME _
) As Long

#End If

#If Mac Then

#If VBA7 Then
Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As LongPtr
End Type

#Else

Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As Long
End Type

#End If

#Else

Private Type utc_SYSTEMTIME
    utc_wYear As Integer
    utc_wMonth As Integer
    utc_wDayOfWeek As Integer
    utc_wDay As Integer
    utc_wHour As Integer
    utc_wMinute As Integer
    utc_wSecond As Integer
    utc_wMilliseconds As Integer
End Type

Private Type utc_TIME_ZONE_INFORMATION
    utc_Bias As Long
    utc_StandardName(0 To 31) As Integer
    utc_StandardDate As utc_SYSTEMTIME
    utc_StandardBias As Long
    utc_DaylightName(0 To 31) As Integer
    utc_DaylightDate As utc_SYSTEMTIME
    utc_DaylightBias As Long
End Type

#End If
' === End VBA-UTC

#If Mac Then
#ElseIf VBA7 Then

Private Declare PtrSafe Sub json_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    json_MemoryDestination As Any, _
    json_MemorySource As Any, _
    ByVal json_ByteLength As Long _
)

#Else

Private Declare Sub json_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    json_MemoryDestination As Any, _
    json_MemorySource As Any, _
    ByVal json_ByteLength As Long _
)

#End If

Private Type json_Options
    ' VBA only stores 15 significant digits, so any numbers larger than that are truncated
    ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
    ' See: http://support.microsoft.com/kb/269370
    '
    ' By default, VBA-JSON will use String for numbers longer than 15 characters that contain only digits
    ' to override set `JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True`
    UseDoubleForLargeNumbers As Boolean

    ' The JSON standard requires object keys to be quoted (" or '), use this option to allow unquoted keys
    AllowUnquotedKeys As Boolean

    ' The solidus (/) is not required to be escaped, use this option to escape them as \/ in Serialize
    EscapeSolidus As Boolean
End Type
Public JsonOptions As json_Options

' ============================================= '
' Public Methods
' ============================================= '

''
' Convert JSON string to object (Dictionary/Collection)
'
' @method Parse
' @param {String} json_String
' @return {Object} (Dictionary or Collection)
' @throws 10001 - JSON parse error
''
Public Function Parse( _
    ByVal JsonString As String _
) As Object
    ' Declare local variables.
    Dim json_Index As Long

    json_Index = 1

    ' Remove vbCr, vbLf, and vbTab from json_String
    JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, VBA.vbNullString), _
        VBA.vbLf, VBA.vbNullString), VBA.vbTab, VBA.vbNullString)

    Call json_SkipSpaces(JsonString, json_Index)
    Select Case VBA.Mid$(JsonString, json_Index, 1)
        Case "{"
            Set Parse = json_ParseObject(JsonString, json_Index)
        Case "["
            Set Parse = json_ParseArray(JsonString, json_Index)
        Case Else
            ' Error: Invalid JSON string
            Call VBA.Err.Raise(10001, "JSONConverter", json_ParseErrorMessage(JsonString, json_Index, "Expecting '{' or '['"))
    End Select
End Function

''
' Convert object (Dictionary/Collection/Array) to JSON
'
' @method Serialize
' @param {Variant} JsonValue (Dictionary, Collection, or Array)
' @param {Integer|String} Whitespace "Pretty" print json with given number of spaces per indentation (Integer) or given string
' @return {String}
''
Public Function Serialize( _
    ByVal JsonValue As Variant, _
    Optional ByVal Whitespace As Variant, _
    Optional ByVal json_CurrentIndentation As Long = 0 _
) As String
    ' Declare local variables.
    Dim json_buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    Dim json_Index As Long
    Dim json_LBound As Long
    Dim json_UBound As Long
    Dim json_IsFirstItem As Boolean
    Dim json_Index2D As Long
    Dim json_LBound2D As Long
    Dim json_UBound2D As Long
    Dim json_IsFirstItem2D As Boolean
    Dim json_Key As Variant
    Dim json_Value As Variant
    Dim json_DateStr As String
    Dim json_Converted As String
    Dim json_SkipItem As Boolean
    Dim json_PrettyPrint As Boolean
    Dim json_Indentation As String
    Dim json_InnerIndentation As String

    json_LBound = -1
    json_UBound = -1
    json_IsFirstItem = True
    json_LBound2D = -1
    json_UBound2D = -1
    json_IsFirstItem2D = True
    json_PrettyPrint = Not IsMissing(Whitespace)

    Select Case VBA.VarType(JsonValue)
        Case VBA.vbNull
            Serialize = "null"
        Case VBA.vbDate
            ' Date
            json_DateStr = ConvertToIso(CDate(JsonValue))

            Serialize = """" & json_DateStr & """"
        Case VBA.vbString
            ' String (or large number encoded as string)
            If (Not JsonOptions.UseDoubleForLargeNumbers) And json_StringIsLargeNumber(JsonValue) Then
                Serialize = JsonValue
            Else
                Serialize = """" & json_Encode(JsonValue) & """"
            End If
        Case VBA.vbBoolean
            If JsonValue Then
                Serialize = "true"
            Else
                Serialize = "false"
            End If
        Case VBA.vbArray To VBA.vbArray + VBA.vbByte
            If json_PrettyPrint Then
                If VBA.VarType(Whitespace) = VBA.vbString Then
                    json_Indentation = VBA.String$(json_CurrentIndentation + 1, Whitespace)
                    json_InnerIndentation = VBA.String$(json_CurrentIndentation + 2, Whitespace)
                Else
                    json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * Whitespace)
                    json_InnerIndentation = VBA.Space$((json_CurrentIndentation + 2) * Whitespace)
                End If
            End If

            ' Array
            Call json_BufferAppend(json_buffer, "[", json_BufferPosition, json_BufferLength)

            On Error Resume Next

            json_LBound = LBound(JsonValue, 1)
            json_UBound = UBound(JsonValue, 1)
            json_LBound2D = LBound(JsonValue, 2)
            json_UBound2D = UBound(JsonValue, 2)

            If (json_LBound >= 0) And (json_UBound >= 0) Then
                For json_Index = json_LBound To json_UBound
                    If json_IsFirstItem Then
                        json_IsFirstItem = False
                    Else
                        ' Append comma to previous line
                        Call json_BufferAppend(json_buffer, ",", json_BufferPosition, json_BufferLength)
                    End If

                    If (json_LBound2D >= 0) And (json_UBound2D >= 0) Then
                        ' 2D Array
                        If json_PrettyPrint Then
                            Call json_BufferAppend(json_buffer, VBA.vbNewLine, json_BufferPosition, json_BufferLength)
                        End If
                        Call json_BufferAppend(json_buffer, json_Indentation & "[", json_BufferPosition, json_BufferLength)

                        For json_Index2D = json_LBound2D To json_UBound2D
                            If json_IsFirstItem2D Then
                                json_IsFirstItem2D = False
                            Else
                                Call json_BufferAppend(json_buffer, ",", json_BufferPosition, json_BufferLength)
                            End If

                            json_Converted = Serialize(JsonValue(json_Index, json_Index2D), Whitespace, json_CurrentIndentation + 2)

                            ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                            If json_Converted = VBA.vbNullString Then
                                ' (nest to only check if converted = VBA.vbNullString)
                                If json_IsUndefined(JsonValue(json_Index, json_Index2D)) Then
                                    json_Converted = "null"
                                End If
                            End If

                            If json_PrettyPrint Then
                                json_Converted = VBA.vbNewLine & json_InnerIndentation & json_Converted
                            End If

                            Call json_BufferAppend(json_buffer, json_Converted, json_BufferPosition, json_BufferLength)
                        Next

                        If json_PrettyPrint Then
                            Call json_BufferAppend(json_buffer, VBA.vbNewLine, json_BufferPosition, json_BufferLength)
                        End If

                        Call json_BufferAppend(json_buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength)
                        json_IsFirstItem2D = True
                    Else
                        ' 1D Array
                        json_Converted = Serialize(JsonValue(json_Index), Whitespace, json_CurrentIndentation + 1)

                        ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                        If json_Converted = VBA.vbNullString Then
                            ' (nest to only check if converted = VBA.vbNullString)
                            If json_IsUndefined(JsonValue(json_Index)) Then
                                json_Converted = "null"
                            End If
                        End If

                        If json_PrettyPrint Then
                            json_Converted = VBA.vbNewLine & json_Indentation & json_Converted
                        End If

                        Call json_BufferAppend(json_buffer, json_Converted, json_BufferPosition, json_BufferLength)
                    End If
                Next
            End If

            On Error GoTo 0

            If json_PrettyPrint Then
                Call json_BufferAppend(json_buffer, VBA.vbNewLine, json_BufferPosition, json_BufferLength)

                If VBA.VarType(Whitespace) = VBA.vbString Then
                    json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                Else
                    json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                End If
            End If

            Call json_BufferAppend(json_buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength)

            Serialize = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)

        ' Dictionary or Collection
        Case VBA.vbObject
            If json_PrettyPrint Then
                If VBA.VarType(Whitespace) = VBA.vbString Then
                    json_Indentation = VBA.String$(json_CurrentIndentation + 1, Whitespace)
                Else
                    json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * Whitespace)
                End If
            End If

            ' Dictionary
            If VBA.TypeName(JsonValue) = "Dictionary" Then
                Call json_BufferAppend(json_buffer, "{", json_BufferPosition, json_BufferLength)
                For Each json_Key In JsonValue
                    ' For Objects, undefined (Empty/Nothing) is not added to object
                    json_Converted = Serialize(JsonValue(json_Key), Whitespace, json_CurrentIndentation + 1)
                    If json_Converted = VBA.vbNullString Then
                        json_SkipItem = json_IsUndefined(JsonValue(json_Key))
                    Else
                        json_SkipItem = False
                    End If

                    If Not json_SkipItem Then
                        If json_IsFirstItem Then
                            json_IsFirstItem = False
                        Else
                            Call json_BufferAppend(json_buffer, ",", json_BufferPosition, json_BufferLength)
                        End If

                        If json_PrettyPrint Then
                            json_Converted = VBA.vbNewLine & json_Indentation & """" & json_Key & """: " & json_Converted
                        Else
                            json_Converted = """" & json_Key & """:" & json_Converted
                        End If

                        Call json_BufferAppend(json_buffer, json_Converted, json_BufferPosition, json_BufferLength)
                    End If
                Next

                If json_PrettyPrint Then
                    Call json_BufferAppend(json_buffer, VBA.vbNewLine, json_BufferPosition, json_BufferLength)

                    If VBA.VarType(Whitespace) = VBA.vbString Then
                        json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                    Else
                        json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                    End If
                End If

                Call json_BufferAppend(json_buffer, json_Indentation & "}", json_BufferPosition, json_BufferLength)

            ' Collection
            ElseIf VBA.TypeName(JsonValue) = "Collection" Then
                Call json_BufferAppend(json_buffer, "[", json_BufferPosition, json_BufferLength)
                For Each json_Value In JsonValue
                    If json_IsFirstItem Then
                        json_IsFirstItem = False
                    Else
                        Call json_BufferAppend(json_buffer, ",", json_BufferPosition, json_BufferLength)
                    End If

                    json_Converted = Serialize(json_Value, Whitespace, json_CurrentIndentation + 1)

                    ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                    If json_Converted = VBA.vbNullString Then
                        ' (nest to only check if converted = VBA.vbNullString)
                        If json_IsUndefined(json_Value) Then
                            json_Converted = "null"
                        End If
                    End If

                    If json_PrettyPrint Then
                        json_Converted = VBA.vbNewLine & json_Indentation & json_Converted
                    End If

                    Call json_BufferAppend(json_buffer, json_Converted, json_BufferPosition, json_BufferLength)
                Next

                If json_PrettyPrint Then
                    Call json_BufferAppend(json_buffer, VBA.vbNewLine, json_BufferPosition, json_BufferLength)

                    If VBA.VarType(Whitespace) = VBA.vbString Then
                        json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                    Else
                        json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                    End If
                End If

                Call json_BufferAppend(json_buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength)
            End If

            Serialize = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)
        Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
            ' Number (use decimals for numbers)
            Serialize = VBA.Replace(JsonValue, ",", ".")
        Case Else
            ' vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
            ' Use VBA's built-in to-string
            On Error Resume Next
            Serialize = JsonValue
            On Error GoTo 0
    End Select
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function json_ParseObject( _
    ByRef json_String As String, _
    ByRef json_Index As Long _
) As Dictionary
    ' Declare local variables.
    Dim json_Key As String
    Dim json_NextChar As String

    Set json_ParseObject = New Dictionary
    Call json_SkipSpaces(json_String, json_Index)
    If VBA.Mid$(json_String, json_Index, 1) <> "{" Then
        Call VBA.Err.Raise(10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '{'"))
    Else
        json_Index = json_Index + 1

        Do
            Call json_SkipSpaces(json_String, json_Index)
            If VBA.Mid$(json_String, json_Index, 1) = "}" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                Call json_SkipSpaces(json_String, json_Index)
            End If

            json_Key = json_ParseKey(json_String, json_Index)
            json_NextChar = json_Peek(json_String, json_Index)
            If (json_NextChar = "[") Or (json_NextChar = "{") Then
                Set json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
            Else
                json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
            End If
        Loop
    End If
End Function

Private Function json_ParseArray( _
    ByRef json_String As String, _
    ByRef json_Index As Long _
) As Collection
    Set json_ParseArray = New Collection

    Call json_SkipSpaces(json_String, json_Index)
    If VBA.Mid$(json_String, json_Index, 1) <> "[" Then
        Call VBA.Err.Raise(10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '['"))
    Else
        json_Index = json_Index + 1

        Do
            Call json_SkipSpaces(json_String, json_Index)
            If VBA.Mid$(json_String, json_Index, 1) = "]" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                Call json_SkipSpaces(json_String, json_Index)
            End If

            Call json_ParseArray.Add(json_ParseValue(json_String, json_Index))
        Loop
    End If
End Function

Private Function json_ParseValue( _
    ByRef json_String As String, _
    ByRef json_Index As Long _
) As Variant
    Call json_SkipSpaces(json_String, json_Index)
    Select Case VBA.Mid$(json_String, json_Index, 1)
        Case "{"
            Set json_ParseValue = json_ParseObject(json_String, json_Index)
        Case "["
            Set json_ParseValue = json_ParseArray(json_String, json_Index)
        Case """", "'"
            json_ParseValue = json_ParseString(json_String, json_Index)
        Case Else
            If VBA.Mid$(json_String, json_Index, 4) = "true" Then
                json_ParseValue = True
                json_Index = json_Index + 4
            ElseIf VBA.Mid$(json_String, json_Index, 5) = "false" Then
                json_ParseValue = False
                json_Index = json_Index + 5
            ElseIf VBA.Mid$(json_String, json_Index, 4) = "null" Then
                json_ParseValue = Null
                json_Index = json_Index + 4
            ElseIf VBA.InStr("+-0123456789", VBA.Mid$(json_String, json_Index, 1)) Then
                json_ParseValue = json_ParseNumber(json_String, json_Index)
            Else
                Call VBA.Err.Raise(10001, "JSONConverter", _
                    json_ParseErrorMessage(json_String, json_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['"))
            End If
    End Select
End Function

Private Function json_ParseString( _
    ByRef json_String As String, _
    ByRef json_Index As Long _
) As String
    ' Declare local variables.
    Dim json_Quote As String
    Dim json_Char As String
    Dim json_Code As String
    Dim json_buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long

    Call json_SkipSpaces(json_String, json_Index)

    ' Store opening quote to look for matching closing quote
    json_Quote = VBA.Mid$(json_String, json_Index, 1)
    json_Index = json_Index + 1

    Do While (json_Index > 0) And (json_Index <= VBA.Len(json_String))
        json_Char = VBA.Mid$(json_String, json_Index, 1)

        Select Case json_Char
            Case "\"
                ' Escaped string, \\, or \/
                json_Index = json_Index + 1
                json_Char = VBA.Mid$(json_String, json_Index, 1)

                Select Case json_Char
                    Case """", "\", "/", "'"
                        Call json_BufferAppend(json_buffer, json_Char, json_BufferPosition, json_BufferLength)
                        json_Index = json_Index + 1
                    Case "b"
                        Call json_BufferAppend(json_buffer, VBA.vbBack, json_BufferPosition, json_BufferLength)
                        json_Index = json_Index + 1
                    Case "f"
                        Call json_BufferAppend(json_buffer, VBA.vbFormFeed, json_BufferPosition, json_BufferLength)
                        json_Index = json_Index + 1
                    Case "n"
                        Call json_BufferAppend(json_buffer, VBA.vbCrLf, json_BufferPosition, json_BufferLength)
                        json_Index = json_Index + 1
                    Case "r"
                        Call json_BufferAppend(json_buffer, VBA.vbCr, json_BufferPosition, json_BufferLength)
                        json_Index = json_Index + 1
                    Case "t"
                        Call json_BufferAppend(json_buffer, VBA.vbTab, json_BufferPosition, json_BufferLength)
                        json_Index = json_Index + 1
                    Case "u"
                        ' Unicode character escape (e.g. \u00a9 = Copyright)
                        json_Index = json_Index + 1
                        json_Code = VBA.Mid$(json_String, json_Index, 4)
                        Call json_BufferAppend(json_buffer, VBA.ChrW(VBA.Val("&h" + json_Code)), json_BufferPosition, json_BufferLength)
                        json_Index = json_Index + 4
                End Select
            Case json_Quote
                json_ParseString = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)
                json_Index = json_Index + 1
                Exit Function
            Case Else
                Call json_BufferAppend(json_buffer, json_Char, json_BufferPosition, json_BufferLength)
                json_Index = json_Index + 1
        End Select
    Loop
End Function

Private Function json_ParseNumber( _
    ByRef json_String As String, _
    ByRef json_Index As Long _
) As Variant
    ' Declare local variables.
    Dim json_Char As String
    Dim json_Value As String
    Dim json_IsLargeNumber As Boolean

    Call json_SkipSpaces(json_String, json_Index)

    Do While (json_Index > 0) And (json_Index <= VBA.Len(json_String))
        json_Char = VBA.Mid$(json_String, json_Index, 1)

        If VBA.InStr("+-0123456789.eE", json_Char) Then
            ' Unlikely to have massive number, so use simple append rather than buffer here
            json_Value = json_Value & json_Char
            json_Index = json_Index + 1
        Else
            ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
            ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
            ' See: http://support.microsoft.com/kb/269370
            '
            ' Fix: Parse -> String, Convert -> String longer than 15/16 characters containing only numbers and decimal points -> Number
            ' (decimal doesn't factor into significant digit count, so if present check for 15 digits + decimal = 16)
            json_IsLargeNumber = VBA.IIf(VBA.InStr(json_Value, "."), VBA.Len(json_Value) >= 17, VBA.Len(json_Value) >= 16)
            If (Not JsonOptions.UseDoubleForLargeNumbers) And json_IsLargeNumber Then
                json_ParseNumber = json_Value
            Else
                ' VBA.Val does not use regional settings, so guard for comma is not needed
                json_ParseNumber = VBA.Val(json_Value)
            End If
            Exit Function
        End If
    Loop
End Function

Private Function json_ParseKey( _
    ByRef json_String As String, _
    ByRef json_Index As Long _
) As String
    ' Declare local variables.
    Dim json_Char As String

    ' Parse key with single or double quotes
    If (VBA.Mid$(json_String, json_Index, 1) = """") Or (VBA.Mid$(json_String, json_Index, 1) = "'") Then
        json_ParseKey = json_ParseString(json_String, json_Index)
    ElseIf JsonOptions.AllowUnquotedKeys Then
        Do While (json_Index > 0) And (json_Index <= VBA.Len(json_String))
            json_Char = VBA.Mid$(json_String, json_Index, 1)
            If (json_Char <> " ") And (json_Char <> ":") Then
                json_ParseKey = json_ParseKey & json_Char
                json_Index = json_Index + 1
            Else
                Exit Do
            End If
        Loop
    Else
        Call VBA.Err.Raise(10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '""' or '''"))
    End If

    ' Check for colon and skip if present or throw if not present
    Call json_SkipSpaces(json_String, json_Index)
    If VBA.Mid$(json_String, json_Index, 1) <> ":" Then
        Call VBA.Err.Raise(10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting ':'"))
    Else
        json_Index = json_Index + 1
    End If
End Function

Private Function json_IsUndefined( _
    ByVal json_Value As Variant _
) As Boolean
    ' Empty / Nothing -> undefined
    Select Case VBA.VarType(json_Value)
        Case VBA.vbEmpty
            json_IsUndefined = True
        Case VBA.vbObject
            Select Case VBA.TypeName(json_Value)
                Case "Empty", "Nothing"
                    json_IsUndefined = True
            End Select
    End Select
End Function

Private Function json_Encode( _
    ByVal json_Text As Variant _
) As String
    ' Reference: http://www.ietf.org/rfc/rfc4627.txt
    ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab

    ' Declare local variables.
    Dim json_Index As Long
    Dim json_Char As String
    Dim json_AscCode As Long
    Dim json_buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long

    For json_Index = 1 To VBA.Len(json_Text)
        json_Char = VBA.Mid$(json_Text, json_Index, 1)
        json_AscCode = VBA.AscW(json_Char)

        ' When AscW returns a negative number, it returns the twos complement form of that number.
        ' To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
        ' https://support.microsoft.com/en-us/kb/272138
        If json_AscCode < 0 Then
            json_AscCode = json_AscCode + 65536
        End If

        ' From spec, ", \, and control characters must be escaped (solidus is optional)

        Select Case json_AscCode
            Case 34
                ' " -> 34 -> \"
                json_Char = "\"""
            Case 92
                ' \ -> 92 -> \\
                json_Char = "\\"
            Case 47
                ' / -> 47 -> \/ (optional)
                If JsonOptions.EscapeSolidus Then
                    json_Char = "\/"
                End If
            Case 8
                ' backspace -> 8 -> \b
                json_Char = "\b"
            Case 12
                ' form feed -> 12 -> \f
                json_Char = "\f"
            Case 10
                ' line feed -> 10 -> \n
                json_Char = "\n"
            Case 13
                ' carriage return -> 13 -> \r
                json_Char = "\r"
            Case 9
                ' tab -> 9 -> \t
                json_Char = "\t"
            Case 0 To 31, 127 To 65535
                ' Non-ascii characters -> convert to 4-digit hex
                json_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_AscCode), 4)
        End Select

        Call json_BufferAppend(json_buffer, json_Char, json_BufferPosition, json_BufferLength)
    Next

    json_Encode = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)
End Function

Private Function json_Peek( _
    ByRef json_String As String, _
    ByVal json_Index As Long, _
    Optional ByVal json_NumberOfCharacters As Long = 1 _
) As String
    ' "Peek" at the next number of characters without incrementing json_Index (ByVal instead of ByRef)
    Call json_SkipSpaces(json_String, json_Index)
    json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)
End Function

Private Sub json_SkipSpaces( _
    ByRef json_String As String, _
    ByRef json_Index As Long _
)
    ' Increment index to skip over spaces
    Do While (json_Index > 0) And (json_Index <= VBA.Len(json_String)) And (VBA.Mid$(json_String, json_Index, 1) = " ")
        json_Index = json_Index + 1
    Loop
End Sub

Private Function json_StringIsLargeNumber( _
    ByRef json_String As Variant _
) As Boolean
    ' Check if the given string is considered a "large number"
    ' (See json_ParseNumber)

    ' Declare local variables.
    Dim json_Length As Long
    Dim json_CharIndex As Long
    Dim json_CharCode As String
    Dim json_Index As Long

    json_Length = VBA.Len(json_String)

    ' Length with be at least 16 characters and assume will be less than 100 characters
    If (json_Length >= 16) And (json_Length <= 100) Then
        json_StringIsLargeNumber = True

        For json_CharIndex = 1 To json_Length
            json_CharCode = VBA.Asc(VBA.Mid$(json_String, json_CharIndex, 1))
            Select Case json_CharCode
                ' Look for .|0-9|E|e
                Case 46, 48 To 57, 69, 101
                    ' Continue through characters
                Case Else
                    json_StringIsLargeNumber = False
                    Exit Function
            End Select
        Next
    End If
End Function

Private Function json_ParseErrorMessage( _
    ByRef json_String As String, _
    ByRef json_Index As Long, _
    ByRef ErrorMessage As String _
)
    ' Provide detailed parse error message, including details of where and what occurred
    '
    ' Example:
    ' Error parsing JSON:
    ' {"abcde":True}
    '          ^
    ' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['

    ' Declare local variables.
    Dim json_StartIndex As Long
    Dim json_StopIndex As Long

    ' Include 10 characters before and after error (if possible)
    json_StartIndex = json_Index - 10
    json_StopIndex = json_Index + 10
    If json_StartIndex <= 0 Then
        json_StartIndex = 1
    End If
    If json_StopIndex > VBA.Len(json_String) Then
        json_StopIndex = VBA.Len(json_String)
    End If

    json_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
        VBA.Mid$(json_String, json_StartIndex, json_StopIndex - json_StartIndex + 1) & VBA.vbNewLine & _
        VBA.Space$(json_Index - json_StartIndex) & "^" & VBA.vbNewLine & _
        ErrorMessage
End Function

Private Sub json_BufferAppend( _
    ByRef json_buffer As String, _
    ByRef json_Append As Variant, _
    ByRef json_BufferPosition As Long, _
    ByRef json_BufferLength As Long _
)
#If Mac Then
    json_buffer = json_buffer & json_Append
#Else
    ' VBA can be slow to append strings due to allocating a new string for each append
    ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
    '
    ' Example:
    ' Buffer: "abc  "
    ' Append: "def"
    ' Buffer Position: 3
    ' Buffer Length: 5
    '
    ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
    ' Buffer: "abc       "
    ' Buffer Length: 10
    '
    ' Copy memory for "def" into buffer at position 3 (0-based)
    ' Buffer: "abcdef    "
    '
    ' Approach based on cStringBuilder from vbAccelerator
    ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp

    ' Declare local variables.
    Dim json_AppendLength As Long
    Dim json_LengthPlusPosition As Long
    Dim json_TemporaryLength As Long

    json_AppendLength = VBA.LenB(json_Append)
    json_LengthPlusPosition = json_AppendLength + json_BufferPosition

    If json_LengthPlusPosition > json_BufferLength Then
        ' Appending would overflow buffer, add chunks until buffer is long enough

        json_TemporaryLength = json_BufferLength
        Do While json_TemporaryLength < json_LengthPlusPosition
            ' Initially, initialize string with 255 characters,
            ' then add large chunks (8192) after that
            '
            ' Size: # Characters x 2 bytes / character
            If json_TemporaryLength = 0 Then
                json_TemporaryLength = json_TemporaryLength + 510
            Else
                json_TemporaryLength = json_TemporaryLength + 16384
            End If
        Loop

        json_buffer = json_buffer & VBA.Space$((json_TemporaryLength - json_BufferLength) \ 2)
        json_BufferLength = json_TemporaryLength
    End If

    ' Copy memory from append to buffer at buffer position
    Call json_CopyMemory( _
        ByVal json_UnsignedAdd(StrPtr(json_buffer), _
        json_BufferPosition), _
        ByVal StrPtr(json_Append), _
        json_AppendLength)

    json_BufferPosition = json_BufferPosition + json_AppendLength
#End If
End Sub

Private Function json_BufferToString( _
    ByRef json_buffer As String, _
    ByVal json_BufferPosition As Long, _
    ByVal json_BufferLength As Long _
) As String
#If Mac Then
    json_BufferToString = json_buffer
#Else
    If json_BufferPosition > 0 Then
        json_BufferToString = VBA.Left$(json_buffer, json_BufferPosition \ 2)
    End If
#End If
End Function

#If VBA7 Then
Private Function json_UnsignedAdd( _
    json_Start As LongPtr, _
    json_Increment As Long _
) As LongPtr
#Else
Private Function json_UnsignedAdd( _
    json_Start As Long, _
    json_Increment As Long _
) As Long
#End If

    If json_Start And &H80000000 Then
        json_UnsignedAdd = json_Start + json_Increment
    ElseIf (json_Start Or &H80000000) < -json_Increment Then
        json_UnsignedAdd = json_Start + json_Increment
    Else
        json_UnsignedAdd = (json_Start + &H80000000) + (json_Increment + &H80000000)
    End If
End Function

''
' VBA-UTC v1.0.3
' (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter
'
' UTC/ISO 8601 Converter for VBA
'
' Errors:
' 10011 - UTC parsing error
' 10012 - UTC conversion error
' 10013 - ISO 8601 parsing error
' 10014 - ISO 8601 conversion error
'
' @module UtcConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' (Declarations moved to top)

' ============================================= '
' Public Methods
' ============================================= '

''
' Parse UTC date to local date
'
' @method ParseUtc
' @param {Date} UtcDate
' @return {Date} Local date
' @throws 10011 - UTC parsing error
''
Public Function ParseUtc( _
    ByVal utc_UtcDate As Date _
) As Date
    On Error GoTo utc_ErrorHandling

#If Mac Then
    ParseUtc = utc_ConvertDate(utc_UtcDate)
#Else

    ' Declare local variables.
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_LocalDate As utc_SYSTEMTIME

    Call utc_GetTimeZoneInformation(utc_TimeZoneInfo)
    Call utc_SystemTimeToTzSpecificLocalTime(utc_TimeZoneInfo, utc_DateToSystemTime(utc_UtcDate), utc_LocalDate)

    ParseUtc = utc_SystemTimeToDate(utc_LocalDate)
#End If

    Exit Function

utc_ErrorHandling:
    Call VBA.Err.Raise(10011, "UtcConverter.ParseUtc", "UTC parsing error: " & VBA.Err.Number & " - " & VBA.Err.Description)
End Function

''
' Convert local date to UTC date
'
' @method ConvertToUrc
' @param {Date} utc_LocalDate
' @return {Date} UTC date
' @throws 10012 - UTC conversion error
''
Public Function ConvertToUtc( _
    ByVal utc_LocalDate As Date _
) As Date
    On Error GoTo utc_ErrorHandling

#If Mac Then
    ConvertToUtc = utc_ConvertDate(utc_LocalDate, utc_ConvertToUtc:=True)
#Else

    ' Declare local variables.
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_UtcDate As utc_SYSTEMTIME

    Call utc_GetTimeZoneInformation(utc_TimeZoneInfo)
    Call utc_TzSpecificLocalTimeToSystemTime(utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate)

    ConvertToUtc = utc_SystemTimeToDate(utc_UtcDate)
#End If

    Exit Function

utc_ErrorHandling:
    Call VBA.Err.Raise(10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & VBA.Err.Number & " - " & VBA.Err.Description)
End Function

''
' Parse ISO 8601 date string to local date
'
' @method ParseIso
' @param {Date} utc_IsoString
' @return {Date} Local date
' @throws 10013 - ISO 8601 parsing error
''
Public Function ParseIso( _
    ByRef utc_IsoString As String _
) As Date
    On Error GoTo utc_ErrorHandling

    ' Declare local variables.
    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String
    Dim utc_OffsetIndex As Long
    Dim utc_HasOffset As Boolean
    Dim utc_NegativeOffset As Boolean
    Dim utc_OffsetParts() As String
    Dim utc_Offset As Date

    utc_Parts = VBA.Split(utc_IsoString, "T")
    utc_DateParts = VBA.Split(utc_Parts(0), "-")
    ParseIso = VBA.DateSerial(VBA.CInt(utc_DateParts(0)), VBA.CInt(utc_DateParts(1)), VBA.CInt(utc_DateParts(2)))

    If UBound(utc_Parts) > 0 Then
        If VBA.InStr(utc_Parts(1), "Z") Then
            utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts(1), "Z", VBA.vbNullString), ":")
        Else
            utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "+")
            If utc_OffsetIndex = 0 Then
                utc_NegativeOffset = True
                utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "-")
            End If

            If utc_OffsetIndex > 0 Then
                utc_HasOffset = True
                utc_TimeParts = VBA.Split(VBA.Left$(utc_Parts(1), utc_OffsetIndex - 1), ":")
                utc_OffsetParts = VBA.Split(VBA.Right$(utc_Parts(1), VBA.Len(utc_Parts(1)) - utc_OffsetIndex), ":")

                Select Case UBound(utc_OffsetParts)
                    Case 0
                        utc_Offset = VBA.TimeSerial(VBA.CInt(utc_OffsetParts(0)), 0, 0)
                    Case 1
                        utc_Offset = VBA.TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), 0)
                    Case 2
                        ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
                        utc_Offset = VBA.TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), Int(VBA.Val(utc_OffsetParts(2))))
                End Select

                If utc_NegativeOffset Then
                    utc_Offset = -utc_Offset
                End If
            Else
                utc_TimeParts = VBA.Split(utc_Parts(1), ":")
            End If
        End If

        Select Case UBound(utc_TimeParts)
        Case 0
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), 0, 0)
        Case 1
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), 0)
        Case 2
            ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), Int(VBA.Val(utc_TimeParts(2))))
        End Select

        ParseIso = ParseUtc(ParseIso)

        If utc_HasOffset Then
            ParseIso = ParseIso + utc_Offset
        End If
    End If

    Exit Function

utc_ErrorHandling:
    Call VBA.Err.Raise(10013, "UtcConverter.ParseIso", "ISO 8601 parsing error for " _
        & utc_IsoString & ": " & VBA.Err.Number & " - " & VBA.Err.Description)
End Function

''
' Convert local date to ISO 8601 string
'
' @method ConvertToIso
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' @throws 10014 - ISO 8601 conversion error
''
Public Function ConvertToIso( _
    ByVal utc_LocalDate As Date _
) As String
    On Error GoTo utc_ErrorHandling

    ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")

    Exit Function

utc_ErrorHandling:
    Call VBA.Err.Raise(10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & VBA.Err.Number & " - " & VBA.Err.Description)
End Function

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Then

Private Function utc_ConvertDate( _
    ByVal utc_Value As Date, _
    Optional ByRef utc_ConvertToUtc As Boolean = False _
) As Date
    ' Declare local variables.
    Dim utc_ShellCommand As String
    Dim utc_Result As utc_ShellResult
    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String

    If utc_ConvertToUtc Then
        utc_ShellCommand = "date -ur `date -jf '%Y-%m-%d %H:%M:%S' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & "' " & _
            " +'%s'` +'%Y-%m-%d %H:%M:%S'"
    Else
        utc_ShellCommand = "date -jf '%Y-%m-%d %H:%M:%S %z' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & " +0000' " & _
            "+'%Y-%m-%d %H:%M:%S'"
    End If

    utc_Result = utc_ExecuteInShell(utc_ShellCommand)

    If utc_Result.utc_Output = VBA.vbNullString Then
        Call VBA.Err.Raise(10015, "UtcConverter.utc_ConvertDate", "'date' command failed")
    Else
        utc_Parts = Split(utc_Result.utc_Output, " ")
        utc_DateParts = Split(utc_Parts(0), "-")
        utc_TimeParts = Split(utc_Parts(1), ":")

        utc_ConvertDate = VBA.DateSerial(utc_DateParts(0), utc_DateParts(1), utc_DateParts(2)) + _
            VBA.TimeSerial(utc_TimeParts(0), utc_TimeParts(1), utc_TimeParts(2))
    End If
End Function

Private Function utc_ExecuteInShell( _
    ByRef utc_ShellCommand As String _
) As utc_ShellResult
    ' Declare local variables.
#If VBA7 Then
    Dim utc_File As LongPtr
    Dim utc_Read As LongPtr
#Else
    Dim utc_File As Long
    Dim utc_Read As Long
#End If

    Dim utc_Chunk As String

    On Error GoTo utc_ErrorHandling
    utc_File = utc_popen(utc_ShellCommand, "r")

    If utc_File = 0 Then
        Exit Function
    End If

    Do While utc_feof(utc_File) = 0
        utc_Chunk = VBA.Space$(50)
        utc_Read = utc_fread(utc_Chunk, 1, VBA.Len(utc_Chunk) - 1, utc_File)
        If utc_Read > 0 Then
            utc_Chunk = VBA.Left$(utc_Chunk, utc_Read)
            utc_ExecuteInShell.utc_Output = utc_ExecuteInShell.utc_Output & utc_Chunk
        End If
    Loop

utc_ErrorHandling:
    utc_ExecuteInShell.utc_ExitCode = utc_pclose(utc_File)
End Function

#Else

Private Function utc_DateToSystemTime( _
    ByVal utc_Value As Date _
) As utc_SYSTEMTIME
    utc_DateToSystemTime.utc_wYear = VBA.Year(utc_Value)
    utc_DateToSystemTime.utc_wMonth = VBA.Month(utc_Value)
    utc_DateToSystemTime.utc_wDay = VBA.Day(utc_Value)
    utc_DateToSystemTime.utc_wHour = VBA.Hour(utc_Value)
    utc_DateToSystemTime.utc_wMinute = VBA.Minute(utc_Value)
    utc_DateToSystemTime.utc_wSecond = VBA.Second(utc_Value)
    utc_DateToSystemTime.utc_wMilliseconds = 0
End Function

Private Function utc_SystemTimeToDate( _
    ByRef utc_Value As utc_SYSTEMTIME _
) As Date
    utc_SystemTimeToDate = VBA.DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + _
        VBA.TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)
End Function

#End If