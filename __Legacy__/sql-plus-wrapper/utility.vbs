' Define a function for outputting syntax errors.
Sub pOutputUnexpectedSyntaxError( _
	ByRef vScriptSyntax, _
	ByRef vErrorMessage _
)
	Call WScript.StdOut.WriteLine("The following syntax is expected '" & vScriptSyntax & "'.")
	Call Err.Raise(1, vScriptName, vErrorMessage)
End Sub

' Define a function for reading text files.
Function pReadTextFile( _
	ByRef vFilePath _
)
	' Declare local variables.
	Dim vAdodbStream: Set vAdodbStream = CreateObject("ADODB.Stream")

	' Load all data from the stream.
	With vAdodbStream
		' Set the charset and open the stream.
		.Charset = "UTF-8"
		Call .Open()

		' Load the specified file and read its data from the stream.
		Call .LoadFromFile(vFilePath)
		pReadTextFile = .ReadText()

		' Close the stream.
		Call .Close()
	End With
End Function

' Define a function for writing text files.
Sub pWriteTextFile( _
	ByRef vFilePath, _
	ByRef vFileContentsString _
)
	' Declare local variables.
	Dim vAdodbStream: Set vAdodbStream = CreateObject("ADODB.Stream")

	' Store all data to the stream.
	With vAdodbStream
		' Set the charset and open the stream.
		.Charset = "UTF-8"
		Call .Open()

		' Write the data to the stream and save it to the specified file.
		Call .WriteText(vFileContentsString)
		Call .SaveToFile(vFilePath, 2)

		' Close the stream.
		Call .Close()
	End With
End Sub

' Define a function for loading environment variables from a specified file.
Sub pLoadEnvironmentVariables( _
	ByRef vFilePath _
)
	Dim vFileTextLines
	Dim vFileTextLine
	Dim vFields

	vFileTextLines = Split(pReadTextFile(vFilePath), vbLf)

	For Each vFileTextLine In vFileTextLines
		If vFileTextLine <> vbNullString Then
			vFields = Split(vFileTextLine, "=")
			vShell.Environment("PROCESS").Item(vFields(0)) = vFields(1)
		End If
	Next
End Sub

Sub pConvertCsvToExcel()
	With CreateObject("Excel.Application")
		With .Workbooks.Add()
			With .Worksheets(1)
				With .QueryTables.Add("TEXT;" & WScript.Arguments(0), .Cells(1, 1))
					.FieldNames = True
					.RowNumbers = False
					.FillAdjacentFormulas = False
					.PreserveFormatting = False
					.RefreshOnFileOpen = False
					.RefreshStyle = 0
					.SavePassword = False
					.SaveData = True
					.AdjustColumnWidth = False
					.RefreshPeriod = 0
					.TextFilePromptOnRefresh = False
					.TextFilePlatform = 65001
					.TextFileStartRow = 1
					.TextFileParseType = 1
					.TextFileTextQualifier = 1
					.TextFileConsecutiveDelimiter = False
					.TextFileTabDelimiter = False
					.TextFileSemicolonDelimiter = False
					.TextFileCommaDelimiter = True
					.TextFileSpaceDelimiter = False
					.TextFileColumnDataTypes = Array(2, 2, 2)
					.TextFileTrailingMinusNumbers = True
					Call .Refresh(False)
				End With
				Call .Cells.QueryTable.Delete()
			End With
			Call .SaveAs(Replace(WScript.Arguments(0), ".csv", ".xlsx"), 51)
			Call .Close(False)
		End With
		Call .Quit()
	End With
End Sub

Sub pDecodeBase64String()
	' Define constants.
	Const vValidCharacters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	Const vGroupLength = 4
	Const vBaseValue = 64

	' Declare local variables.
	Dim vInputString
	Dim vInputStringLength
	Dim vInputStringIndex
	Dim vInputStringCharacter
	Dim vGroupStringIndex
	Dim vByteCount
	Dim vByteValue
	Dim vPartialByteValue
	Dim vPartialOutputString
	Dim vOuputString

	' Load the input string from the standard input stream.
	vInputString = WScript.StdIn.ReadLine()

	' The source must consists of groups with the specified length in characters.
	vInputStringLength = Len(vInputString)
	If (vInputStringLength Mod vGroupLength) <> 0 Then
		Call Err.Raise(1, "Base64Decode", "The input string has an invalid length.")
	End If

	For vInputStringIndex = 1 To vInputStringLength Step vGroupLength
		' Each data group encodes up to 3 actual bytes.
		vByteCount = vGroupLength - 1
		vByteValue = 0

		For vGroupStringIndex = 1 To vGroupLength
			' Convert each character into 6 bits of data, and add it to an integer for temporary storage.
			' If an '=' character is encountered, there is one fewer data byte.
			' There can only be a maximum of 2 '=' in the whole string.

			vInputStringCharacter = Mid(vInputString, vInputStringIndex + vGroupStringIndex - 1, 1)

			If vInputStringCharacter = "=" Then
				vByteCount = vByteCount - 1
				vPartialByteValue = 0
			Else
				vPartialByteValue = InStr(1, vValidCharacters, vInputStringCharacter, vbBinaryCompare) - 1

				If vPartialByteValue = -1 Then
					Call Err.Raise(2, "Base64Decode", "An invalid character was found in the input string.")
				End If
			End If

			vByteValue = vBaseValue * vByteValue + vPartialByteValue
		Next

		' Hex splits the long to 6 groups with 4 bits.
		vByteValue = Hex(vByteValue)

		' Add leading zeros.
		vByteValue = String(6 - Len(vByteValue), "0") & vByteValue

		' Convert the 3 byte hex integer (6 chars) to 3 characters.
		vPartialOutputString = Chr(CByte("&H" & Mid(vByteValue, 1, 2))) _
			+ Chr(CByte("&H" & Mid(vByteValue, 3, 2))) _
			+ Chr(CByte("&H" & Mid(vByteValue, 5, 2)))

		' Add the proper amount of characters to the output string.
		vOuputString = vOuputString & Left(vPartialOutputString, vByteCount)
	Next

	' Write the output string to the standard output stream.
	Call WScript.StdOut.WriteLine(vOuputString)
End Sub
