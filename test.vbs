Option Explicit

Dim vString

' Initialize the msxml dom document object.
With CreateObject("MSXML2.DOMDocument.6.0")
	' Configure to load files asynchronously.
	.async = False

	' Load the build configuration xml file.
	Call .load("test.xml")

	vString = .selectSingleNode("test").Text

	Call WScript.Echo(vString)

	Call WScript.Echo(Len(vString))

	Call WScript.Echo(Left(vString, 2))
	Call WScript.Echo(Right(vString, 2))
End With
