VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibClsDBRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Introduction

' If you are working with classes in VBA (see Class Modules for more details) it is often useful to make one member of a class the default member. For example, in the Excel Range object, the default member is Value. This allows you to omit the member name and use code like the following:

'     Range("A1") = 1234
' ' is the same as
'     Range("A1").Value = 1234
' Because Value is the default member, it may be omitted in the code. Creating a default member of a class is also very useful (necessary, really) when you are working with a customized Collection class. (See Creating A Custom Collection Class for more information about custom Collection classes.) In this case, you would likely specify the Item method as the default member. This allows you to use code like the following:

'     V = Coll(2)
'     ' is the same as
'     V = Coll.Item(2)
' SectionBreak

' Creating A Default Member In VBA

' VBA does not directly support the creation of a default member of a class. That is, there is nothing in the VBA IDE that allows you to specify a default member. However, VBA does respect the default method if it is specified in a class. To specify a method as the default member, you need to Export the class module to a text file, edit that text file in NotePad or your favorite text editor, add an Attribute directive to the method, and then Import the text file back into the VBA Project.

' First, export the class module to a text file. In VBA, go to the File menu and choose Export File.... In the Save dialog that appears, navigate to some folder (it doesn't matter which folder), and save the class file as text with a cls extension. Next, select Remove... from the File menu and choose No in the Do you want to export? dialog. Next, open Notepad ( C:\Windows\Notepad.exe) or another text editor, and open the cls that you saved in the Export step. In the text file, go to the method that you want to make the default, and add the following line of code.

' Attribute Value.VB_UserMemId = 0
' An Attribute directive is an instruction to the compiler indicating various conditions for compilation. The Attribute directives are not visible in the VBA Editor and they cannot be added by the VBA Editor. You must use a text editor to add Attribute directives. If you are making the Value property the default member of your class, your code in Notepad should look similar to the following:

'     Property Get Value() As Long
'         Attribute Value.VB_UserMemId = 0
'         Value = Whatever
'     End Property
' You can make a Sub, Function, or Property the default member of the class, but only one procedure in the module may be the default member. Once you have added the Attribute directive to the text file, save the file and exit from NotePad. Now, in the VBA Editor, go to the File menu and choose Import File.... In the Open dialog that appears, navigate to the folder in which you saved the cls file and import it into VBA. Because Attribute directives are not visible in the VBA Editor, you will not see any changes in your code.

' Once the Attribute directive is in place, you can use code like the following:

'     Dim CC As CMyClassName
'     Set CC = New CMyClassName
'     CC.Value = 123
'     ' is the same as
'     CC = 123
' This page last updated: 2-May-2008
