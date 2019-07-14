VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibClsSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vDictionary As Dictionary

' TODO: Refactor to not use class initialization but the New keyword.
Private Sub Class_Initialize()
    ' Initialize the underlying dictionary.
    Set vDictionary = New Dictionary
End Sub

Public Function Exists( _
    ByRef vValue As Variant _
) As Boolean
    ' Check in the underlying dictionary.
    Exists = vDictionary.Exists(vValue)
End Function

Public Function Add( _
    ByRef vValue As Variant _
) As CUtilitySet
    ' Verify that the value does not already exist in the set.
    If Not Exists(vValue) Then
        ' Add to the underlying dictionary.
        Call vDictionary.Add(vValue, Null)
    End If

    ' Return the current instance for chaining.
    Set Add = Me
End Function

Public Function Remove( _
    ByRef vValue As Variant _
) As CUtilitySet
    ' Remove from the underlying dictionary.
    Call vDictionary.Remove(vValue)

    ' Return the current instance for chaining.
    Set Remove = Me
End Function

' TODO: Refactor to Items, which returns an array.
Public Function ToCollection() As Collection
    ' Declare local variables.
    Dim vValue As Variant

    ' Initialize the result.
    Set ToCollection = New Collection

    ' Transfer values to the the result.
    For Each vValue In vDictionary
        Call ToCollection.Add(vValue)
    Next
End Function

' TODO: Populate from an array or collection.
