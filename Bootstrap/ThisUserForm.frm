VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ThisUserForm
   Caption         =   "Main"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5910
   OleObjectBlob   =   "ThisUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ThisUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vFileSystemObject As New FileSystemObject

Private Sub UserForm_Initialize()
    With Application
        Me.Width = .Width - 8
        Me.Height = .Height - 12
    End With

    With ThisWebBrowser
        .Width = Me.InsideWidth + 4
        .Height = Me.InsideHeight + 4
    End With

    With vFileSystemObject
        Call ThisWebBrowser.Navigate(.BuildPath(.BuildPath(ThisWorkbook.Path, "Assets"), "Main.html"))
    End With
End Sub

Private Sub UserForm_Layout()
    With Application
        .Left = Me.Left
        .Top = Me.Top
    End With
End Sub

Private Sub ThisWebBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If VBA.InStr(URL, "#") = 0 Then
        Call ModController.Navigate(VBA.vbNullString)
    Else
        Call ModController.Navigate(VBA.Right(URL, VBA.Len(URL) - VBA.InStr(URL, "#")))
    End If
End Sub
