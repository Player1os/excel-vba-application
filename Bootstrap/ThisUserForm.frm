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

' Requires Runtime

Private vOriginalApplicationLeft As Long
Private vOriginalApplicationTop As Long
Private vOriginalApplicationWidth As Long
Private vOriginalApplicationHeight As Long

Private Const vMinimumApplicationWidth As Long = 104
Private Const vMinimumApplicationHeight As Long = 30
Private Const vApplicationPadding As Long = 50
Private Const vWebBrowserPadding As Long = 4

Private Sub UserForm_Activate()
    ' Set the title and icon of the current window.
    Caption = Runtime.ProjectName()
    Call Runtime.SetActiveWindowIcon

    ' Populate the current window with standard controls and maximize it.
    Call Runtime.PopulateActiveWindowTitlebar
    Call Runtime.MaximizeActiveWindow

    ' Load the main HTML file as the basis of the pages to be displayed in the embedded web browser.
    With Runtime.FileSystemObject()
        Call ThisWebBrowser.Navigate(.BuildPath(.BuildPath(ThisWorkbook.Path, "Assets"), "Main.html") & "#" & Runtime.NavigatePath())
    End With
End Sub

Private Sub UserForm_Initialize()
    ' Load the excel application instance.
    With Application
        ' Store the original dimensions.
        vOriginalApplicationLeft = .Left
        vOriginalApplicationTop = .Top
        vOriginalApplicationWidth = .Width
        vOriginalApplicationHeight = .Height

        ' Shrink the application window.
        .Width = vMinimumApplicationWidth
        .Height = vMinimumApplicationHeight
    End With
End Sub

Private Sub UserForm_Terminate()
    ' Load the excel application instance.
    With Application
        ' Restore the original dimensions.
        .Left = vOriginalApplicationLeft
        .Top = vOriginalApplicationTop
        .Width = vOriginalApplicationWidth
        .Height = vOriginalApplicationHeight
    End With
End Sub

Private Sub UserForm_Layout()
    ' Load the excel application instance.
    With Application
        ' Check whether there is a need to move the application window.
        If _
            (.Left < (Left + vApplicationPadding)) _
            And ((.Left + .Width) > (Left + Width + vApplicationPadding)) _
            And (.Top < (Top + vApplicationPadding)) _
            And ((.Top + .Height) > (Top + Height + vApplicationPadding)) _
        Then
            Exit Sub
        End If

        ' Move the application window to the center of the user form.
        .Left = Me.Left + (Width - .Width) / 2
        .Top = Me.Top + (Height - .Height) / 2
    End With
End Sub

Private Sub UserForm_Resize()
    ' Resize the embedded web browser.
    With ThisWebBrowser
        .Width = InsideWidth + vWebBrowserPadding
        .Height = InsideHeight + vWebBrowserPadding
    End With
End Sub

Private Sub ThisWebBrowser_DocumentComplete( _
    ByVal pDisp As Object, _
    URL As Variant _
)
    ' Extract the fragment portion of the original url.
    Call Runtime.Navigate(Right(URL, Len(URL) - InStr(URL, "#")))
End Sub

Public Sub SetInnerHtml( _
    vHtmlText As String _
)
    ThisWebBrowser.Document.body.innerHtml = vHtmlText
End Sub
