Attribute VB_Name = "LibModMail"
Option Explicit
Option Private Module

' Requires MRuntime
' Requires MUtilityFile

Private Const vOlMailItem As Long = 0
Private Const vEmailAddressRegExpText As String = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"

Public Function ValidateAddress( _
    ByRef vEmailAddress As String _
) As Boolean
    With New RegExp
        .IgnoreCase = True
        .Global = True
        .Pattern = vEmailAddressRegExpText

        ValidateAddress = .Test(vEmailAddress)
    End With
End Function

' TODO: Remove the option to send in background.
' TODO: Think about using only HTML emails.
Public Sub Send( _
    ByRef vRecipientDictionary As Dictionary, _
    ByRef vSubject As String, _
    ByRef vBody As String, _
    Optional ByVal vSenderAddress As String = VBA.vbNullString, _
    Optional ByVal vAttachmentFilePaths As Collection = Nothing, _
    Optional ByVal vIsHtmlBody As Boolean = False, _
    Optional ByVal vIsBackgroundSendingEnabled As Boolean = True _
)
    ' Declare local variables.
    Dim vIsLoggedOn As Boolean
    Dim vOutlookApplication As Object
    Dim vMailItem As Object
    Dim vAccount As Object
    Dim vSelectedAccount As Object
    Dim vAttachmentFilePath As Variant

    ' Setup error handling.
    On Error GoTo HandleError:

    ' Initialize the outlook application.
    Set vOutlookApplication = CreateObject("Outlook.Application")
    Set vMailItem = vOutlookApplication.CreateItem(vOlMailItem)
    Call vOutlookApplication.Session.Logon
    vIsLoggedOn = True

    With vMailItem
        ' Check whether the-default sender account is to be used.
        If vSenderAddress <> VBA.vbNullString Then
            ' Search for the specified account to be used as the sender.
            For Each vAccount In vOutlookApplication.Session.Accounts
                If vAccount.DisplayName = vSenderAddress Then
                    Set vSelectedAccount = vAccount
                    Exit For
                End If
            Next

            ' If no matching account has been found, use the first available one.
            If vSelectedAccount Is Nothing Then
                Call MRuntime.RaiseError("MUtilityMail.Send", _
                    "No suitable account has been found with the specified sender address.")
            End If

            ' Set the selected account in the mail item.
            Set .SendUsingAccount = vSelectedAccount
        End If

        ' Parse the recepient data.
        If vRecipientDictionary.Exists("To") Then
            .To = vRecipientDictionary("To")
        End If
        If vRecipientDictionary.Exists("CC") Then
            .CC = vRecipientDictionary("CC")
        End If
        If vRecipientDictionary.Exists("BCC") Then
            .BCC = vRecipientDictionary("BCC")
        End If

        If Not (vAttachmentFilePaths Is Nothing) Then
            For Each vAttachmentFilePath In vAttachmentFilePaths
                If Not MUtilityFile.Exists(CStr(vAttachmentFilePath)) Then
                    Call MRuntime.RaiseError("MUtilityMail.Send", "The file to be attached cannot be found.")
                End If

                Call .Attachments.Add(CStr(vAttachmentFilePath))
            Next
        End If

        ' Set the message content.
        .Subject = vSubject
        If vIsHtmlBody Then
            .HTMLBody = vBody
        Else
            .Body = vBody
        End If

        ' Send the prepared message.
        If vIsBackgroundSendingEnabled Then
            Call .Send
        Else
            Call .Display(True)
        End If
    End With

Terminate:
    ' Reset error handling.
    On Error GoTo 0

    ' Log off the session if logged on successfully.
    If vIsLoggedOn Then
        Call vOutlookApplication.Session.Logoff
    End If

    ' Re-raise the error if needed.
    Call MRuntime.ReRaiseError

    ' Exit the procedure.
    Exit Sub

HandleError:
    ' Store the error for further handling.
    Call MRuntime.StoreError

    ' Resume to procedure termination.
    Resume Terminate:
End Sub
