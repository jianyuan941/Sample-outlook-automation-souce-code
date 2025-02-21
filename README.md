# Sample-outlook-automation-souce-code
This VBA code allow to access specific subfolder in outlook, read then extract attachment and store in pc with specific folder name

=====================================================================================================================================
Sub listfolder()
Dim olApp As Outlook.Application
Dim olNs As Outlook.NameSpace
Dim olAccount As Outlook.Account
Dim olFolder As Outlook.Folder

Set olApp = New Outlook.Application
Set olNs = olApp.GetNamespace("MAPI")

For Each olAccount In olNs.Accounts
If olAccount.SmtpAddress = "<primary email account>" Then
Set olFolder = olNs.Folders(olAccount.DisplayName).Folders("Inbox")
    ListSubFolders olFolder
    Exit For
End If
Next olAccount
Set olFolder = Nothing
Set olAccount = Nothing
Set olNs = Nothing
Set olApp = Nothing

End Sub

Sub ListSubFolders(ByVal parentFolder As Outlook.Folder)
Dim subFolder As Outlook.Folder
Dim folderPath As String
Dim olMail As Outlook.MailItem
Dim latestMail As Outlook.MailItem
Dim hasAttachment As Boolean
Dim subjectwords() As String
Dim subjectPreview As String
Dim subjectPreview2 As String
Dim spacePosition As Variant
Dim firstword As Variant
Dim remainingtext As Variant
Dim remainingEmailsCount As Long

hasAttachment = False
Set latestMail = Nothing

For Each subFolder In parentFolder.Folders
    If subFolder.Name = "NEW CASE" Then
        For Each olMail In subFolder.Items
            If TypeOf olMail Is Outlook.MailItem Then
                If latestMail Is Nothing Then
                    Set latestMail = olMail
                ElseIf olMail.ReceivedTime > latestMail.ReceivedTime Then
                    Set latestMail = olMail
                End If
            End If
        Next olMail
        If Not latestMail Is Nothing Then
            If latestMail.Attachments.Count > 0 Then
             hasAttachment = True
             'MsgBox "yes has attachment"
            End If
            
            'MsgBox Trim(Mid(latestMail.Subject, spacePosition + 5))
            
            
            'subjectwords = Split(latestMail.Subject, " ")
            
            'If UBound(subjectwords) >= 1 Then
            '    ReDim Preserve subjectwords(0 To 1)
            '    subjectPreview = Join(subjectwords, " ")
            '    subjectPreview2 = latestMail.Subject
            '    spacePosition = InStr(subjectPreview2, "")
            '    If spacePosition > 0 Then
            '        remainingtext = Trim(Mid(subjectPreview2, spacePosition + 4))
            '    End If

            'Else
            '    subjectPreview = latestMail.Subject
            'End If
            '    MsgBox remainingtext
                
            If hasAttachment Then
                Dim strParentFolder As String
                Dim fullfolderpath As String
                Dim saveattachment As String
                Dim attachment As Outlook.attachment
                
                strParentFolder = "D:\JY\wip"
                fullfolderpath = strParentFolder & "\" & Trim(Mid(latestMail.Subject, spacePosition + 5)) & "\"
                'MsgBox fullfolderpath
                
                If Len(Dir(fullfolderpath, vbDirectory)) > 0 Then
                Else
                    MkDir fullfolderpath
                End If
                
                For Each attachment In latestMail.Attachments
                    saveattachment = fullfolderpath & attachment.FileName
                    attachment.SaveAsFile saveattachment
                Next attachment
                latestMail.Delete
                
                remainingEmailsCount = subFolder.Items.Count
                If remainingEmailsCount <> 0 Then
                    Call listfolder
                End If
            Else
            MsgBox "There is not attachment"
            End If
        Else
        MsgBox "No emails found in the inbox"
        End If
        Exit Sub
    End If
    ListSubFolders subFolder
Next subFolder
End Sub

