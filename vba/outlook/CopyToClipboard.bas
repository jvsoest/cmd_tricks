Attribute VB_Name = "Module1"
'Adds a link to the currently selected message to the clipboard
Sub ArchiveCopySelectedMsgInfo()

Dim objOutlook As Outlook.Application
Dim objNameSpace As Outlook.NameSpace
Dim objSourceFolder As Outlook.MAPIFolder
Dim obDestFolder As Outlook.MAPIFolder
Dim objMail As Outlook.MailItem
Dim senderString As String
Dim subjectString As String
Dim dateString As String
Dim outlookMessageId As String

Dim strHeader As String
Dim messageId As String
Dim regEx As Object
Dim M1 As Object
Dim M As Object

mailboxNameString = "j.vansoest@maastrichtuniversity.nl" 'Enter your email address here!

Set objOutlook = Application
Set objNameSpace = objOutlook.GetNamespace("MAPI")
Set objSourceFolder = Application.ActiveExplorer.CurrentFolder
'Set objSourceFolder = objNameSpace.GetDefaultFolder(olFolderInbox) ' From the inbox

'Set obDestFolder = objNameSpace.Folders(mailboxNameString).Folders("Archive") ' To archive folder

' Catch when multiple emails were selected
If Application.ActiveExplorer.Selection.Count <> 1 Then
    MsgBox ("Multiple emails were selected, please select only 1.")
    Exit Sub
End If

Set objMail = Application.ActiveExplorer.Selection.Item(1) ' Select the email
'Set objMail = Application.ActiveExplorer.Selection.Item(1).Move(obDestFolder) ' Move the email
objMail.UnRead = False ' Mark as read

strHeader = GetInetHeaders(objMail)
strHeader = Replace(strHeader, vbCrLf, "")

Set regEx = CreateObject("VBScript.RegExp")
With regEx
  .Pattern = "(Message-ID:\s(.*?)>)"
  .Global = True
End With

If regEx.test(strHeader) Then
    Set M1 = regEx.Execute(strHeader)
    For Each M In M1
        messageId = M
    Next
End If

senderString = "From: " + objMail.SenderName + " (" + objMail.SenderEmailAddress + ")"
subjectString = "Subject: " + objMail.Subject
dateString = "Date: " + Format(objMail.ReceivedTime, "dd-mm-yyyy hh:mm")

messageId = Replace(messageId, "Message-ID: ", "")
messageId = Replace(messageId, "<", "%3C")
messageId = Replace(messageId, ">", "%3E")
messageId = "message://" + messageId

With New HTMLDocument
    SetClipBoardText = .parentWindow.ClipboardData.SetData("Text", senderString + vbNewLine + dateString + vbNewLine + subjectString + vbNewLine + messageId)
End With

'Dim cmdString As String
'cmdString = "addTask.bat -ui -cb """ + Replace(objMail.Subject, """", "") + """"
'Shell (cmdString)

MsgBox ("Email metadata sent to clipboard")

End Sub

Function GetInetHeaders(olkMsg As Outlook.MailItem) As String
    ' Purpose: Returns the internet headers of a message.'
    ' Written: 4/28/2009'
    ' Author:  BlueDevilFan'
    ' //techniclee.wordpress.com/
    ' Outlook: 2007'
    Const PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
    Dim olkPA As Outlook.PropertyAccessor
    Set olkPA = olkMsg.PropertyAccessor
    GetInetHeaders = olkPA.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS)
    Set olkPA = Nothing
End Function
