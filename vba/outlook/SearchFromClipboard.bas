Attribute VB_Name = "Module2"
Public Sub SearchMessageId()
    Dim msgId As String
    Dim mailboxNameString As String
    Dim found As Boolean
    Dim objSourceFolder As Outlook.Folder
    
    msgId = InputBox("Enter Message-ID:")
    msgId = Replace(msgId, "message://", "")
    msgId = Replace(msgId, "%3C", "<")
    msgId = Replace(msgId, "%3E", ">")
    
    Set objOutlook = Application
    Set objNameSpace = objOutlook.GetNamespace("MAPI")
    
    
    mailboxNameString = "j.vansoest@maastrichtuniversity.nl" 'Enter your email address here!
    Set objSourceFolder = objNameSpace.Folders(mailboxNameString).Folders("Postvak IN")
    found = SearchMessageIdInFolder(objSourceFolder, msgId)
    
    If Not found Then
        Set objSourceFolder = objNameSpace.Folders(mailboxNameString).Folders("Archive")
        found = SearchMessageIdInFolder(objSourceFolder, msgId)
    End If
    
End Sub

Function SearchMessageIdInFolder(objSourceFolder As Outlook.Folder, msgId As String) As Boolean
    SearchMessageIdInFolder = False
    Dim strHeader As String
    Dim messageId As String
    Dim regEx As Object
    Dim M1 As Object
    Dim M As Object
    
    For Each Item In objSourceFolder.Items
        If TypeOf Item Is Outlook.MailItem Then
            Dim oMail As Outlook.MailItem: Set oMail = Item
            
            strHeader = GetInetHeaders(oMail)

            Set regEx = CreateObject("VBScript.RegExp")
            With regEx
              .Pattern = "(Message-ID:\s(.*))"
              .Global = True
            End With
            
            If regEx.test(strHeader) Then
                Set M1 = regEx.Execute(strHeader)
                For Each M In M1
                    messageId = Replace(M, "Message-ID: ", "")
                    messageId = Replace(messageId, vbLf, "")
                    messageId = Replace(messageId, vbCr, "")
                Next
            End If
            
            messageId = Trim(messageId)
            msgId = Trim(msgId)
            If messageId = msgId Then
                oMail.Display
                SearchMessageIdInFolder = True
                Exit Function
            End If
        End If
    Next
    
End Function

Function GetInetHeaders(olkMsg As Outlook.MailItem) As String
    ' Purpose: Returns the internet headers of a message.'
    ' Written: 4/28/2009'
    ' Author:  BlueDevilFan'
    ' //techniclee.wordpress.com/
    ' Outlook: 2007'
    On Error GoTo errhand
    Const PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
    Dim olkPA As Outlook.PropertyAccessor
    Set olkPA = olkMsg.PropertyAccessor
    GetInetHeaders = olkPA.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS)
    Set olkPA = Nothing
    Exit Function
errhand:
    GetInetHeaders = ""
End Function

