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
    
    
    Set objSourceFolder = objNameSpace.Folders("j.vansoest@maastrichtuniversity.nl")
    found = MyFolderSearch(objSourceFolder, msgId)
    
    If Not found Then
        Set objSourceFolder = objNameSpace.Folders("johan.vansoest@medicaldataworks.nl")
        found = MyFolderSearch(objSourceFolder, msgId)
    End If
    
    If Not found Then
        MsgBox ("Could not find the requested email")
    End If
    
End Sub

Function MyFolderSearch(folderToSearch As Outlook.Folder, msgId As String) As Boolean

    Dim subFolder As Outlook.Folder
    Dim found As Boolean
    
    'check for result in given folder
    found = SearchMessageIdInFolder(folderToSearch, msgId)
    If found Then
        MyFolderSearch = found
        Exit Function
    End If
    
    For Each subFolder In folderToSearch.Folders
        Debug.Print ("Folder name: " + subFolder.FolderPath)
        found = MyFolderSearch(subFolder, msgId)
        
        If found Then
            Exit For
        End If
    Next
    
    MyFolderSearch = found
End Function

Function SearchMessageIdInFolder(objSourceFolder As Outlook.Folder, msgId As String) As Boolean
    SearchMessageIdInFolder = False
    Dim strHeader As String
    Dim messageId As String
    Dim regEx As Object
    Dim M1 As Object
    Dim M As Object
    
    Dim itemList As Outlook.Items
    
    Set itemList = objSourceFolder.Items

    'Ignore if sorting is not possible
    On Error Resume Next
    itemList.Sort ("ReceivedTime")
    
    For Each Item In itemList
        If TypeOf Item Is Outlook.MailItem Then
            Dim oMail As Outlook.MailItem: Set oMail = Item
            
            strHeader = GetInetHeaders(oMail)
            strHeader = Replace(strHeader, vbCrLf, "")

            Set regEx = CreateObject("VBScript.RegExp")
            With regEx
              .Pattern = "(Message-ID:\s(.*?)>)"
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

