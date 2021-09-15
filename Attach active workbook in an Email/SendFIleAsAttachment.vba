Sub SendFIleAsAttachment()
    'Declare your variables
    'Set reference to Microsoft Outlook Object library
        Dim OLApp As Outlook.Application
        Dim OLMail As Object
    'Open Outlook start a new mail item
        Set OLApp = New Outlook.Application
        Set OLMail = OLApp.CreateItem()
        OLApp.Session.Logon  
    'Build your mail item and send
        With OLMail
        .To = "admin@datapigtechnologies.com; mike@datapigtechnologies.com"
        .CC = ""
        .BCC = ""
        .Subject = "This is the Subject line"
        .Body = "Hi there"
        .Attachments.Add ActiveWorkbook.FullName
        .Display  'Change to .Send to send without reviewing
        End With
    'Memory cleanup
        Set OLMail = Nothing
        Set OLApp = Nothing
End Sub