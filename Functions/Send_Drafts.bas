Sub SendAllYourMailboxDrafts()

' Outlook VBA for sending a batch of emails

     SendAllDrafts "linda2020130"
     
End Sub


Sub SendAllDrafts(mailbox As String)

' Send out all emails in the draft
' Put mailbox name as input 

     Dim folder As MAPIFolder
     Dim msg As Outlook.MailItem
     Dim count As Integer

     Set folder = Outlook.GetNamespace("MAPI").Folders(mailbox)
     Set folder = folder.Folders("Drafts")

     If MsgBox("Are you sure to send out " & folder.Items.count & " emails in " & mailbox & " ?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub

     count = 0
     Do While folder.Items.count > 0
      Set msg = folder.Items(1)
      msg.Send
      count = count + 1
     Loop

     MsgBox count & " emails have been sent.", vbInformation + vbOKOnly
     
End Sub
