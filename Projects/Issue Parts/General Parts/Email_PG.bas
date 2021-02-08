Sub Generate_PGHead_Email(name, att, htm As String, i As Integer)
' 開啟Outlook, 自動擬PG Issue Part草稿

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    Dim table, body, time As String
    table = GetHtm(htm)
    time = Sheets("Inv. Balance").Range("N1").Value
    
    body = HtmlBodyTemplateForPGHead(name, time) & HtmlBodyDefinition() & table & "<br><br><br>" & GetSignature()
    
    With OutMail
        .To = Sheets("Group List").Range("D" & i).Value
        .cc = Sheets("Group List").Range("E" & i).Value
        .BCC = ""
        .Subject = Format(Date, "YYYYMMDD") & " Backlog and Shortage Issue"
        Dim files As Variant, file As Variant
        filepath = Sheets("Group List").Range("G" & i).Value
        If filepath <> "" Then
            files = Split(filepath, ",")
            For Each file In files
                .attachments.Add file
            Next
        End If
        .attachments.Add att
        .htmlbody = body
        '.display
        .Save
        '.Send
        
    End With


End Sub
Function HtmlBodyTemplateForPGHead(name, time As String)
'內文首段模板for PG Head

    Dim html As String
    
    
    html = "<!DOCTYPE html><html><body>" & _
            "<font style=" & Chr(34) & "font-family:Calibri; font-size: 11pt;" & Chr(34) & ">" & _
            "Dear " & name & ", <br /><br />Attached is your team's <B><U>working parts summary</U></B>" & _
            " and <B><U>issue part list</U></B>, based on data from <B><U>EIC-Inventory Balance Table</U></B>" & _
            " on " & Format(Date, "YYYY/MM/DD") & " " & time & "<br />" & _
            "Please remind your PMs to take actions (push out backlog, update forecast, ..., etc.) if necessary.<br />" & _
            "For more details, please refer to <B><U>MRP</U></B>.<br /><br /></font></body></html>"
            
    
    HtmlBodyTemplateForPGHead = html
    
End Function