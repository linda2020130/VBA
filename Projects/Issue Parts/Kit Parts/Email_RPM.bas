Sub Generate_RPM_Email(firstname, vendor, att, htm As String, i As Integer)
' 開啟Outlook, 自動擬RPM Issue Part草稿

    Set OutApp = CreateObject("Outlook.Application")
    Set Outmail = OutApp.CreateItem(0)
    Dim table, body, time As String
    table = GetHtm(htm)
    time = Sheets("Inv. Balance").Range("O1").Value
    body = HtmlBodyTemplateForRPM(firstname, vendor, time) & HtmlBodyDefinition() & table & "<br><br><br>" & GetSignature()
    With Outmail
        .To = Sheets("RPM List").Range("C" & i).Value
        .cc = Sheets("RPM List").Range("D" & i).Value
        .BCC = ""
        .Subject = Format(Date, "YYYYMMDD") & " Backlog and Shortage Issue RPM - " & vendor
        .Attachments.Add att
        .htmlbody = body
        .Save
    End With

End Sub

Function HtmlBodyTemplateForRPM(firstname, vendor, time As String)
'內文首段模板for RPM

    Dim html As String
    
    html = "<!DOCTYPE html><html><body>" & _
            "<font style=" & Chr(34) & "font-family:Calibri; font-size: 11pt;" & Chr(34) & ">" & _
            "Dear " & firstname & ", <br /><br />Attached is the <B><U>working parts summary</U></B>" & _
            " and <B><U>issue part list</U></B> of <B>" & vendor & "</B> for <B>RPM View</B>, based on data from <B><U>EIC-Inventory Balance Table</U></B>" & _
            " on " & Format(Date, "YYYY/MM/DD") & " " & time & "<br />" & _
            "Please remind your PMs to take actions (push out backlog, update forecast, ..., etc.) if necessary.<br />" & _
            "For more details, please refer to <B><U>MRP</U></B>.<br /><br /></font></body></html>"
    
    HtmlBodyTemplateForRPM = html
    
End Function