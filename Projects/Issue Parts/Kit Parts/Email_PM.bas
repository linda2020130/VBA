Sub Generate_PM_Email(name, att, htm As String, i As Integer)
' 開啟Outlook, 自動擬PM Issue Part草稿

    Set OutApp = CreateObject("Outlook.Application")
    Set Outmail = OutApp.CreateItem(0)
    Dim table, body, time As String
    table = GetHtm(htm)
    time = Sheets("Inv. Balance").Range("O1").Value
    body = HtmlBodyTemplateForPM(name, time) & HtmlBodyDefinition() & table & "<br><br><br>" & GetSignature()
    With Outmail
        .To = Sheets("PM List").Range("E" & i).Value
        .cc = Sheets("PM List").Range("F" & i).Value
        .BCC = ""
        .Subject = Format(Date, "YYYYMMDD") & " Backlog and Shortage Issue"
        .Attachments.Add att
        .htmlbody = body
        .Save
    End With

End Sub

Function HtmlBodyTemplateForPM(name, time As String)
'內文首段模板for PM

    Dim html As String
    
    html = "<!DOCTYPE html><html><body>" & _
            "<font style=" & Chr(34) & "font-family:Calibri; font-size: 11pt;" & Chr(34) & ">" & _
            "Dear " & name & ", <br /><br />Attached is your <B><U>working parts summary</U></B>" & _
            " and <B><U>issue part list</U></B>, based on data from <B><U>EIC-Inventory Balance Table</U></B>" & _
            " on " & Format(Date, "YYYY/MM/DD") & " " & time & "<br />" & _
            "<span style=" & Chr(34) & "background-color:#FFFF00" & Chr(34) & ">Please go to <B>MRP</B> and check <B>one by one</B></span>" & _
            " and adjust your latest forecast, backlog, ..., etc. if necessary.<br /><br /></font></body></html>"
    
    HtmlBodyTemplateForPM = html
    
End Function