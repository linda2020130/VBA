Function GetSignature()

' Get outlook signature to be inserted to email content

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    SigPath = "C:\Users\lindac\AppData\Roaming\Microsoft\Signatures\Linda Chou.htm"
    Set f_SignatureObj = fso.OpenTextFile(SigPath, 1, False, 0)
    GetSignature = f_SignatureObj.ReadAll
    f_SignatureObj.Close
    Set fso = Nothing

End Function

Function GetHtm(htm_path As String)

' Get htm file to be inserted to email content
' Put htm path including file name as input for GetHtm function

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f_htmObj = fso.OpenTextFile(htm_path, 1, False, 0)
    GetHtm = f_htmObj.ReadAll
    f_htmObj.Close
    Set fso = Nothing

End Function

Function HtmlBody(name, time As String)

' Get email body content
' Put receiver's name and data downloaded time as input for HtmlBody function

    Dim html1, html2 As String
    html1 = "<!DOCTYPE html><html><body>" & _
            "<font style=" & Chr(34) & "font-family:Calibri; font-size: 11pt;" & Chr(34) & ">" & _
            "Dear " & name & ", <br /><br />Attached is your <B><U>Weekly Summary</U></B>" & _
            ", based on data from system on " & Format(Date, "YYYY/MM/DD") & " " & time & "<br />" & _
            "<span style=" & Chr(34) & "background-color:#FFFF00" & Chr(34) & ">Please go on system for more details</span>" & _
            " and take any action if necessary.<br /><br /></font></body></html>"
    
    ' Create another html content for another paragraph if font-family, font_size...is different
    html2 = "<!DOCTYPE html><html><body>" & _
            "<font style=" & Chr(34) & "font-family:Arial; font-size: 10.0pt;" & Chr(34) & ">" & _
            "If you have any question, please feel free to contact me.<br /><br /></font></body></html>"
            
    ' Combine htmls together
    HtmlBodyTemplateForPM = html1 & html2
    
End Function


Sub Generate_Email(name, attchment, htm As String, i As Integer)

' Generate drafts in outlook according to "Receiver List"
' Attach file in the email (attachment = path + file)
' Insert htm table in the email content (htm = path + file)

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    Dim table, body, time As String
    table = GetHtm(htm)
    time = Sheets("Summary").Range("B1").Value
    body = HtmlBody(name, time) & table & "<br><br><br>" & GetSignature()
    
    With OutMail
        .To = Sheets("Receiver List").Range("E" & i).Value
        .cc = Sheets("Receiver List").Range("F" & i).Value
        .BCC = ""
        .Subject = Format(Date, "YYYYMMDD") & " Weekly Summary"
        .attachments.Add attchment
        .htmlbody = body
        '.display
        .Save
        '.Send
    End With

End Sub