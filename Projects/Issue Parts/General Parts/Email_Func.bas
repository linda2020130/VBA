Function HtmlBodyDefinition()
'內文Issue Part定義與結語

    Dim html1, html2 As String
    html1 = "<!DOCTYPE html><html><body>" & _
            "<font style=" & Chr(34) & "font-family:Calibri; font-size: 10.0pt;" & Chr(34) & ">" & _
            "Definition:<br /><ol><li> Working Part: number of all parts with inventory or backlog or forecast or billing.<br />" & _
            "<li> Stock Ratio Issue: number of all working parts with stock ratio > 1.5.<br />" & _
            "<li> <B>Backlog Issue</B>: number of all working parts with <U>stock ratio > 1.5 for two consecutive months <B>AND</B> still have backlog</U>.<br />" & _
            "<li> <B>Shortage Issue</B>: number of all working parts with possible shortage at the end of next month.<br /><br /></font></body></html>"
    
    html2 = "<!DOCTYPE html><html><body>" & _
            "<font style=" & Chr(34) & "font-family:Calibri; font-size: 11pt;" & Chr(34) & ">" & _
            "If you have any question, please contact me or Jerry Chen (ext. 2300).<br /><br /></font></body></html>"
    
    HtmlBodyDefinition = html1 & html2

End Function

Function GetHtm(htm_path As String)
'讀取htm檔

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f_htmObj = fso.OpenTextFile(htm_path, 1, False, 0)
    GetHtm = f_htmObj.ReadAll
    f_htmObj.Close
    Set fso = Nothing

End Function

Function GetSignature()
' 讀取Outlook Email的簽名檔

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    SigPath = "C:\Users\lindac\AppData\Roaming\Microsoft\Signatures\Linda Chou.htm"
    Set f_SignatureObj = fso.OpenTextFile(SigPath, 1, False, 0)
    GetSignature = f_SignatureObj.ReadAll
    f_SignatureObj.Close
    Set fso = Nothing

End Function
