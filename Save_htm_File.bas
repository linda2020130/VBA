Sub Save_htm_File()

' Save table in ActiveWorkbook as html file
' htm file can be inserted to email content

    Dim mypath, myfile As String
    mypath = "D:\Users\lindac\Desktop\" & Format(Date, "YYYYMM") & "\" & Format(Date, "MMDD") & "\"
    myfile = "Summary Table_" & Format(Date, "YYYYMMDD") & ".htm"
    Dim table As Range
    Set table = Sheets("Summary").Range("A1:J" & Range("J1").End(xlDown).row)
    Dim htmfile As String
    htmfile = mypath & myfile
    
    With ActiveWorkbook.PublishObjects.Add(xlSourceRange, htmfile, "Summary", _
        table.Address, xlHtmlStatic)
        .Publish (False)
        .AutoRepublish = False
    End With

End Sub