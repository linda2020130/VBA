Sub Copy_Sheets()

' Copy sheets("Summary", "table1", "table2") from Test.xlsm to new workwork and save as "Summary_YYYYMMDD.xlsx"

    ' Generate new workbook
    Workbooks.Add
    
    ' Save file with certain name
    Dim mypath, myfile As String
    mypath = "D:\Users\lindac\Desktop\" & Format(Date, "YYYYMM") & "\" & Format(Date, "MMDD") & "\"
    myfile = "Summary_" & Format(Date, "YYYYMMDD") & ".xlsx"
    ActiveWorkbook.SaveAs Filename:=mypath & myfile, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    ' Copy sheets from original workbook
    Windows("Test.xlsm").Activate
    ' Put copied sheet after the first sheet in the new workbook
    Sheets(Array("Summary", "Table1", "Table2")).Copy After:=Workbooks(myfile).Sheets(1)
    Application.DisplayAlerts = False
    ' Delete worksheets
    Sheets(Array("Sheet1", "Sheet2", "Sheet3")).Delete

    ' Save and close new workbook
    Workbooks(myfile).Close SaveChanges:=True

End Sub