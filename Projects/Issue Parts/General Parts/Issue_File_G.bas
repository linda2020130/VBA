Sub Generate_PM_Issue_File()
' 產生各PM的Issue Part List, 並儲存在所屬群別裡

    ' 開啟Timer
    t0 = Timer
    
    ' 產生用來跑迴圈的PM Table
    Application.ScreenUpdating = False
    Dim row_pm As Integer
    row_pm = Sheets("PM List").Cells(rows.Count, "A").End(xlUp).row
    Dim group_long, pm, location, group_short, pm_first As String
    Dim i As Integer
    For i = 2 To row_pm
        With Worksheets("PM List")
            group_long = .Range("A" & i).Value
            pm = .Range("B" & i).Value
            location = .Range("C" & i).Value
            group_short = .Range("D" & i).Value
            pm_first = .Range("G" & i).Value
        End With
        ' Raw Data裡篩選群別(PM), PM, 出貨倉
        With Worksheets("Raw Data")
            .AutoFilterMode = False
            Dim row_rdata As Integer
            row_rdata = .Cells(rows.Count, "A").End(xlUp).row
            .Range("A2:CD" & row_rdata).AutoFilter Field:=2, Criteria1:=group_long
            .Range("A2:CD" & row_rdata).AutoFilter Field:=5, Criteria1:=pm
            .Range("A2:CD" & row_rdata).AutoFilter Field:=7, Criteria1:=location
        End With
        
        ' 呼叫副程匯入資料
        Import_Data
        
        ' 更新Backlog Issue和Shortage Issue工作表的月份
        Update_Month
        
        ' 刪除non-working part的資料
        With Worksheets("Inv. Balance")
            Dim y, row_inv As Integer
            row_inv = .Cells(rows.Count, "A").End(xlUp).row
            For y = row_inv To 6 Step -1
                If .Range("A" & y) = "0" Then
                    .rows(y).EntireRow.Delete
                End If
            Next
            row_inv = .Cells(rows.Count, "A").End(xlUp).row
        End With
        
        If row_inv > 5 Then
            ' 呼叫副程序產生Summary
            Generate_Summary
            
            ' htm檔名, 當作副程序引數
            Dim htmfile As String
            htmfile = "D:\Users\lindac\documents\Issue Part\" & Format(Date, "YYYYMM") & "\" & _
                Format(Date, "MMDD") & "\" & group_short & "\" & "Summary_" & Format(Date, "YYYYMMDD") & _
                "_" & group_short & "_" & pm & ".htm"
                
            ' 呼叫副程序將Summary儲存成htm檔(供後續Outlook郵件夾帶進內文)
            Save_htm_File (htmfile)
            
            With Worksheets("Inv. Balance")
                ' Inv. Balance裡依金額排序
                '.Range("A5:CS" & row_inv).Sort key1:=.Range("P5"), order1:=xlDescending, Header:=xlYes
                ' 篩選出Backlog Issue Part
                .Range("A5:CS" & row_inv).AutoFilter Field:=3, Criteria1:="1"
                Dim row_b_issue As Integer
                row_b_issue = .Cells(rows.Count, "A").End(xlUp).row
                ' 若有Backlog Issue的料, 複製到Backlog Issue工作表裡
                If row_b_issue > 5 Then
                    .Range("A6:CS" & row_b_issue).SpecialCells(xlCellTypeVisible).Copy
                    With Worksheets("Backlog Issue")
                        .Range("A4").PasteSpecial xlPasteValues
                        .Range("A1:Y1").Copy
                        .Range("A4:Y" & .Range("A3").End(xlDown).row).PasteSpecial xlPasteFormats
                        .Range("B2").Copy
                        .Range("B2").PasteSpecial xlPasteValues
                        Application.CutCopyMode = False
                    End With
                End If
                ' 篩選Shortage Issue Part
                .Range("A5:CS" & row_inv).AutoFilter Field:=3
                .Range("A5:CS" & row_inv).AutoFilter Field:=5, Criteria1:="1"
                Dim row_s_issue As Integer
                row_s_issue = .Cells(rows.Count, "A").End(xlUp).row
                ' 若有Shortage Issue的料, 複製到Shortage Issue工作表裡
                If row_s_issue > 5 Then
                    .Range("A6:CS" & row_s_issue).SpecialCells(xlCellTypeVisible).Copy
                    With Worksheets("Shortage Issue")
                        .Range("A4").PasteSpecial xlPasteValues
                        .Range("A1:Y1").Copy
                        .Range("A4:Y" & .Range("A3").End(xlDown).row).PasteSpecial xlPasteFormats
                        .Range("C2").Copy
                        .Range("C2").PasteSpecial xlPasteValues
                        Application.CutCopyMode = False
                    End With
                End If
            End With
            
            ' 僅複製有Issue List的工作表
            If row_b_issue > 5 And row_s_issue > 5 Then
                ' 新增活頁簿
                Workbooks.Add
                ' 將新活頁簿依指定名稱儲存
                Dim mypath, myfile As String
                mypath = "D:\Users\lindac\Documents\Issue Part\" & Format(Date, "YYYYMM") _
                    & "\" & Format(Date, "MMDD") & "\" & group_short & "\"
                myfile = "Issue Part List_" & Format(Date, "YYYYMMDD") & "_" _
                    & group_short & "_" & pm & ".xlsx"
                ActiveWorkbook.SaveAs Filename:=mypath & myfile, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                ' 複製Backlog Issue和Shortage Issue工作表到新活頁簿
                Windows("Issue Parts_template_for LC.xlsm").Activate
                Sheets(Array("Summary", "Backlog Issue", "Shortage Issue")).Copy After:=Workbooks(myfile).Sheets(1)
                Application.DisplayAlerts = False
                Sheets(Array("工作表1")).Delete
                ' 刪除Backlog和Shortage Issue裡的第一列
                Sheets("Shortage Issue").rows("1:1").Delete
                Sheets("Backlog Issue").rows("1:1").Delete
                ' 儲存新活頁簿後關閉
                Workbooks(myfile).Close SaveChanges:=True
                
                ' 呼叫副程序自動擬Outlook郵件
                Dim attfile As String
                attfile = mypath & myfile
                Generate_PM_Email pm_first, attfile, htmfile, i
                
                Windows("Issue Parts_template_for LC.xlsm").Activate
                ' 刪除原Backlog Issue裡的資料
                With Worksheets("Backlog Issue")
                    .rows("4:" & .Range("A4").End(xlDown).row).Delete
                End With
                ' 刪除原Shortage Issue裡的資料
                With Worksheets("Shortage Issue")
                    .rows("4:" & .Range("A4").End(xlDown).row).Delete
                End With
            
            ElseIf row_b_issue > 5 Then
                ' 新增活頁簿
                Workbooks.Add
                ' 將新活頁簿依指定名稱儲存
                mypath = "D:\Users\lindac\Documents\Issue Part\" & Format(Date, "YYYYMM") _
                    & "\" & Format(Date, "MMDD") & "\" & group_short & "\"
                myfile = "Issue Part List_" & Format(Date, "YYYYMMDD") & "_" _
                    & group_short & "_" & pm & ".xlsx"
                ActiveWorkbook.SaveAs Filename:=mypath & myfile, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                ' 複製Backlog Issue和Shortage Issue工作表
                Windows("Issue Parts_template_for LC.xlsm").Activate
                Sheets(Array("Summary", "Backlog Issue")).Copy After:=Workbooks(myfile).Sheets(1)
                Application.DisplayAlerts = False
                Sheets(Array("工作表1")).Delete
                ' 刪除backlog Issue裡的第一列
                Sheets("Backlog Issue").rows("1:1").Delete
                ' 儲存新活頁簿後關閉
                Workbooks(myfile).Close SaveChanges:=True
                
                ' 呼叫副程序自動擬Outlook郵件
                attfile = mypath & myfile
                Generate_PM_Email pm_first, attfile, htmfile, i
                
                Windows("Issue Parts_template_for LC.xlsm").Activate
                ' 刪除原Backlog Issue裡的資料
                With Worksheets("Backlog Issue")
                    .rows("4:" & .Range("A4").End(xlDown).row).Delete
                End With
            
            ElseIf row_s_issue > 5 Then
                ' 新增活頁簿
                Workbooks.Add
                ' 將新活頁簿依指定名稱儲存
                mypath = "D:\Users\lindac\Documents\Issue Part\" & Format(Date, "YYYYMM") _
                    & "\" & Format(Date, "MMDD") & "\" & group_short & "\"
                myfile = "Issue Part List_" & Format(Date, "YYYYMMDD") & "_" _
                    & group_short & "_" & pm & ".xlsx"
                ActiveWorkbook.SaveAs Filename:=mypath & myfile, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                ' 複製Backlog Issue和Shortage Issue工作表
                Windows("Issue Parts_template_for LC.xlsm").Activate
                Sheets(Array("Summary", "Shortage Issue")).Copy After:=Workbooks(myfile).Sheets(1)
                Application.DisplayAlerts = False
                Sheets(Array("工作表1")).Delete
                ' 刪除shortage Issue裡的第一列
                Sheets("Shortage Issue").rows("1:1").Delete
                ' 儲存新活頁簿後關閉
                Workbooks(myfile).Close SaveChanges:=True
                
                ' 呼叫副程序自動擬Outlook郵件
                attfile = mypath & myfile
                Generate_PM_Email pm_first, attfile, htmfile, i
                
                Windows("Issue Parts_template_for LC.xlsm").Activate
                ' 刪除原Shortage Issue裡的資料
                With Worksheets("Shortage Issue")
                    .rows("4:" & .Range("A4").End(xlDown).row).Delete
                End With
            End If
            
            ' 復原Inv. Balance裡的篩選
            With Worksheets("Inv. Balance")
                .Range("A5:CS" & row_inv).AutoFilter Field:=5
                .Range("I1").Select
            End With
        End If
    Next
    
    Application.ScreenUpdating = True
    
    '結算總運行時間
    t0 = Timer - t0
    MsgBox "完成各PM的Issue File" & vbCrLf & t0 \ 60 & "分 " & t0 Mod 60 & "秒"
    
End Sub

Sub Generate_Group_Issue_File()
' 產生各群的Issue Part List, 並儲存在所屬群別裡

    ' 開啟Timer
    t0 = Timer
    
    ' 產生用來跑迴圈的PG Table
    Application.ScreenUpdating = False
    Dim n As Integer
    n = Sheets("Group List").Cells(rows.Count, "A").End(xlUp).row
    Dim group_long, location, group_short, pghead_first As String
    Dim i As Integer
    For i = 2 To n
        With Worksheets("Group List")
            group_long = .Range("A" & i).Value
            location = .Range("B" & i).Value
            group_short = .Range("C" & i).Value
            pghead_first = .Range("F" & i).Value
        End With
        ' Raw Data裡篩選群別(PM), 出貨倉
        With Worksheets("Raw Data")
            .AutoFilterMode = False
            Dim k As Integer
            k = .Cells(rows.Count, "A").End(xlUp).row
            .Range("A2:AW" & k).AutoFilter Field:=2, Criteria1:=group_long
            .Range("A2:AW" & k).AutoFilter Field:=7, Criteria1:=location
        End With
        
        ' 呼叫副程序匯入資料
        Import_Data
        
        ' 更新Backlog Issue和Shortage Issue工作表的月份
        Update_Month
        
        ' 刪除non-working part的資料
        With Worksheets("Inv. Balance")
            Dim f As Integer
            f = .Cells(rows.Count, "A").End(xlUp).row
            For y = f To 6 Step -1
                If .Range("A" & y) = "0" Then
                    .rows(y).EntireRow.Delete
                End If
            Next
        End With

        ' 呼叫副程序產生Summary
        Generate_Summary
        
        ' htm檔名, 當作副程序引數
        Dim htmfile As String
        htmfile = "D:\Users\lindac\documents\Issue Part\" & Format(Date, "YYYYMM") & "\" & _
            Format(Date, "MMDD") & "\" & group_short & "\" & "Summary_" & Format(Date, "YYYYMMDD") & _
            "_" & group_short & ".htm"
            
        ' 呼叫副程序將Summary儲存成htm檔(供後續Outlook郵件夾帶進內文)
        Save_htm_File (htmfile)
        
        With Worksheets("Inv. Balance")
            ' Inv. Balance裡依金額排序後
            Dim row_inv As Integer
            row_inv = .Cells(rows.Count, "A").End(xlUp).row
            '.Range("A5:CS" & row_inv).Sort key1:=.Range("P5"), order1:=xlDescending, Header:=xlYes
            ' 篩選出Backlog Issue Part
            .Range("A5:CS" & row_inv).AutoFilter Field:=3, Criteria1:="1"
            ' 複製到backlog Issue工作表裡
            .Range("A6:CS" & .Range("A6").End(xlDown).row).SpecialCells(xlCellTypeVisible).Copy
            With Worksheets("Backlog Issue")
                .Range("A4").PasteSpecial xlPasteValues
                .Range("A1:Y1").Copy
                .Range("A4:Y" & .Range("A3").End(xlDown).row).PasteSpecial xlPasteFormats
                .Range("B2").Copy
                .Range("B2").PasteSpecial xlPasteValues
                Application.CutCopyMode = False
            End With
            ' 篩選出Shortage Issue Part
            .Range("A5:CS" & row_inv).AutoFilter Field:=3
            .Range("A5:CS" & row_inv).AutoFilter Field:=5, Criteria1:="1"
            ' 複製到Shortage Issue工作表裡
            .Range("A6:CS" & .Range("A6").End(xlDown).row).SpecialCells(xlCellTypeVisible).Copy
            With Worksheets("Shortage Issue")
                .Range("A4").PasteSpecial xlPasteValues
                .Range("A1:Y1").Copy
                .Range("A4:Y" & .Range("A3").End(xlDown).row).PasteSpecial xlPasteFormats
                .Range("C2").Copy
                .Range("C2").PasteSpecial xlPasteValues
                Application.CutCopyMode = False
            End With
        End With
        
        ' 新增活頁簿
        Workbooks.Add
        ' 將新活頁簿依指定名稱儲存
        Dim mypath, myfile As String
        mypath = "D:\Users\lindac\Documents\Issue Part\" & Format(Date, "YYYYMM") _
            & "\" & Format(Date, "MMDD") & "\" & group_short & "\"
        myfile = "Issue Part List_" & Format(Date, "YYYYMMDD") & "_" _
            & group_short & ".xlsx"
        ActiveWorkbook.SaveAs Filename:=mypath & myfile, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        ' 複製Backlog Issue, Shortage Issue和Summary工作表到新檔案
        Windows("Issue Parts_template_for LC.xlsm").Activate
        Sheets(Array("Summary", "Backlog Issue", "Shortage Issue")).Copy After:=Workbooks(myfile).Sheets(1)
        Application.DisplayAlerts = False
        Sheets(Array("工作表1")).Delete
        ' 刪除Backlog和Shortage Issue裡的第一列
        Sheets("Shortage Issue").rows("1:1").Delete
        Sheets("Backlog Issue").rows("1:1").Delete
        ' 儲存新活頁簿後關閉
        Workbooks(myfile).Close SaveChanges:=True
        
        ' 呼叫副程序自動擬Outlook郵件
        Dim attfile As String
        attfile = mypath & myfile
        Generate_PGHead_Email pghead_first, attfile, htmfile, i
        
        Windows("Issue Parts_template_for LC.xlsm").Activate
        ' 刪除原Backlog Issue裡的資料
        With Worksheets("Backlog Issue")
            .rows("4:" & .Range("A4").End(xlDown).row).Delete
        End With
        ' 刪除原Shortage Issue裡的資料
        With Worksheets("Shortage Issue")
            .rows("4:" & .Range("A4").End(xlDown).row).Delete
        End With
        
        ' 復原Inv. Balance裡的篩選
        With Worksheets("Inv. Balance")
            .Range("A5:CS" & row_inv).AutoFilter Field:=5
            .Range("I1").Select
        End With
    Next
    
    Application.ScreenUpdating = True
    
    '結算總運行時間
    t0 = Timer - t0
    MsgBox "完成各群的Issue File" & vbCrLf & t0 \ 60 & "分 " & t0 Mod 60 & "秒"
    
End Sub