Sub Generate_PM_Issue_File()
' 產生各PM的Issue Part List, 並儲存在所屬群別裡

    ' 開啟Timer
    t0 = Timer
    
    ' 產生用來跑迴圈的PM Table
    Application.ScreenUpdating = False
    Dim row_pm As Integer
    row_pm = Sheets("PM List").Cells(rows.Count, "A").End(xlUp).row
    Dim group_long, pm, location, group_short, pm_first, vendor, email As String
    Dim i As Integer
    For i = 2 To row_pm
        With Worksheets("PM List")
            group_long = .Range("A" & i).Value
            pm = .Range("B" & i).Value
            location = .Range("C" & i).Value
            group_short = .Range("D" & i).Value
            pm_first = .Range("G" & i).Value
            vendor = .Range("H" & i).Value
            email = .Range("I" & i).Value
        End With
        ' Raw Data裡篩選群別(PM), PM, 出貨倉, 供應商
        With Worksheets("Raw Data")
            .AutoFilterMode = False
            Dim row_rdata As Integer
            row_rdata = Sheets("Raw Data").Cells(rows.Count, "A").End(xlUp).row
            .Range("A2:BC" & row_rdata).AutoFilter Field:=2, Criteria1:=group_long
            .Range("A2:BC" & row_rdata).AutoFilter Field:=3, Criteria1:=vendor
            .Range("A2:BC" & row_rdata).AutoFilter Field:=5, Criteria1:=pm
            .Range("A2:BC" & row_rdata).AutoFilter Field:=7, Criteria1:=location
        End With
        
        ' 呼叫副程序匯入資料
        Import_Data
        
        ' 更新Backlog Issue和Shortage Issue工作表的月份
        Update_Month
        
        Dim row_inv As Integer
        row_inv = Sheets("Inv. Balance").Cells(rows.Count, "O").End(xlUp).row
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
                '.Range("A5:CZ" & row_inv).Sort key1:=.Range("W5"), order1:=xlDescending, Header:=xlYes
                ' 篩選出Backlog Issue Part
                .Range("A5:CZ" & row_inv).AutoFilter Field:=3, Criteria1:="1"
                Dim row_b_issue As Integer
                row_b_issue = .Cells(rows.Count, "A").End(xlUp).row
                ' 若有Backlog Issue的料, 複製到Backlog Issue工作表裡
                If row_b_issue > 5 Then
                    .Range("A6:CZ" & row_b_issue).SpecialCells(xlCellTypeVisible).Copy
                    With Worksheets("Backlog Issue")
                        .Range("A4").PasteSpecial xlPasteValues
                        .Range("A1:Z1").Copy
                        .Range("A4:Z" & .Range("A3").End(xlDown).row).PasteSpecial xlPasteFormats
                        .Range("B2").Copy
                        .Range("B2").PasteSpecial xlPasteValues
                        Application.CutCopyMode = False
                    End With
                End If
                ' 篩選Shortage Issue Part
                .Range("A5:CZ" & row_inv).AutoFilter Field:=3
                .Range("A5:CZ" & row_inv).AutoFilter Field:=5, Criteria1:="1"
                Dim row_s_issue As Integer
                row_s_issue = .Cells(rows.Count, "A").End(xlUp).row
                ' 若有Shortage Issue的料, 複製到Shortage Issue工作表裡
                If row_s_issue > 5 Then
                    .Range("A6:CZ" & row_s_issue).SpecialCells(xlCellTypeVisible).Copy
                    With Worksheets("Shortage Issue")
                        .Range("A4").PasteSpecial xlPasteValues
                        .Range("A1:Z1").Copy
                        .Range("A4:Z" & .Range("A3").End(xlDown).row).PasteSpecial xlPasteFormats
                        .Range("C2").Copy
                        .Range("C2").PasteSpecial xlPasteValues
                        Application.CutCopyMode = False
                    End With
                End If
            End With
        
            ' 僅複製有Issue List的工作表
            Dim mypath, myfile As String
            Dim attfile As String
            Dim rng As Range
            If row_b_issue > 5 And row_s_issue > 5 Then
                ' 新增活頁簿
                Workbooks.Add
                ' 將新活頁簿依指定名稱儲存
                mypath = "D:\Users\lindac\Documents\Issue Part\" & Format(Date, "YYYYMM") _
                    & "\" & Format(Date, "MMDD") & "\" & group_short & "\"
                myfile = "Issue Part List_" & Format(Date, "YYYYMMDD") & "_" _
                    & group_short & "_" & pm & "_" & vendor & ".xlsx"
                ActiveWorkbook.SaveAs Filename:=mypath & myfile, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                ' 複製Backlog Issue和Shortage Issue工作表
                Windows("Issue Parts_Kit Parts_template.xlsm").Activate
                Sheets(Array("Summary", "Backlog Issue", "Shortage Issue")).Copy After:=Workbooks(myfile).Sheets(1)
                Application.DisplayAlerts = False
                Sheets(Array("工作表1")).Delete
                ' 刪除Backlog和Shortage Issue裡的第一列
                Sheets("Shortage Issue").rows("1:1").Delete
                Sheets("Backlog Issue").rows("1:1").Delete
                ' 儲存新活頁簿後關閉
                Workbooks(myfile).Close SaveChanges:=True
                
                attfile = mypath & myfile
                If email = "O" Then
                    ' 呼叫副程序自動擬Outlook郵件
                    Generate_PM_Email pm_first, attfile, htmfile, i
                Else
                    Windows("Issue Parts_template_for LC.xlsm").Activate
                    With Worksheets("PM List")
                        Set rng = .Range("B:B").Find(pm, lookat:=xlWhole)
                        If .Range("H" & rng.row) = "" Then
                            .Range("H" & rng.row).Value = attfile
                        Else
                            .Range("H" & rng.row).Value = .Range("H" & rng.row).Value & "," & attfile
                        End If
                    End With
                End If
                
                Windows("Issue Parts_Kit Parts_template.xlsm").Activate
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
                    & group_short & "_" & pm & "_" & vendor & ".xlsx"
                ActiveWorkbook.SaveAs Filename:=mypath & myfile, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                ' 複製Backlog Issue和Shortage Issue工作表
                Windows("Issue Parts_Kit Parts_template.xlsm").Activate
                Sheets(Array("Summary", "Backlog Issue")).Copy After:=Workbooks(myfile).Sheets(1)
                Application.DisplayAlerts = False
                Sheets(Array("工作表1")).Delete
                ' 刪除backlog Issue裡的第一列
                Sheets("Backlog Issue").rows("1:1").Delete
                ' 儲存新活頁簿後關閉
                Workbooks(myfile).Close SaveChanges:=True
                
                attfile = mypath & myfile
                If email = "O" Then
                    ' 呼叫副程序自動擬Outlook郵件
                    Generate_PM_Email pm_first, attfile, htmfile, i
                Else
                    Windows("Issue Parts_template_for LC.xlsm").Activate
                    With Worksheets("PM List")
                        Set rng = .Range("B:B").Find(pm, lookat:=xlWhole)
                        If .Range("H" & rng.row) = "" Then
                            .Range("H" & rng.row).Value = attfile
                        Else
                            .Range("H" & rng.row).Value = .Range("H" & rng.row).Value & "," & attfile
                        End If
                    End With
                End If
                
                Windows("Issue Parts_Kit Parts_template.xlsm").Activate
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
                    & group_short & "_" & pm & "_" & vendor & ".xlsx"
                ActiveWorkbook.SaveAs Filename:=mypath & myfile, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                ' 複製Backlog Issue和Shortage Issue工作表
                Windows("Issue Parts_Kit Parts_template.xlsm").Activate
                Sheets(Array("Summary", "Shortage Issue")).Copy After:=Workbooks(myfile).Sheets(1)
                Application.DisplayAlerts = False
                Sheets(Array("工作表1")).Delete
                ' 刪除shortage Issue裡的第一列
                Sheets("Shortage Issue").rows("1:1").Delete
                ' 儲存新活頁簿後關閉
                Workbooks(myfile).Close SaveChanges:=True
                
                attfile = mypath & myfile
                If email = "O" Then
                    ' 呼叫副程序自動擬Outlook郵件
                    Generate_PM_Email pm_first, attfile, htmfile, i
                Else
                    Windows("Issue Parts_template_for LC.xlsm").Activate
                    With Worksheets("PM List")
                        Set rng = .Range("B:B").Find(pm, lookat:=xlWhole)
                        If .Range("H" & rng.row) = "" Then
                            .Range("H" & rng.row).Value = attfile
                        Else
                            .Range("H" & rng.row).Value = .Range("H" & rng.row).Value & "," & attfile
                        End If
                    End With
                End If
                
                Windows("Issue Parts_Kit Parts_template.xlsm").Activate
                ' 刪除原Shortage Issue裡的資料
                With Worksheets("Shortage Issue")
                    .rows("4:" & .Range("A4").End(xlDown).row).Delete
                End With
            End If
        
            ' 復原Inv. Balance裡的篩選
            With Worksheets("Inv. Balance")
                .Range("A5:CS" & row_inv).AutoFilter Field:=5
                .Range("J1").Select
            End With
        End If
    Next
    
    Application.ScreenUpdating = True
    
    '結算總運行時間
    t0 = Timer - t0
    MsgBox "完成各PM的Issue File" & vbCrLf & t0 \ 60 & "分 " & t0 Mod 60 & "秒"
    
End Sub

Sub Generate_PG_Issue_File()
' 產生PG的Issue Part List, 並儲存在PG資料夾裡

    ' 開啟Timer
    t0 = Timer

    ' 產生用來跑迴圈的PG Table
    Application.ScreenUpdating = False
    Dim row_pg As Integer
    row_pg = Sheets("Group List").Cells(rows.Count, "A").End(xlUp).row
    
    Dim vendor, group, pg_first, group_short As String
    Dim i As Integer
    For i = 2 To row_pg
        With Worksheets("Group List")
            vendor = .Range("A" & i).Value
            group = .Range("B" & i).Value
            pg_first = .Range("E" & i).Value
            group_short = .Range("F" & i).Value
        End With
        ' Raw Data裡篩選群別(PM)
        With Worksheets("Raw Data")
            .AutoFilterMode = False
            Dim row_rdata As Integer
            row_rdata = .Cells(rows.Count, "A").End(xlUp).row
            '.Range("A2:AW" & row_rdata).AutoFilter Field:=3, Criteria1:=vendor
            .Range("A2:AW" & row_rdata).AutoFilter Field:=2, Criteria1:=group
        End With
        
        ' 呼叫副程序匯入資料
        Import_Data
        
        ' 更新Backlog Issue和Shortage Issue工作表的月份
        Update_Month
        
'        ' 刪除non-working part的資料
'        Dim f As Integer
'        f = Sheets("Inv. Balance").Cells(rows.Count, "A").End(xlUp).row
'        For y = f To 6 Step -1
'            If Range("A" & y) = "0" Then
'                rows(y).EntireRow.Delete
'            End If
'        Next
        
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
            Dim row_inv As Integer
            row_inv = .Cells(rows.Count, "O").End(xlUp).row
            ' Inv. Balance裡依金額排序
            '.Range("A5:CZ" & row_inv).Sort key1:=.Range("W5"), order1:=xlDescending, Header:=xlYes
            ' 篩選出Backlog Issue Part
            .Range("A5:CZ" & row_inv).AutoFilter Field:=3, Criteria1:="1"
            Dim row_b_issue As Integer
            row_b_issue = .Cells(rows.Count, "A").End(xlUp).row
            ' 若有Backlog Issue的料, 複製到Backlog Issue工作表裡
            If row_b_issue > 5 Then
                .Range("A6:CZ" & row_b_issue).SpecialCells(xlCellTypeVisible).Copy
                With Worksheets("Backlog Issue")
                    .Range("A4").PasteSpecial xlPasteValues
                    .Range("A1:Z1").Copy
                    .Range("A4:Z" & .Range("A3").End(xlDown).row).PasteSpecial xlPasteFormats
                    .Range("B2").Copy
                    .Range("B2").PasteSpecial xlPasteValues
                    Application.CutCopyMode = False
                End With
            End If
            ' 篩選Shortage Issue Part
            .Range("A5:CZ" & row_inv).AutoFilter Field:=3
            .Range("A5:CZ" & row_inv).AutoFilter Field:=5, Criteria1:="1"
            Dim row_s_issue As Integer
            row_s_issue = .Cells(rows.Count, "A").End(xlUp).row
            ' 若有Shortage Issue的料, 複製到Shortage Issue工作表裡
            If row_s_issue > 5 Then
                .Range("A6:CZ" & row_s_issue).SpecialCells(xlCellTypeVisible).Copy
                With Worksheets("Shortage Issue")
                    .Range("A4").PasteSpecial xlPasteValues
                    .Range("A1:Z1").Copy
                    .Range("A4:Z" & .Range("A3").End(xlDown).row).PasteSpecial xlPasteFormats
                    .Range("C2").Copy
                    .Range("C2").PasteSpecial xlPasteValues
                    Application.CutCopyMode = False
                End With
            End If
        End With
        
        ' 僅複製有Issue List的工作表
        Dim mypath, myfile As String
        If row_b_issue > 5 And row_s_issue > 5 Then
            ' 新增活頁簿
            Workbooks.Add
            ' 將新活頁簿依指定名稱儲存
            mypath = "D:\Users\lindac\Documents\Issue Part\" & Format(Date, "YYYYMM") _
                & "\" & Format(Date, "MMDD") & "\" & group_short & "\"
            myfile = "Issue Part List_" & Format(Date, "YYYYMMDD") & "_" _
                & group_short & "_Kit.xlsx"
'            mypath = "D:\Users\lindac\Documents\Issue Part\" & Format(Date, "YYYYMM") _
'                & "\" & Format(Date, "MMDD") & "\"
'            myfile = "Issue Part List_" & Format(Date, "YYYYMMDD") & "_" & vendor & ".xlsx"
            ActiveWorkbook.SaveAs Filename:=mypath & myfile, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ' 複製Backlog Issue和Shortage Issue工作表
            Windows("Issue Parts_Kit Parts_template.xlsm").Activate
            Sheets(Array("Summary", "Backlog Issue", "Shortage Issue")).Copy After:=Workbooks(myfile).Sheets(1)
            Application.DisplayAlerts = False
            Sheets(Array("工作表1")).Delete
            ' 刪除Backlog和Shortage Issue裡的第一列
            Sheets("Shortage Issue").rows("1:1").Delete
            Sheets("Backlog Issue").rows("1:1").Delete
            ' 儲存新活頁簿後關閉
            Workbooks(myfile).Close SaveChanges:=True
            
'            ' 呼叫副程序自動擬Outlook郵件
'            Dim attfile As String
'            attfile = mypath & myfile
'            Generate_RPM_Email rpm_first, vendor, attfile, htmfile, i

            attfile = mypath & myfile
            Windows("Issue Parts_template_for LC.xlsm").Activate
            With Worksheets("Group List")
                Set rng = .Range("C:C").Find(group_short, lookat:=xlWhole)
                If .Range("G" & rng.row) = "" Then
                    .Range("G" & rng.row).Value = attfile
                Else
                    .Range("G" & rng.row).Value = .Range("G" & rng.row).Value & "," & attfile
                End If
            End With
            
            Windows("Issue Parts_Kit Parts_template.xlsm").Activate
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
                & group_short & "_Kit.xlsx"
'            mypath = "D:\Users\lindac\Documents\Issue Part\" & Format(Date, "YYYYMM") _
'                & "\" & Format(Date, "MMDD") & "\"
'            myfile = "Issue Part List_" & Format(Date, "YYYYMMDD") & "_" & vendor & ".xlsx"
            ActiveWorkbook.SaveAs Filename:=mypath & myfile, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ' 複製Backlog Issue和Shortage Issue工作表
            Windows("Issue Parts_Kit Parts_template.xlsm").Activate
            Sheets(Array("Summary", "Backlog Issue")).Copy After:=Workbooks(myfile).Sheets(1)
            Application.DisplayAlerts = False
            Sheets(Array("工作表1")).Delete
            ' 刪除backlog Issue裡的第一列
            Sheets("Backlog Issue").rows("1:1").Delete
            ' 儲存新活頁簿後關閉
            Workbooks(myfile).Close SaveChanges:=True
            
'            ' 呼叫副程序自動擬Outlook郵件
'            attfile = mypath & myfile
'            Generate_RPM_Email rpm_first, vendor, attfile, htmfile, i

            attfile = mypath & myfile
            Windows("Issue Parts_template_for LC.xlsm").Activate
            With Worksheets("Group List")
                Set rng = .Range("C:C").Find(group_short, lookat:=xlWhole)
                If .Range("G" & rng.row) = "" Then
                    .Range("G" & rng.row).Value = attfile
                Else
                    .Range("G" & rng.row).Value = .Range("G" & rng.row).Value & "," & attfile
                End If
            End With
            
            Windows("Issue Parts_Kit Parts_template.xlsm").Activate
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
                & group_short & "_Kit.xlsx"
'            mypath = "D:\Users\lindac\Documents\Issue Part\" & Format(Date, "YYYYMM") _
'                & "\" & Format(Date, "MMDD") & "\"
'            myfile = "Issue Part List_" & Format(Date, "YYYYMMDD") & "_" & vendor & ".xlsx"
            ActiveWorkbook.SaveAs Filename:=mypath & myfile, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ' 複製Backlog Issue和Shortage Issue工作表
            Windows("Issue Parts_Kit Parts_template.xlsm").Activate
            Sheets(Array("Summary", "Shortage Issue")).Copy After:=Workbooks(myfile).Sheets(1)
            Application.DisplayAlerts = False
            Sheets(Array("工作表1")).Delete
            ' 刪除shortage Issue裡的第一列
            Sheets("Shortage Issue").rows("1:1").Delete
            ' 儲存新活頁簿後關閉
            Workbooks(myfile).Close SaveChanges:=True
            
'            ' 呼叫副程序自動擬Outlook郵件
'            attfile = mypath & myfile
'            Generate_RPM_Email rpm_first, vendor, attfile, htmfile, i

            attfile = mypath & myfile
            Windows("Issue Parts_template_for LC.xlsm").Activate
            With Worksheets("Group List")
                Set rng = .Range("C:C").Find(group_short, lookat:=xlWhole)
                If .Range("G" & rng.row) = "" Then
                    .Range("G" & rng.row).Value = attfile
                Else
                    .Range("G" & rng.row).Value = .Range("G" & rng.row).Value & "," & attfile
                End If
            End With
            
            Windows("Issue Parts_Kit Parts_template.xlsm").Activate
            ' 刪除原Shortage Issue裡的資料
            With Worksheets("Shortage Issue")
                .rows("4:" & .Range("A4").End(xlDown).row).Delete
            End With
        End If
    
        ' 復原Inv. Balance裡的篩選
        With Worksheets("Inv. Balance")
            .Range("A5:CS" & row_inv).AutoFilter Field:=5
            .Range("J1").Select
        End With
    Next
    
    Application.ScreenUpdating = True
    
    '結算總運行時間
    t0 = Timer - t0
    MsgBox "完成各PG的Issue File" & vbCrLf & t0 \ 60 & "分 " & t0 Mod 60 & "秒"
    
End Sub