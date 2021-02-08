Sub Update_PM_List()
' 更新PM的Database

    With Worksheets("PM Database")
        Application.ScreenUpdating = False
        .AutoFilterMode = False
        Dim r_raw, r_pm As Integer
        r_raw = Sheets("Raw Data").Cells(rows.Count, "A").End(xlUp).row
        Sheets("Raw Data").Range("B2:B" & r_raw & ",E2:E" & r_raw).Copy .Range("J1")
        .Range("J1:K" & r_raw).RemoveDuplicates Columns:=2, Header:=xlYes
        r_raw = .Cells(rows.Count, "J").End(xlUp).row
        r_pm = .Cells(rows.Count, "A").End(xlUp).row

        ' 在Raw Data裡但不存在資料庫的PM
        ' 新增PM與群別至資料庫底部並在Run?欄位寫New?
        Dim rng As Range
        Dim count_new, count_delete As Integer
        Dim pm As String
        count_new = 0
        count_delete = 0
        For i = 2 To r_raw
            pm = .Range("K" & i).Value
            Set rng = .Range("B:B").Find(pm, lookat:=xlWhole)
            If rng Is Nothing Then
                .Range("A" & r_pm + 1).Value = .Range("J" & i).Value
                .Range("B" & r_pm + 1).Value = .Range("K" & i).Value
                .Range("H" & r_pm + 1).Value = "New?"
                count_new = count_new + 1
                r_pm = .Cells(rows.Count, "A").End(xlUp).row
            End If
        Next
        
        ' 不在Raw Data裡但仍在資料庫裡的PM
        ' 若原本在Run?欄位是OK, 則改寫成Delete?
        ' 已確認為X的保持不變
        For j = 2 To r_pm
            pm = .Range("B" & j).Value
            Set rng = .Range("K:K").Find(pm, lookat:=xlWhole)
            If rng Is Nothing Then
                If .Range("H" & j).Value = "OK" Then
                    .Range("H" & j).Value = "Delete?"
                    count_delete = count_delete + 1
                End If
            End If
        Next
        
        ' 刪除複製過來的資料
        .Columns("J:K").Delete
        
        ' 重新依群別排序
        .Range("A1:H" & r_pm).Sort key1:=.Range("A1"), order1:=xlAscending, _
            key2:=Range("B1"), order2:=xlAscending, Header:=xlYes
        .Range("A1:H1").AutoFilter
        
        Application.ScreenUpdating = True
        MsgBox "新增: " & count_new & " 個" & vbCrLf & "待刪除: " & count_delete & " 個"
    End With
    
End Sub
