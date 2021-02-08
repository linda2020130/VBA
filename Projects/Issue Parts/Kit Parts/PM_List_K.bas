Sub Update_PM_List()
' 更新PM List

    With Worksheets("PM List")
        Application.ScreenUpdating = False
        .AutoFilterMode = False
        Dim r_raw, r_pm As Integer
        r_raw = Sheets("Raw Data").Cells(rows.Count, "A").End(xlUp).row
        Sheets("Raw Data").Range("B2:B" & r_raw & ",E2:E" & r_raw & ",C2:C" & r_raw).Copy .Range("L1")
        .Range("O1").Value = "=RC[-2]&RC[-1]"
        .Range("O1").Autofill Destination:=.Range("O1:O" & r_raw), Type:=xlFillDefault
        .Range("O1:O" & r_raw).Copy
        .Range("O1:O" & r_raw).PasteSpecial xlPasteValues
        .Range("L1:O" & r_raw).RemoveDuplicates Columns:=4, Header:=xlYes
        r_raw = .Cells(rows.Count, "L").End(xlUp).row
        r_pm = .Cells(rows.Count, "A").End(xlUp).row
        .Range("J2").Value = "=RC[-2]&RC[-8]"
        .Range("J2").Autofill Destination:=.Range("J2:J" & r_pm), Type:=xlFillDefault
        .Range("J2:J" & r_pm).Copy
        .Range("J2:J" & r_pm).PasteSpecial xlPasteValues
        
        ' 在Raw Data裡但不在List裡的PM
        Dim rng As Range
        Dim count_new, count_delete As Integer
        Dim vendor_pm As String
        count_new = 0
        count_delete = 0
        For i = 2 To r_raw
            vendor_pm = .Range("O" & i).Value
            Set rng = .Range("J:J").Find(vendor_pm, lookat:=xlWhole)
            If rng Is Nothing Then
                .Range("A" & r_pm + 1).Value = .Range("L" & i).Value
                .Range("B" & r_pm + 1).Value = .Range("N" & i).Value
                .Range("H" & r_pm + 1).Value = .Range("M" & i).Value
                count_new = count_new + 1
                r_pm = .Cells(rows.Count, "A").End(xlUp).row
            End If
        Next
        
        ' 不在Raw Data裡但仍在List裡的PM 寫上Delete?
        For j = 2 To r_pm
            vendor_pm = .Range("J" & j).Value
            Set rng = .Range("O:O").Find(vendor_pm, lookat:=xlWhole)
            If rng Is Nothing Then
                .Range("I" & j).Value = "Delete?"
                count_delete = count_delete + 1
            End If
        Next
        
        ' 刪除多餘資料
        .Columns("J:O").Delete
        
        ' 重新依群別排序
        .Range("A1:I" & r_pm).Sort key1:=.Range("A1"), order1:=xlAscending, _
            key2:=Range("B1"), order2:=xlAscending, Header:=xlYes
        .Range("A1:I1").AutoFilter
        .Range("I1").Select
        
        Application.ScreenUpdating = True
        MsgBox "新增: " & count_new & " 個" & vbCrLf & "待刪除: " & count_delete & " 個"
    End With
    
End Sub