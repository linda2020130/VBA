Sub Adjust_Data()
' 將下載下來的資料整理成可匯入形式
' 清除數值皆為0的資料列
' 調整群別名稱

    Application.ScreenUpdating = False
    With Worksheets("Raw Data")
        .AutoFilterMode = False
        Dim k As Integer
        k = .Cells(rows.Count, "A").End(xlUp).row
        ' 將數值皆為0的資料列刪除
        .Range("BD3").FormulaR1C1 = "=IF(SUM(RC8:RC55)=0,0,1)"
        .Range("BD3").Autofill Destination:=Range("BD3:BD" & k)
        .Range("BD3:BD" & k).Copy
        .Range("BD3:BD" & k).PasteSpecial xlPasteValues
        .Range("A2:BD" & k).Sort key1:=.Range("BD2"), order1:=xlAscending, Header:=xlYes
        .Range("A2:BD" & k).AutoFilter Field:=56, Criteria1:="0"
        Dim row_0 As Integer
        row_0 = .Cells(rows.Count, "A").End(xlUp).row
        If row_0 >= 3 Then
            .rows("3:" & row_0).Delete
        End If
        .AutoFilterMode = False
        .Columns("BD").Delete
        k = .Cells(rows.Count, "A").End(xlUp).row
        ' 將北群二,三改成北群
        Dim rng As Range
        Set rng = .Range("B:B").Find("北中國事業群", lookat:=xlPart)
        If Not rng Is Nothing Then
            .Range("A2:BC" & k).AutoFilter Field:=2, Criteria1:="=北中國事業群二", _
                Operator:=xlOr, Criteria2:="=北中國事業群三"
            .Range("B3:B" & .Range("B3").End(xlDown).row).Replace "二", ""
            .Range("B3:B" & .Range("B3").End(xlDown).row).Replace "三", ""
            .AutoFilterMode = False
        End If
        ' 將南群一,三改成南群
        Set rng = .Range("B:B").Find("南中國事業群", lookat:=xlPart)
        If Not rng Is Nothing Then
            .Range("A2:BC" & k).AutoFilter Field:=2, Criteria1:="=南中國事業群一", _
                Operator:=xlOr, Criteria2:="=南中國事業群三"
            .Range("B3:B" & .Range("B3").End(xlDown).row).Replace "一", ""
            .Range("B3:B" & .Range("B3").End(xlDown).row).Replace "三", ""
            .AutoFilterMode = False
        End If
        .Range("A2:BC2").AutoFilter
        .Range("A1").Select
    End With
    Application.ScreenUpdating = True
    MsgBox "資料清理完畢"
    
End Sub