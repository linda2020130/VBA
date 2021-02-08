Sub Adjust_Data()
' 將下載下來的資料整理成可匯入形式

    Application.ScreenUpdating = False
    With Worksheets("Raw Data")
        .AutoFilterMode = False
        Dim k As Integer
        k = .Cells(rows.Count, "A").End(xlUp).row
        ' 將數值皆為0的資料列刪除
        .Range("CE3").FormulaR1C1 = "=IF(SUM(RC8:RC77)=0,0,1)"
        .Range("CE3").Autofill Destination:=Range("CE3:CE" & k)
        .Range("CE3:CE" & k).Copy
        .Range("CE3:CE" & k).PasteSpecial xlPasteValues
        .Range("A2:CE" & k).Sort key1:=.Range("CE2"), order1:=xlAscending, Header:=xlYes
        .Range("A2:CE" & k).AutoFilter Field:=83, Criteria1:="0"
        Dim row_0 As Integer
        row_0 = .Cells(rows.Count, "A").End(xlUp).row
        If row_0 >= 3 Then
            .rows("3:" & row_0).Delete
        End If
        .AutoFilterMode = False
        .Columns("CE").Delete
        k = .Cells(rows.Count, "A").End(xlUp).row
        ' 將北群二,三改成北群
        Dim rng As Range
        Set rng = .Range("B:B").Find("北中國事業群", lookat:=xlPart)
        If Not rng Is Nothing Then
            .Range("A2:CD" & k).AutoFilter Field:=2, Criteria1:="=北中國事業群二", _
                Operator:=xlOr, Criteria2:="=北中國事業群三"
            .Range("B3:B" & .Range("B3").End(xlDown).row).Replace "二", ""
            .Range("B3:B" & .Range("B3").End(xlDown).row).Replace "三", ""
            .AutoFilterMode = False
        End If
        ' 將南群一,三改成南群
        Set rng = .Range("B:B").Find("南中國事業群", lookat:=xlPart)
        If Not rng Is Nothing Then
            .Range("A2:CD" & k).AutoFilter Field:=2, Criteria1:="=南中國事業群一", _
                Operator:=xlOr, Criteria2:="=南中國事業群三"
            .Range("B3:B" & .Range("B3").End(xlDown).row).Replace "一", ""
            .Range("B3:B" & .Range("B3").End(xlDown).row).Replace "三", ""
            .AutoFilterMode = False
        End If
        ' 代入公式計算單價
        .Range("BZ3").FormulaR1C1 = "=IF(RC12=0,0,RC8/RC12)"
        .Range("CA3").FormulaR1C1 = "=IF(RC73=0,0,RC69/RC73)"
        .Range("CB3").FormulaR1C1 = "=IF(RC74=0,0,RC70/RC74)"
        .Range("CC3").FormulaR1C1 = "=IF(RC75=0,0,RC71/RC75)"
        .Range("CD3").FormulaR1C1 = "=IF(RC78=0,IF(RC79=0,IF(RC80=0,IF(RC81=0,0,RC81),RC80),RC79),RC78)"
        ' 轉單價公式為數值
        .Range("BZ3:CD3").Autofill Destination:=.Range("BZ3:CD" & k)
        .Range("BZ3:CD" & k).Copy
        .Range("BZ3:CD" & k).PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        .Range("A2:CD2").AutoFilter
        .Range("A1").Select
    End With
    Application.ScreenUpdating = True
    MsgBox "單價計算完畢"
    
End Sub