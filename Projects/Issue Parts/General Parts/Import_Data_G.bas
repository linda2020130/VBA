Sub Import_Data()
' 從Raw Data匯入資料到Inv. Balance

    'Application.ScreenUpdating = False
    
    ' 復原篩選, 刪除Inv. Balance裡原有資料
    With Worksheets("Inv. Balance")
        .AutoFilterMode = False
        .Range("6:" & Range("A5").End(xlDown).row).Delete
    End With
    
    ' 從Raw Data複製資料貼到Inv. Balance的指定位置
    With Worksheets("Raw Data")
        Dim i As Integer
        i = .Cells(rows.Count, "A").End(xlUp).row
        .Range("A3:H" & i & ",L3:R" & i & ",V3:AB" & i & ",AF3:AL" & i & ",AP3:AV" & i & _
            ",AZ3:BF" & i & ",BJ3:BO" & i & ",CD3:CD" & i).SpecialCells(xlCellTypeVisible).Copy
        Sheets("Inv. Balance").Range("I6").PasteSpecial xlPasteValues
    End With
    
    ' 前後分別複製貼上Issue Part判斷公式,並複製貼上中間Backlog顏色格式
    With Worksheets("Inv. Balance")
        .Range("BG2:CS2").Copy .Range("BG6")
        Dim row_part As Integer
        row_part = .Cells(rows.Count, "L").End(xlUp).row
        If row_part > 6 Then
            .Range("BG6:CS6").Autofill Destination:=.Range("BG6:CS" & row_part)
        End If
        .Range("BG6:CS" & row_part).Copy
        .Range("BG6:CS" & row_part).PasteSpecial xlPasteValues
        .Range("A2:H2").Copy .Range("A6")
        If row_part > 6 Then
            .Range("A6:H6").Autofill Destination:=.Range("A6:H" & row_part)
        End If
        .Range("A6:H" & row_part).Copy
        .Range("A6:H" & row_part).PasteSpecial xlPasteValues
        .Range("R2:BB2").Copy
        .Range("R6:BB" & row_part).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        .Range("A5:CS5").AutoFilter
        .Range("I1").Select
    End With
    
    'Application.ScreenUpdating = True
    
End Sub
