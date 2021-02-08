Sub Import_Data()
' 從Raw Data匯入資料到Inv. Balance
    
    't0 = Timer
    'Application.ScreenUpdating = False
    ' 復原篩選, 刪除Inv. Balance裡原有資料
    With Worksheets("Inv. Balance")
        .AutoFilterMode = False
        .Range("6:" & .Range("L5").End(xlDown).row).Delete
    End With
    
    ' 從Raw Data複製資料
    With Worksheets("Raw Data")
        .Range("A3:BC" & .Range("A2").End(xlDown).row).SpecialCells(xlCellTypeVisible).Copy
        Sheets("Inv. Balance").Range("J6").PasteSpecial xlPasteValues
    End With
    
    ' 在Inv. Balance裡複製貼上料號類型公式
    With Worksheets("Inv. Balance")
        Dim row_inv As Integer
        row_inv = .Cells(rows.Count, "L").End(xlUp).row
        If row_inv > 5 Then
            .Range("I2").Copy
            .Range("I6:I" & row_inv).PasteSpecial xlPasteFormulas
            .Range("I6:I" & row_inv).Copy
            .Range("I6:I" & row_inv).PasteSpecial xlPasteValues
            Application.CutCopyMode = False
        End If
    End With
    
    ' 將套件料的Backlog寫入外包料
    Add_Backlog_FCST_to_Components
    ' 將一般料的批次刪除
    Remove_Batch
    ' 將一般料的同料號資料列合併
    Combine_Common_Parts
    ' 刪除套件料
    Delete_Set_Parts
    
    With Worksheets("Inv. Balance")
        .AutoFilterMode = False
        row_inv = .Cells(rows.Count, "O").End(xlUp).row
        If row_inv > 5 Then
            ' 前後分別複製貼上Issue Part判斷公式
            .Range("BM2:CZ2").Copy .Range("BM6:CZ" & row_inv)
            .Range("BM6:CZ" & row_inv).Copy
            .Range("BM6:CZ" & row_inv).PasteSpecial xlPasteValues
            .Range("A2:H2").Copy .Range("A6:H" & row_inv)
            .Range("A6:H" & row_inv).Copy
            .Range("A6:H" & row_inv).PasteSpecial xlPasteValues
            Application.CutCopyMode = False
            With .Range("A6:H" & row_inv, "BM6:CZ" & row_inv).Font
                .ThemeColor = xlThemeColorLight1
            End With
            ' 複製貼上中間Backlog顏色格式
            .Range("Q2:BL2").Copy
            .Range("Q6:BL" & row_inv).PasteSpecial xlPasteFormats
        End If
        ' 指定列加上篩選
        .Range("A5:CZ5").AutoFilter
    End With
    
    'Application.ScreenUpdating = True
    ' MsgBox "資料匯入完成。"
    't0 = Timer - t0
    'MsgBox "資料匯入完成。" & vbCrLf & "花費時間: " & t0 \ 60 & "分 " & t0 Mod 60 & "秒"
    
End Sub