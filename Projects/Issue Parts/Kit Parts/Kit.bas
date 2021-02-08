Sub Add_Backlog_FCST_to_Components()
' 將套件料的Backlog和FCST加到外包料上

    't0 = Timer
    'Application.ScreenUpdating = False
    ' 將外包料的Backlog數字歸零(避免後續產生重複加的情形)
    With Worksheets("Inv. Balance")
        .Outline.ShowLevels ColumnLevels:=3
        Dim row_inv As Integer
        row_inv = .Cells(rows.Count, "I").End(xlUp).row
        .Range("A5:CZ" & row_inv).AutoFilter Field:=9, Criteria1:="外包料"
        .Range("U6:V" & row_inv & ",AC6:AD" & row_inv & ",AK6:AL" & row_inv & ",AS6:AT" & row_inv & _
            ",BA6:BB" & row_inv & ",BI6:BJ" & row_inv).FormulaR1C1 = "0"
        .Range("A5:CZ" & row_inv).AutoFilter Field:=9
        .Outline.ShowLevels ColumnLevels:=1
    End With
    
    ' 抓Backlog和FCST數字到Kit Table工作表裡
    With Worksheets("Kit Table")
        .Range("G3").FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC4,'Inv. Balance'!C15:C64,8,FALSE)),0,VLOOKUP(RC4,'Inv. Balance'!C15:C64,8,FALSE))*RC6"
        .Range("H3").FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC4,'Inv. Balance'!C15:C64,4,FALSE)),0,VLOOKUP(RC4,'Inv. Balance'!C15:C64,4,FALSE))*RC6"
        .Range("I3").FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC4,'Inv. Balance'!C15:C64,16,FALSE)),0,VLOOKUP(RC4,'Inv. Balance'!C15:C64,16,FALSE))*RC6"
        .Range("J3").FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC4,'Inv. Balance'!C15:C64,12,FALSE)),0,VLOOKUP(RC4,'Inv. Balance'!C15:C64,12,FALSE))*RC6"
        .Range("K3").FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC4,'Inv. Balance'!C15:C64,24,FALSE)),0,VLOOKUP(RC4,'Inv. Balance'!C15:C64,24,FALSE))*RC6"
        .Range("L3").FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC4,'Inv. Balance'!C15:C64,20,FALSE)),0,VLOOKUP(RC4,'Inv. Balance'!C15:C64,20,FALSE))*RC6"
        .Range("M3").FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC4,'Inv. Balance'!C15:C64,32,FALSE)),0,VLOOKUP(RC4,'Inv. Balance'!C15:C64,32,FALSE))*RC6"
        .Range("N3").FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC4,'Inv. Balance'!C15:C64,28,FALSE)),0,VLOOKUP(RC4,'Inv. Balance'!C15:C64,28,FALSE))*RC6"
        .Range("O3").FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC4,'Inv. Balance'!C15:C64,40,FALSE)),0,VLOOKUP(RC4,'Inv. Balance'!C15:C64,40,FALSE))*RC6"
        .Range("P3").FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC4,'Inv. Balance'!C15:C64,36,FALSE)),0,VLOOKUP(RC4,'Inv. Balance'!C15:C64,36,FALSE))*RC6"
        .Range("Q3").FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC4,'Inv. Balance'!C15:C64,48,FALSE)),0,VLOOKUP(RC4,'Inv. Balance'!C15:C64,48,FALSE))*RC6"
        .Range("R3").FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC4,'Inv. Balance'!C15:C64,44,FALSE)),0,VLOOKUP(RC4,'Inv. Balance'!C15:C64,44,FALSE))*RC6"
        .Range("S3").FormulaR1C1 = "=SUM(RC7,RC9,RC11,RC13,RC15,RC17)"
        .Range("T3").FormulaR1C1 = "=SUM(RC8,RC10,RC12,RC14,RC16,RC18)"
        .Range("G3:T3").Autofill Destination:=.Range("G3:T" & .Range("E3").End(xlDown).row)
        .Range("G3:T" & .Range("E3").End(xlDown).row).Copy
        .Range("G3:T" & .Range("E3").End(xlDown).row).PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End With
    
    '將套件料的Backlog和FCST數字加進外包料的Backlog數字裡
    Dim total_backlog, total_fcst As Single
    Dim n As Integer
    Dim rng, rng2 As Range
    Dim pn_son, pn_mom As String
    n = Sheets("Kit Table").Cells(rows.Count, "E").End(xlUp).row
    
    For r = 3 To n
        total_backlog = Sheets("Kit Table").Range("S" & r).Value
        total_fcst = Sheets("Kit Table").Range("T" & r).Value
        '將套件料的Backlog數字加進外包料的Backlog數字裡
        If total_backlog > 0 Then
            pn_son = Sheets("Kit Table").Range("E" & r)
            pn_mom = Sheets("Kit Table").Range("D" & r)
            Set rng = Sheets("Inv. Balance").Range("O:O").Find(pn_son, lookat:=xlWhole)
            
            ' 若找不到外包料號, 則於資料末端插入新的一列
            If rng Is Nothing Then
                With Worksheets("Inv. Balance")
                    Set rng2 = .Range("O:O").Find(pn_mom, lookat:=xlWhole)
                    .Range("I" & row_inv + 1).Value = "外包料"
                    .Range("I" & row_inv).Copy
                    .Range("I" & row_inv + 1).PasteSpecial xlPasteFormats
                    .Range("J" & rng2.row & ":N" & rng2.row).Copy .Range("J" & row_inv + 1 & ":N" & row_inv + 1)
                    Application.CutCopyMode = False
                    .Range("O" & row_inv + 1).Value = pn_son
                    .Range("P" & row_inv + 1).Value = .Range("P" & rng2.row).Value
                    .Range("Q" & row_inv + 1 & ":BL" & row_inv + 1).Value = 0
                    Set rng = .Range("O:O").Find(pn_son, lookat:=xlWhole)
                    row_inv = .Cells(rows.Count, "I").End(xlUp).row
                End With
            End If
            
            With Worksheets("Inv. Balance")
                .Range("V" & rng.row).Value = .Range("V" & rng.row).Value + Sheets("Kit Table").Range("G" & r).Value
                .Range("AD" & rng.row).Value = .Range("AD" & rng.row).Value + Sheets("Kit Table").Range("I" & r).Value
                .Range("AL" & rng.row).Value = .Range("AL" & rng.row).Value + Sheets("Kit Table").Range("K" & r).Value
                .Range("AT" & rng.row).Value = .Range("AT" & rng.row).Value + Sheets("Kit Table").Range("M" & r).Value
                .Range("BB" & rng.row).Value = .Range("BB" & rng.row).Value + Sheets("Kit Table").Range("O" & r).Value
                .Range("BJ" & rng.row).Value = .Range("BJ" & rng.row).Value + Sheets("Kit Table").Range("Q" & r).Value
            End With
        End If
        
        '將套件料的FCST數字加進外包料的FCST數字裡
        If total_fcst > 0 Then
            pn_son = Sheets("Kit Table").Range("E" & r)
            pn_mom = Sheets("Kit Table").Range("D" & r)
            Set rng = Sheets("Inv. Balance").Range("O:O").Find(pn_son, lookat:=xlWhole)
            
            ' 若找不到外包料號, 則於資料末端插入新的一列
            If rng Is Nothing Then
                With Worksheets("Inv. Balance")
                    Set rng2 = .Range("O:O").Find(pn_mom, lookat:=xlWhole)
                    .Range("I" & row_inv + 1).Value = "外包料"
                    .Range("I" & row_inv).Copy
                    .Range("I" & row_inv + 1).PasteSpecial xlPasteFormats
                    .Range("J" & rng2.row & ":N" & rng2.row).Copy .Range("J" & row_inv + 1 & ":N" & row_inv + 1)
                    Application.CutCopyMode = False
                    .Range("O" & row_inv + 1).Value = pn_son
                    .Range("P" & row_inv + 1).Value = .Range("P" & rng2.row).Value
                    .Range("Q" & row_inv + 1 & ":BL" & row_inv + 1).Value = 0
                    Set rng = .Range("O:O").Find(pn_son, lookat:=xlWhole)
                    row_inv = .Cells(rows.Count, "I").End(xlUp).row
                End With
            End If
            
            With Worksheets("Inv. Balance")
                .Range("R" & rng.row).Value = .Range("R" & rng.row).Value + Sheets("Kit Table").Range("H" & r).Value
                .Range("Z" & rng.row).Value = .Range("Z" & rng.row).Value + Sheets("Kit Table").Range("J" & r).Value
                .Range("AH" & rng.row).Value = .Range("AH" & rng.row).Value + Sheets("Kit Table").Range("L" & r).Value
                .Range("AP" & rng.row).Value = .Range("AP" & rng.row).Value + Sheets("Kit Table").Range("N" & r).Value
                .Range("AX" & rng.row).Value = .Range("AX" & rng.row).Value + Sheets("Kit Table").Range("P" & r).Value
                .Range("BF" & rng.row).Value = .Range("BF" & rng.row).Value + Sheets("Kit Table").Range("R" & r).Value
            End With
        End If
    Next
    
    ' 將Kit Table的Backlog和FCST數字欄位刪除
    With Worksheets("Kit Table")
        .Columns("G:T").Delete
        .Range("A1").Copy
        .Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End With
    'Application.ScreenUpdating = True
    
    ' MsgBox "套件料的backlog和FCST數量已寫入外包料"
    't0 = Timer - t0
    'MsgBox "數值寫入外包料花費時間 " & vbCrLf & t0 \ 60 & "分 " & t0 Mod 60 & "秒"
    
End Sub

Sub Remove_Batch()
' 一般料的批次去除

    With Worksheets("Inv. Balance")
        Dim row_inv As Integer
        Dim edompn As String
        row_inv = .Cells(rows.Count, "O").End(xlUp).row
        For i = 6 To row_inv
            If .Range("I" & i) = "一般料" Then
                edompn = .Range("O" & i).Value
                If InStr(edompn, "(") > 0 Then
                    .Range("O" & i).Value = Mid(edompn, 1, InStr(edompn, "(") - 1)
                End If
            End If
        Next
    End With
    
    ' MsgBox "一般料去批次完成"

End Sub

Sub Combine_Common_Parts()
' 將相同料號的資料列合併成一列
    'Application.ScreenUpdating = False
    With Worksheets("Inv. Balance")
        Dim row_inv As Integer
        row_inv = .Cells(rows.Count, "O").End(xlUp).row
        .Range("A5:CZ" & row_inv).Sort key1:=Range("J5"), order1:=xlDescending, _
        key2:=Range("I5"), order2:=xlAscending, key3:=Range("O5"), order3:=xlAscending, Header:=xlYes
        For j = row_inv To 7 Step -1
            If .Range("I" & j) = "一般料" And .Range("I" & j - 1) = "一般料" Then
                If .Range("O" & j) = .Range("O" & j - 1) Then
                    .Range("Q" & j & ":AT" & j).Copy
                    .Range("Q" & j - 1 & ":AT" & j - 1).PasteSpecial Paste:=xlPasteAll, Operation:=xlAdd, _
                        SkipBlanks:=False, Transpose:=False
                    .rows(j).Delete
                End If
            End If
        Next
    End With
    'Application.ScreenUpdating = True
    ' MsgBox "同料號資料列合併完成"

End Sub

Sub Delete_Set_Parts()
' 將套件料的資料列刪除
    
    With Worksheets("Inv. Balance")
        Dim n As Integer
        n = .Cells(rows.Count, "O").End(xlUp).row
        For i = n To 6 Step -1
            If .Range("I" & i) = "套件料" Then
                .rows(i).Delete
            End If
        Next
    End With
    
    ' MsgBox "套件料已刪除"

End Sub