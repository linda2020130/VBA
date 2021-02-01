Sub GetBalance()

    Dim rowCount, columnCount, totalRow As Integer
    totalRow = Range("D1").End(xlDown).Row
    columnCount = Range("A1").End(xlToRight).Column
    

    ' FCST 欄位
    Dim fixedColumn, firstWeekOfFixedColumn, columnFcst, columnEndWeek As Integer
    Dim fixedColumnMonth As String
    fixedColumn = Sheets("User Input").Range("B2").Value + 6
    fixedColumnMonth = Left(Range(Chr(fixedColumn + 64) & 1).Value, 2)
    columnFcst = 3
    columnEndWeek = 0
    Rows("1:1").AutoFilter
    ActiveSheet.Range("A1:" & Chr(columnCount + 64) & totalRow).AutoFilter Field:=6, Criteria1:=Array( _
        "FCST-MB", "FCST-NB"), Operator:=xlFilterValues
    rowCount = Range("F1").End(xlDown).Row
    ' FCST - by week and month
    For i = 7 To columnCount
        If Asc(Left(Range(Chr(i + 64) & 1).Value, 1)) < 65 Then
            If Left(Range(Chr(i + 64) & 1).Value, 2) <> Left(Range(Chr(i + 63) & 1).Value, 2) Then
                Range(Chr(i + 64) & "2:" & Chr(i + 64) & rowCount).SpecialCells(xlCellTypeVisible).FormulaR1C1 = _
                    "=IF(LEFT(R1C,2)=LEFT(R1C[-1],2),"""",IF(RC6=""FCST-MB"",IFERROR(VLOOKUP(C5,'MB FCST'!C1:C13," & columnFcst & ",0),0),IFERROR(VLOOKUP(C5,'NB FCST'!C1:C13," & columnFcst & ",0),0)))"
                    columnFcst = columnFcst + 1
                If Left(Range(Chr(i + 64) & 1).Value, 2) = fixedColumnMonth Then
                    firstWeekOfFixedColumn = i
                End If
            End If
        Else
            If columnEndWeek = 0 Then
                columnEndWeek = i - 1
            End If
            Range(Chr(i + 64) & "2:" & Chr(i + 64) & rowCount).SpecialCells(xlCellTypeVisible).FormulaR1C1 = _
                "=IF(LEFT(R1C,2)=LEFT(R1C[-1],2),"""",IF(RC6=""FCST-MB"",IFERROR(VLOOKUP(C5,'MB FCST'!C1:C13," & columnFcst & ",0),0),IFERROR(VLOOKUP(C5,'NB FCST'!C1:C13," & columnFcst & ",0),0)))"
            columnFcst = columnFcst + 1
        End If
    Next
    ' FCST - fixedColumn
    If fixedColumn = firstWeekOfFixedColumn + 1 Then
        Range(Chr(fixedColumn + 64) & "2:" & Chr(fixedColumn + 64) & rowCount).SpecialCells(xlCellTypeVisible).FormulaR1C1 = _
            "=IF(LEFT(R[2]C[-" & fixedColumn - 6 & "],3)=""MRP"", MAX(RC[-1]-R[2]C[-1],0), MAX(RC[-1]-R[1]C[-1], 0))"
    ElseIf fixedColumn = firstWeekOfFixedColumn + 2 Then
        Range(Chr(fixedColumn + 64) & "2:" & Chr(fixedColumn + 64) & rowCount).SpecialCells(xlCellTypeVisible).FormulaR1C1 = _
            "=IF(LEFT(R[2]C[-" & fixedColumn - 6 & "],3)=""MRP"", MAX(RC[-2]-R[2]C[-2]-R[2]C[-1],0), MAX(RC[-2]-R[1]C[-2]-R[1]C[-1], 0))"
    ElseIf fixedColumn = firstWeekOfFixedColumn + 3 Then
        Range(Chr(fixedColumn + 64) & "2:" & Chr(fixedColumn + 64) & rowCount).SpecialCells(xlCellTypeVisible).FormulaR1C1 = _
            "=IF(LEFT(R[2]C[-" & fixedColumn - 6 & "],3)=""MRP"", MAX(RC[-3]-R[2]C[-3]-R[2]C[-2]-R[2]C[-1],0), MAX(RC[-3]-R[1]C[-3]-R[1]C[-2]-R[1]C[-1], 0))"
    ElseIf fixedColumn = firstWeekOfFixedColumn + 4 Then
        Range(Chr(fixedColumn + 64) & "2:" & Chr(fixedColumn + 64) & rowCount).SpecialCells(xlCellTypeVisible).FormulaR1C1 = _
            "=IF(LEFT(R[2]C[-" & fixedColumn - 6 & "],3)=""MRP"", MAX(RC[-4]-R[2]C[-4]-R[2]C[-3]-R[2]C[-2]-R[2]C[-1],0), MAX(RC[-4]-R[1]C[-4]-R[1]C[-3]-R[1]C[-2]-R[1]C[-1], 0))"
    End If
    
    ' MRP 欄位
    ActiveSheet.Range("A1:" & Chr(columnCount + 64) & totalRow).AutoFilter Field:=6, Criteria1:=Array( _
        "MRP", "MRP-MB", "MRP-NB"), Operator:=xlFilterValues
    rowCount = Range("F1").End(xlDown).Row
    For i = 7 To columnCount
        Range(Chr(i + 64) & "2:" & Chr(i + 64) & rowCount).SpecialCells(xlCellTypeVisible).FormulaR1C1 = _
        "=IF(RC6=""MRP-NB"",0,IFERROR(VLOOKUP(C5,MRP!C1:C20," & i - 5 & ",0),0))"
    Next
    
    ' Backlog 欄位
    ActiveSheet.Range("A1:" & Chr(columnCount + 64) & totalRow).AutoFilter Field:=6, Criteria1:="Backlog", Operator:=xlFilterValues
    rowCount = Range("F1").End(xlDown).Row
    For i = 8 To columnCount
        Range(Chr(i + 64) & "2:" & Chr(i + 64) & rowCount).SpecialCells(xlCellTypeVisible).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(C4,Backlog!C1:C20," & i - 5 & ",0),0)"
    Next

    ' Shipment 欄位
    ActiveSheet.Range("A1:" & Chr(columnCount + 64) & totalRow).AutoFilter Field:=6, Criteria1:=Array("Shipment-MB", "Shipment-NB"), Operator:=xlFilterValues
    rowCount = Range("F1").End(xlDown).Row
    For i = 7 To columnCount
        Range(Chr(i + 64) & "2:" & Chr(i + 64) & rowCount).SpecialCells(xlCellTypeVisible).FormulaR1C1 = _
           "=IF(RC6=""Shipment-MB"", IFERROR(VLOOKUP(C5,'MB Shipment'!C1:C20," & i - 5 & ",0),0), IFERROR(VLOOKUP(C5,'NB Shipment'!C1:C20," & i - 5 & ",0),0))"
    Next
    
    ' Balance-MRP 欄位
    ActiveSheet.Range("A1:" & Chr(columnCount + 64) & totalRow).AutoFilter Field:=6, Criteria1:="Balance-MRP", Operator:=xlFilterValues
    rowCount = Range("F1").End(xlDown).Row
    ' First week
    Range("G2:G" & rowCount).SpecialCells(xlCellTypeVisible).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(C4,Stock!C1:C2,2,0),0)"
    ' Update fixedColumn
    ActiveSheet.AutoFilterMode = False
    Dim fcstSum, mrpSum As Long
    For i = 6 To rowCount
        If Range("F" & i).Value = "Balance-MRP" Then
            ' 非共用料
            If Left(Range("F" & i - 4).Value, 4) = "FCST" Then
                fcstSum = Application.WorksheetFunction.Sum(Range(Chr(fixedColumn + 64) & i - 4 & ":" & Chr(columnEndWeek + 64) & i - 4))
                mrpSum = Application.WorksheetFunction.Sum(Range(Chr(fixedColumn + 64) & i - 3 & ":" & Chr(columnEndWeek + 64) & i - 3))
                ' Second week
                Range("H" & i).FormulaR1C1 = "=R[-1]C[-1]+RC[-1]-R[-3]C+R[-2]C"
                ' Third week to last week
                Range("I" & i & ":" & Chr(columnEndWeek + 64) & i).FormulaR1C1 = "=RC[-1]-R[-3]C+R[-2]C"
                ' Month欄
                Range(Chr(columnEndWeek + 65) & i & ":" & Chr(columnCount + 64) & i).FormulaR1C1 = "=RC[-1]-R[-4]C+R[-2]C"
                ' 若FCST>MRP, 更新Balance-MRP算法
                If fcstSum > mrpSum Then
                    Range(Chr(fixedColumn + 64) & i & ":" & Chr(columnEndWeek + 64) & i).FormulaR1C1 = "=RC[-1]-R[-4]C+R[-2]C"
                ' 若FCST<MRP, 標記顏色
                ElseIf fcstSum < mrpSum Then
                    Range(Chr(fixedColumn + 64) & i).Interior.Color = 65535
                End If
            ' 共用料
            Else
                fcstSum = Application.WorksheetFunction.Sum(Range(Chr(fixedColumn + 64) & i - 7 & ":" & Chr(columnEndWeek + 64) & i - 6))
                mrpSum = Application.WorksheetFunction.Sum(Range(Chr(fixedColumn + 64) & i - 5 & ":" & Chr(columnEndWeek + 64) & i - 4))
                ' Second week
                Range("H" & i).FormulaR1C1 = "=R[-2]C[-1]+R[-1]C[-1]+RC[-1]-R[-4]C-R[-5]C+R[-3]C"
                ' Third week to last week
                Range("I" & i & ":" & Chr(columnEndWeek + 64) & i).FormulaR1C1 = "=RC[-1]-R[-4]C-R[-5]C+R[-3]C"
                ' Month欄
                Range(Chr(columnEndWeek + 65) & i & ":" & Chr(columnCount + 64) & i).FormulaR1C1 = "=RC[-1]-R[-6]C-R[-7]C+R[-3]C"
                ' 若FCST>MRP, 更新Balance-MRP算法
                If fcstSum > mrpSum Then
                    Range(Chr(fixedColumn + 64) & i & ":" & Chr(columnEndWeek + 64) & i).FormulaR1C1 = "=RC[-1]-R[-6]C-R[-7]C+R[-3]C"
                ' 若FCST<MRP, 標記顏色
                ElseIf fcstSum < mrpSum Then
                    Range(Chr(fixedColumn + 64) & i).Interior.Color = 65535
                End If
            End If
        End If
    Next
    
    
    ' Balance-Shipment 欄位
    ActiveSheet.Range("A1:" & Chr(columnCount + 64) & totalRow).AutoFilter Field:=6, Criteria1:="Balance-Shipment", Operator:=xlFilterValues
    rowCount = Range("F1").End(xlDown).Row
    Range("H2:H" & rowCount).SpecialCells(xlCellTypeVisible).FormulaR1C1 = _
        "=IF(R[-4]C6=""Backlog"", R[-1]C[-1]+R[-4]C-R[-3]C-R[-2]C,R[-1]C[-1]-R[-2]C+R[-3]C)"
    Range("I2:" & Chr(columnCount + 64) & rowCount).SpecialCells(xlCellTypeVisible).FormulaR1C1 = _
        "=IF(R[-4]C6=""Backlog"", RC[-1]-R[-2]C-R[-3]C+R[-4]C, RC[-1]-R[-2]C+R[-3]C)"
    ' 複製格式
    Range("G2:" & Chr(columnCount + 64) & rowCount).SpecialCells(xlCellTypeVisible).Borders(xlEdgeBottom).LineStyle = xlDouble
        
    ' R/Y/G 欄位
    Dim weekR, weekY As Integer
    weekR = Sheets("User Input").Range("B3").Value
    weekY = Sheets("User Input").Range("B4").Value
    ActiveSheet.Range("A1:" & Chr(columnCount + 64) & totalRow).AutoFilter Field:=6, Criteria1:="Balance-MRP", Operator:=xlFilterValues
    rowCount = Range("F1").End(xlDown).Row
    Range("C2:C" & rowCount).SpecialCells(xlCellTypeVisible).FormulaR1C1 = _
        "=IF(COUNTIF(RC[4]:R[1]C[" & weekR + 3 & "],""<0"")>0,""R"",IF(COUNTIF(RC[" & weekR + 4 & "]:R[1]C[" & weekY + 3 & "],""<0"")>0,""Y"",""G""))"
        
    ActiveSheet.Range("A1:" & Chr(columnCount + 64) & totalRow).AutoFilter Field:=6, Criteria1:="Balance-Shipment", Operator:=xlFilterValues
    rowCount = Range("F1").End(xlDown).Row
    Range("C2:C" & rowCount).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=R[-1]C"
    
    ActiveSheet.Range("A1:" & Chr(columnCount + 64) & totalRow).AutoFilter Field:=6, Criteria1:=Array( _
        "FCST-MB", "FCST-NB", "MRP-MB", "MRP-NB", "Backlog", "Shipment-MB", "Shipment-NB"), Operator:=xlFilterValues
    rowCount = Range("F1").End(xlDown).Row
    Range("C2:C" & rowCount).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=R[1]C"
    
    ActiveSheet.AutoFilterMode = False
    Range("D1:D" & totalRow).Copy
    Range("C1:C" & totalRow).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    Range("G2:" & Chr(columnCount + 64) & totalRow).NumberFormatLocal = "#,##0_ ;[紅色]-#,##0 "
    Cells.EntireColumn.AutoFit
    Range("A1").Select
        
End Sub
