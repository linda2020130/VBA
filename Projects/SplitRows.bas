Sub SplitRows()
' 拆分資料列

    Application.ScreenUpdating = False
    With Worksheets("總表")
        Dim i, j, totalRow As Integer
        totalRow = .Range("B1").End(xlDown).Row
        j = 2
        .Range("O2").FormulaR1C1 = "=VLOOKUP(RC[-13],'Mapping table'!C8:C14,7,FALSE)"
        .Range("O2").AutoFill Destination:=.Range("O2:O" & totalRow), Type:=xlFillDefault
        For i = 2 To totalRow
            Dim qty, mpq As Double
            qty = .Range("C" & i).Value
            mpq = .Range("O" & i).Value
            Do While qty > 0
                .Range("A" & i & ":N" & i).Copy Sheets("拆分表").Range("A" & j)
                If qty > mpq Then
                    Sheets("拆分表").Range("C" & j).Value = mpq
                Else
                    Sheets("拆分表").Range("C" & j).Value = qty
                End If
                qty = qty - mpq
                j = j + 1
            Loop
        Next
        .Columns("O").Delete
    End With
    
    With Worksheets("拆分表")
        totalRow = .Range("B1").End(xlDown).Row
        .Range("D2").FormulaR1C1 = "=VLOOKUP(RC[-2],'Mapping table'!C8:C11,4,FALSE)"
        .Range("D2").AutoFill Destination:=.Range("D2:D" & totalRow), Type:=xlFillDefault
        .Range("L2").FormulaR1C1 = "=VLOOKUP(RC[-10],'Mapping table'!C8:C17,10,FALSE)"
        .Range("L2").AutoFill Destination:=.Range("L2:L" & totalRow), Type:=xlFillDefault
        .Range("H2").FormulaR1C1 = "=TEXT(MONTH(RC[-2]),""mmm"")&""'""&RIGHT(YEAR(RC[-2]),2)"
        .Range("H2").AutoFill Destination:=.Range("H2:H" & totalRow), Type:=xlFillDefault
    End With
    
    Application.ScreenUpdating = True
    
End Sub