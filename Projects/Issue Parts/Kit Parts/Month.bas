Sub Update_Month()
' 更新Backlog Issue和Shortage Issue的月份

    With Worksheets("Raw Data")
        If .Range("H1").Value <> Sheets("Backlog Issue").Range("M2").Value Then
            .Range("H1").Copy
            Sheets("Backlog Issue").Range("M2:P2").PasteSpecial xlPasteValues
            Sheets("Shortage Issue").Range("M2:P2").PasteSpecial xlPasteValues
            .Range("P1").Copy
            Sheets("Backlog Issue").Range("Q2:R2").PasteSpecial xlPasteValues
            Sheets("Shortage Issue").Range("Q2:R2").PasteSpecial xlPasteValues
            .Range("X1").Copy
            Sheets("Backlog Issue").Range("S2:T2").PasteSpecial xlPasteValues
            Sheets("Shortage Issue").Range("S2:T2").PasteSpecial xlPasteValues
            .Range("AF1").Copy
            Sheets("Backlog Issue").Range("U2:V2").PasteSpecial xlPasteValues
            Sheets("Shortage Issue").Range("U2:V2").PasteSpecial xlPasteValues
            .Range("AN1").Copy
            Sheets("Backlog Issue").Range("W2:X2").PasteSpecial xlPasteValues
            Sheets("Shortage Issue").Range("W2:X2").PasteSpecial xlPasteValues
            .Range("AV1").Copy
            Sheets("Backlog Issue").Range("Y2:Z2").PasteSpecial xlPasteValues
            Sheets("Shortage Issue").Range("Y2:Z2").PasteSpecial xlPasteValues
            Application.CutCopyMode = False
        End If
    End With

End Sub