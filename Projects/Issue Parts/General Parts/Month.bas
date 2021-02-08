Sub Update_Month()
' 更新Backlog Issue和Shortage Issue的月份

    With Worksheets("Raw Data")
        If .Range("H1").Value <> Sheets("Backlog Issue").Range("L2").Value Then
            .Range("H1").Copy
            Sheets("Backlog Issue").Range("L2:O2").PasteSpecial xlPasteValues
            Sheets("Shortage Issue").Range("L2:O2").PasteSpecial xlPasteValues
            .Range("R1").Copy
            Sheets("Backlog Issue").Range("P2:Q2").PasteSpecial xlPasteValues
            Sheets("Shortage Issue").Range("P2:Q2").PasteSpecial xlPasteValues
            .Range("AB1").Copy
            Sheets("Backlog Issue").Range("R2:S2").PasteSpecial xlPasteValues
            Sheets("Shortage Issue").Range("R2:S2").PasteSpecial xlPasteValues
            .Range("AL1").Copy
            Sheets("Backlog Issue").Range("T2:U2").PasteSpecial xlPasteValues
            Sheets("Shortage Issue").Range("T2:U2").PasteSpecial xlPasteValues
            .Range("AV1").Copy
            Sheets("Backlog Issue").Range("V2:W2").PasteSpecial xlPasteValues
            Sheets("Shortage Issue").Range("V2:W2").PasteSpecial xlPasteValues
            .Range("BF1").Copy
            Sheets("Backlog Issue").Range("X2:Y2").PasteSpecial xlPasteValues
            Sheets("Shortage Issue").Range("X2:Y2").PasteSpecial xlPasteValues
            Application.CutCopyMode = False
        End If
    End With

End Sub