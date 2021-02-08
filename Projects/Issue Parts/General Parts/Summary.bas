Sub Generate_Summary()
' 更新Pivot Table並複製結果至Summary

    ' 刪除原Summary資料
    Sheets("Summary").Range("A1").CurrentRegion.Clear
    
    ' 更新Pivot Table, 並將Table貼至Summary
    With Worksheets("Pivot Table")
        .PivotTables("樞紐分析表1").PivotCache.Refresh
        .Range("A1:J" & .Range("F1").End(xlDown).row).Copy
        With Worksheets("Summary")
            .Range("A1").PasteSpecial xlPasteValues
            .Range("A1").PasteSpecial xlPasteFormats
        End With
        Application.CutCopyMode = False
    End With
    
    With Worksheets("Summary")
        .Range("A1").Copy
        .Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End With
    
End Sub
Sub Save_htm_File(file As String)
' 將Summary Table另存成htm檔案, 以供outlook郵件使用

    Dim table As Range
    With Worksheets("Summary")
        Set table = .Range("A1:J" & .Range("J1").End(xlDown).row)
    End With
    With ActiveWorkbook.PublishObjects.Add(xlSourceRange, file, "Summary", _
        table.Address, xlHtmlStatic)
        .Publish (False)
        .AutoRepublish = False
    End With

End Sub