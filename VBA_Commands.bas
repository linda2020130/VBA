Attribute VB_Name = "Basic Commands"
Sub Basic_Commands()
Attribute Basic_Commands.VB_Description = ""
Attribute Basic_Commands.VB_ProcData.VB_Invoke_Func = " \n14"
'
    ' 抓資料末欄/列等相關語法
    totalRow = Sheets("Sheet1").Range("A1").End(xlDown).Row
    totalColumn = Sheets("Sheet1").Range("A1").End(xlToRight).Column
    ' 數字轉字母(e.g. 1->A, 2->B...)
    Chr(totalColumn + 64)

    ' 刪除相關語法
    ' Delete row 2
    Sheets("Sheet1").rows(2).Delete
    ' Delete column A
    Sheets("Sheet1").columns("A").Delete
    ' Delete data from row 2 to bottom of column A in Sheet1
    Sheets("Sheet1").rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ' Delete data from row 2 to bottom of column A in Sheet1
    ' can substitute column A for other column
    Sheets("Sheet1").rows("2:" & Range("A2").End(xlDown).row).Delete


    ' 清除相關語法
    ' Clear content in current region(以指定儲存格為基準, 直到碰到空白行列)
    Sheets("Sheet1").Range("A1").CurrentRegion.Clear


    ' 複製貼上相關語法
    ' Copy data from certain columns to bottom of column B in Sheet1
    ' Copy all
    Sheets("Sheet1").Range("B2:E" & Range("B2").End(xlDown).row).Copy
    ' Copy visible cells only
    Sheets("Sheet1").Range("B2:E" & Range("B2").End(xlDown).row).SpecialCells(xlCellTypeVisible).Copy
    ' Paste all
    Sheets("Sheet2").Range("A1").Paste
    Application.CutCopyMode = False
    ' Paste value only 
    Sheets("Sheet2").Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    ' Paste format only
    Sheets("Sheet2").Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    ' Paste formula only
    Sheets("Sheet2").Range("A1").PasteSpecial xlPasteFormulas
    Application.CutCopyMode = False
    ' Copy and paste in one line (each line work the same)
    Sheets("Sheet1").Range("B2:E2").Copy Sheets("Sheet2").Range("B1:E1")
    Sheets("Sheet1").Range("B2:E2").Copy Sheets("Sheet2").Range("B1")
    

    ' 修改儲存格內公式
    ' Change formula of certain cell in Sheet1
    Sheets("Sheet1").Range("F1:G1").FormulaR1C1 = "=INDEX(顯示值所在陣列,查詢索引值,顯示值所在第?個欄位)"
    ' Change formula of certain cell - formula=IF(sum($A1:$E1)=0,0,1) (鎖住column)
    Sheets("Sheet1").Range("F1").FormulaR1C1 = "=IF(SUM(RC1:RC5)=0,0,1)"
    ' Range("J9") = Range("B9") => row keeps unchanged; column shifts from J to B (-8)
    Sheets("Sheet1").Range("J9").FormulaR1C1 = "=RC[-8]"
    ' Range("K9") = Range("J13") => row shifts from 9 to 13 (+4); column shifts from K to J (-1)
    Sheets("Sheet1").Range("K9").FormulaR1C1 = "=R[4]C[-1]"
    ' 只修改篩選出來的資料列
    Sheets("Sheet1").Range("A1:D100").SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=R[1]C"


    ' 修改名稱管理員內的公式
    ' 名稱範圍會隨著資料的增減改變
    ' Change formula of "Column A" in Name Manager (公式 -> 名稱管理員) to catch different numbers of row if data changed
    ' "Column A" will be used to make a drop down list without header(row1)
    ActiveWorkbook.Names("Column A").RefersToR1C1 = "=OFFSET(Sheet2!R2C1,,,COUNTA(Sheet2!C1)-1)"


    ' 填滿相關語法
    ' 已知填滿目的地
    ' Autofill formula of Range("A1:F1") with range("A1")
    Sheets("Sheet1").Range("A1").Autofill Destination:=Range("A1:F1"), Type:=xlFillDefault
    ' Autofill formula of Range("A1:D2") with range("A1:A2")
    Sheets("Sheet1").Range("A1:A2").Autofill Destination:=Range("A1:D2"), Type:=xlFillDefault
    ' 自動填滿至特定欄的資料末端
    ' Autofill columns a to c with Range("A2:C2")
    Sheets("Sheet1").Range("A2:C2").Autofill Destination:=Range("A2:C" & Range("C2").End(xlDown).row)


    ' 篩選相關語法
    ' Turn on autoFilter at certain range
    Sheets("Sheet1").Range("A5:L5").AutoFilter
    ' Turn off autofilter
    Sheets("Sheet1").AutoFilterMode = False
    ' Filter out column C (3rd column) with value = 1 (Only show value with 1)
    Sheets("Sheet1").Range("A5:L" & Range("L5").End(xlDown).row).AutoFilter Field:=3, Criteria1:="1"
    ' Filter out column C (3rd column) with value not equals to 1 (Only show values that is not 1)
    Sheets("Sheet1").Range("A5:L" & Range("L5").End(xlDown).row).AutoFilter Field:=3, Criteria1:="<>1"
    ' Filter out column F (6th column) with value = A or B or C
    ActiveSheet.Range("A1:" & Chr(totalColumn + 64) & totalRow).AutoFilter Field:=6, Criteria1:=Array( _
        "A", "B", "C"), Operator:=xlFilterValues


    ' 排序相關語法
    ' Sort by column A then B then C (priority: C>B>A) on column A to G including header on row 1
    Dim i as integer
    i = Sheets("Sheet1").Cells(rows.Count, "G").End(xlUp).row
    Sheets("Sheet1").Range("A1:G" & i).Sort key1:=Range("C1"), order1:=xlDescending, _
        key2:=Range("B1"), order2:=xlAscending, key3:=Range("A1"), order3:=xlAscending, Header:=xlYes

    
    ' 尋找相關用法
    ' Find cells with string "Pig" in column B
    ' Do something if rng is not empty(meaning there is at least one cell with "Pig")
    Dim rng As Range
    Set rng = Sheets("Sheet1").Range("B:B").Find("Pig", lookat:=xlPart)
    If Not rng Is Nothing Then
    ...
    End If
    ' Find cells equal to string "Pig" in column B
    ' Do something if rng is (Not) empty(meaning there is no cell equals to "Pig")
    Set rng = Sheets("Sheet1").Range("B:B").Find("Pig", lookat:=xlWhole)
    If (Not) rng Is Nothing Then
    ...
    End If


    ' 字串相關語法
    ' 搜尋字串位置
    ' Find position of substring in string
    ' 預設大小寫視同不同, 增加比對方式參數來將字母大小寫視為相同
    Dim pos As Integer
    pos = InStr("Hello, World.", "world", vbTextCompare)  # pos=8
    ' 擷取字串
    ' 從左邊擷取5個字
    Dim word as String 
    word = Left("Hello, world.", 5)  # word = Hello
    ' 從右邊擷取6個
    word = Right("Hello, world.", 6)  # word = world.
    ' 從第6個字元開始擷取2個
    word = Mid("This is a message.", 6, 2)  # word = is
    ' 從第6個字元開始擷取至最後
    word = Mid("This is a message.", 6)  # word = is a message.
    ' 移除空白
    ' Remove space at left side
    Dim mystr, newstr as String
    mystr = "  Hello!  "
    LTrim(mystr)
    ' Remove space at right side
    RTrim(mystr)
    ' Remove space from both sides
    mystr = Trim(mystr)
    ' 取代字串
    ' 起始位置1, 替換次數-1(代表不限制)
    newstr = Replace(mystr, "ello", "i", 1, -1, vbTextCompare)  # newstr = "  Hi!  "
    ' 移除特定字串
    ' Remove () at the end of string
    Dim pn as String
    pn = ActiveSheet.Range("A2").Value
    If InStr(pn, "(") > 0 Then
        ActiveSheet.Range("A2").Value = Mid(pn, 1, InStr(pn, "(") - 1)
    End If


    ' 取代相關語法
    ' Replace "$" with "*" in Column C
    Sheets("Sheet1").Range("C2:C" & .Range("C2").End(xlDown).row).Replace "$", "*"
    ' Remove "%" in column C
    Sheets("Sheet1").Range("C2:C" & .Range("C2").End(xlDown).row).Replace "%", ""


    ' 插入資料列相關語法
    ' Insert one row before row 5 (original row 5 will become row 6)
    Sheets("Sheet1").rows(5).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ' Insert three columns before column A (original column A will become column D)
    Sheets("Sheet1").Columns("A:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove


    ' 更新資料相關語法
    ' Refresh data source in Pivot Table
    Sheets("Pivot Table").PivotTables("樞紐分析表1").PivotCache.Refresh


    ' 借用Excel函數相關語法
    ' sum
    total = Application.WorksheetFunction.Sum(Range("A1:C1"))


    ' 格式相關語法
    ' Change font color
    With ActiveSheet.Range("A").Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
    ' Change background color to yellow
    ActiveSheet.Range("A1").Interior.Color = 65535
    ' Add double line at the bottom border
    ActiveSheet.Range("A1").Borders(xlEdgeBottom).LineStyle = xlDouble
    ' Mark red if value is less than 0
    ActiveSheet.Range("A2:" & Chr(totalColumn + 64) & totalRow).NumberFormatLocal = "#,##0_ ;[紅色]-#,##0 "
    ' Autofit columns
    ActiveSheet.Cells.EntireColumn.AutoFit


    ' 群組相關語法
    ' 打開群組
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
    ' 收起群組
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1


    ' 判斷式相關語法
    ' If.. elseif...else (可只有if判斷式)
    If ...  Then
    ElseIf ... Then
    Else
    End If


    ' 迴圈相關語法
    ' for迴圈
    For i = 10 to 1 Step -1
        ...
        If count >= 0 And count <> 5 Then
            Exit For
        End If
        ...
    Next
    ' do while迴圈(先檢查condition再跑迴圈內容)
    i = 1
    Do While i <= 10
        i = i + 1
    Loop
    ' do迴圈(先跑迴圈內容再檢查condition)
    i = 1
    Do 
        i = i + 1
    Loop While i <= 10


    ' 增加運行速度的方法
    ' 使用With...End With
    With Worksheets("Sheet3")
        .AutoFilterMode = False
        .Range("2:" & Range("A2").End(xlDown).row).Delete
    End With
    ' 運行前關閉螢幕更新 (運行後再打開)
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    
    
    ' 計時器相關語法
    ' Set timer
    t0 = Timer  ' put at the beginning
    t0 = Timer - t0    ' put at the end
    ' vbCrLf代表對話框換行
    MsgBox "Complete! " & vbCrLf & "Total running time: " & t0 \ 60 & "mins " & t0 Mod 60 & "secs."



    
End Sub