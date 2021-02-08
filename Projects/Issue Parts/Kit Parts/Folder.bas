Sub Create_Folder()
' 自動產生Issue Part的資料夾, 依照資料下載日期和群別分類

    Dim folder_y, folder_m As String
    folder_y = "D:\Users\lindac\Documents\Issue Part\" & Format(Date, "YYYYMM")

    If Dir(folder_y, vbDirectory) = "" Then
        MkDir folder_y
    End If
        
    folder_m = folder_y & "\" & Format(Date, "MMDD")
    If Dir(folder_m, vbDirectory) = "" Then
        MkDir folder_m
    
        Dim folder_1, folder_2, folder_3, folder_4, folder_5, folder_6, _
            folder_7, folder_8, folder_sc, folder_nc As String
        
        folder_1 = folder_m & "\PG1"
        folder_2 = folder_m & "\PG2"
        folder_3 = folder_m & "\PG3"
        folder_4 = folder_m & "\PG4"
        folder_5 = folder_m & "\PG5"
        folder_6 = folder_m & "\PG6"
        folder_7 = folder_m & "\PG7"
        folder_8 = folder_m & "\PG8"
        folder_sc = folder_m & "\SC"
        folder_nc = folder_m & "\NC"
        
        MkDir folder_1:  MkDir folder_2: MkDir folder_3: MkDir folder_4: _
            MkDir folder_5: MkDir folder_6: MkDir folder_7: MkDir folder_8: _
            MkDir folder_sc: MkDir folder_nc:
    End If
    
End Sub