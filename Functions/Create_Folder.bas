Sub Create_Folder()

' Create folder to save new excel workbook

    Dim folder_y, folder_m As String
    folder_y = "D:\Users\lindac\Desktop\" & Format(Date, "YYYYMM")

    ' Check if folder already create, if not then create folder
    If Dir(folder_y, vbDirectory) = "" Then
        MkDir folder_y
    End If
        
    folder_m = folder_y & "\" & Format(Date, "MMDD")
    If Dir(folder_m, vbDirectory) = "" Then
        MkDir folder_m
    End If
    
End Sub
