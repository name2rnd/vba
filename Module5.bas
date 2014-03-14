Attribute VB_Name = "Module5"
Sub create_attendance_list()
    selected = select_range
    row_top = selected(1)
    row_bottom = selected(2)
    
    Call system_create_attendance_list(row_top, row_bottom)
End Sub

Sub system_create_attendance_list(ByVal row_top, ByVal row_bottom)

    fam_col = "B"
    nam_col = "C"
    surnam_col = "D"
    otdel_col = "G"
    last_col = "O"
    
    list_num_col = "A"
    list_fio_col = "B"
    list_otdel_col = "C"
    
    'перед этим идет шапка
    list_row = 3
    list_name = "учет посещаемости"
    
    Set OurList = ActiveWorkbook.Sheets(list_name)
    Set SourceList = ActiveWorkbook.ActiveSheet
    
    counter = 0
    
    'заполняем данными
    For row = row_top To row_bottom
        
        With SourceList
            fam = .Cells(row, fam_col)
            nam = .Cells(row, nam_col)
            surnam = .Cells(row, surnam_col)
            otdel = .Cells(row, otdel_col)
        End With
        
        fio = get_fio(fam, nam, surnam)
        If fio <> "" Then
            counter = counter + 1
            With OurList
                'так как строки вставляются на ходу, то list_row не меняется
                .Rows(list_row).Insert
                
                'formats
                .Rows(list_row).RowHeight = 20
                .Range(.Cells(list_row, list_num_col), .Cells(list_row, list_fio_col)).HorizontalAlignment = xlHAlignLeft
                .Range(.Cells(list_row, list_num_col), .Cells(list_row, last_col)).Borders().LineStyle = xlContinuous
                .Cells(list_row, list_otdel_col).Font.Size = 10
                .Cells(list_row, list_otdel_col).Borders(xlEdgeRight).Weight = xlThick
                .Cells(list_row, "H").Borders(xlEdgeRight).Weight = xlThick
                .Cells(list_row, "N").Borders(xlEdgeRight).Weight = xlThick
                If counter = 1 Then
                    .Range(.Cells(list_row, "D"), .Cells(list_row, "N")).Borders(xlEdgeBottom).Weight = xlThick
                End If
                
                'пишем данные
                .Cells(list_row, list_num_col).value = counter
                .Cells(list_row, list_fio_col).value = fio
                .Cells(list_row, list_otdel_col).value = otdel
                
            End With
        End If
    Next row
    With OurList
        .Range(.Cells(list_row, "A"), .Cells(list_row + counter - 1, "O")).Sort Key1:=.Cells(3, "A")
        .Activate
    End With
    
End Sub
