Attribute VB_Name = "Module1"
'create_journal_list - формирует итоговую таблицу для журнала вида
'№ | Фамилия И.О. | Подразделение
'в листе с именем "auto-list"

Sub create_journal_list()
    Set Rng = Application.InputBox("В колонке с фамилией выберите все ячейки, по которым сформировать таблицу", Type:=8)
    If Nothing Is Rng Then
        MsgBox "Что-нибудь нужно выбрать."
    Else
        row_top = Rng.Cells(1).row
        row_bottom = row_top + Rng.Cells.Count - 1
        Call system_create_journal_list(row_top, row_bottom)
    End If
End Sub

Sub system_create_journal_list(ByVal row_top, ByVal row_bottom)
    fam_col = "B"
    nam_col = "C"
    surnam_col = "D"
    otdel_col = "G"
    
    list_num_col = "A"
    list_fio_col = "B"
    list_otdel_col = "C"
    
    list_row = 1
    list_name = "список журнала"
    
    'чистим все данные в листе
    ActiveWorkbook.Sheets(list_name).Cells.ClearContents
    
    'пишем шапку в таблицу
    ActiveWorkbook.Sheets(list_name).Cells(list_row, list_num_col).value = "№"
    ActiveWorkbook.Sheets(list_name).Cells(list_row, list_fio_col).value = "Фамилия И.О."
    ActiveWorkbook.Sheets(list_name).Cells(list_row, list_otdel_col).value = "Подразделение"
    list_row = list_row + 1
    
    'заполняем данными
    For row = row_top To row_bottom
        fam = ActiveWorkbook.ActiveSheet.Cells(row, fam_col)
        nam = ActiveWorkbook.ActiveSheet.Cells(row, nam_col)
        surnam = ActiveWorkbook.ActiveSheet.Cells(row, surnam_col)
        otdel = ActiveWorkbook.ActiveSheet.Cells(row, otdel_col)
        fio = get_fio(fam, nam, surnam)
        If fio <> "" Then
            ActiveWorkbook.Sheets(list_name).Cells(list_row, list_num_col).value = list_row - 1
            ActiveWorkbook.Sheets(list_name).Cells(list_row, list_fio_col).value = fio
            ActiveWorkbook.Sheets(list_name).Cells(list_row, list_otdel_col).value = otdel
            list_row = list_row + 1
        End If
    Next row
    ActiveWorkbook.Sheets(list_name).Activate
End Sub


