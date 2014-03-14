Attribute VB_Name = "normativs"
'sex - пол (м/ж)
'pullups - подт€гивани€
'pushups - отжимани€
'situps - пресс
'gym - жим гири
'ten_to_ten - челночный бег 10х10
'four_to_twenty - челночный бег 4х20

'Function count_normativ_strength - формула дл€ расчета норматива —ила.
'Ќа вход надо колонки: пол, подт€гивани€, отжимани€, пресс, жим гири

'Function count_normativ_speed - формула дл€ расчета норматива Ѕыстрота и Ћовкость
'Ќа вход надо колонки: пол, челночный бег 10х10, челночный бег 4х20

'‘ормулы count_normativ_strength и count_normativ_strength по значению колонки ѕол определ€ют лист,
'с которого нужно брать нормативы и расчитывают балл по значению каждого из измерений

Function count_normativ_strength(sex, Optional pullups, Optional pushups, Optional situps, Optional gym)
    sheetname = "" 'им€ листа, с которого будут братьс€ нормативы
    pullups_col = "" 'колонка, где нормативы дл€ подт€гиваний
    pushups_col = "" 'колонка, где нормативы дл€ отжиманий
    situps_col = "" 'колонка, где нормативы дл€ пресса
    gym_col = "" 'колонка, где нормативы дл€ жима гири
    row_top = 0 'начало таблицы с нормативами
    row_bottom = 0 'конец таблицы с нормативами
    balls_col = "A" 'колонка, где итоговый балл за упражнение
    'если мужской пол
    If sex = "м" Then
        sheetname = "нормативы-мужчины"
        pullups_col = "B"
        pushups_col = "C"
        gym_col = "D"
        row_top = 9
        row_bottom = 109
    'если женский пол
    ElseIf sex = "ж" Then
        sheetname = "нормативы-женщины"
        pushups_col = "B"
        situps_col = "C"
        row_top = 8
        row_bottom = 108
    End If
    result = 0
    If sheetname <> "" Then
        'баллы за потт€гивани€
        pullups_ball = search_desc(sheetname, pullups_col, balls_col, row_bottom, row_top, pullups)
        'баллы за отжимани€
        pushups_ball = search_desc(sheetname, pushups_col, balls_col, row_bottom, row_top, pushups)
        'баллый за пресс
        situps_ball = search_desc(sheetname, situps_col, balls_col, row_bottom, row_top, situps)
        'баллы за жим гири
        gym_ball = search_desc(sheetname, gym_col, balls_col, row_bottom, row_top, gym)
        
        result = pullups_ball + pushups_ball + situps_ball + gym_ball
    End If
    count_normativ_strength = result
End Function

Function count_normativ_speed(sex, Optional ten_to_ten, Optional four_to_twenty)
    sheetname = "" 'им€ листа, с которого будут братьс€ нормативы
    ten_to_ten_col = "" 'колонка, где нормативы 10х10
    four_to_twenty_col = "" 'колонка, где нормативы 4х20
    row_top = 0 'начало таблицы с нормативами
    row_bottom = 0 'конец таблицы с нормативами
    balls_col = "A" 'колонка, где итоговый балл за упражнение
    'если мужской пол
    If sex = "м" Then
        sheetname = "нормативы-мужчины"
        ten_to_ten_col = "E"
        four_to_twenty_col = "F"
        row_top = 9
        row_bottom = 109
    'если женский пол
    ElseIf sex = "ж" Then
        sheetname = "нормативы-женщины"
        ten_to_ten_col = "D"
        row_top = 8
        row_bottom = 108
    End If
    result = 0
    If sheetname <> "" Then
        'баллы за челночный бег 10х1
        ten_to_ten_ball = search_asc(sheetname, ten_to_ten_col, balls_col, row_top, row_bottom, ten_to_ten)
        four_to_twenty_ball = search_asc(sheetname, four_to_twenty_col, balls_col, row_top, row_bottom, four_to_twenty)
        result = ten_to_ten_ball + four_to_twenty_ball
    End If
    count_normativ_speed = result
End Function

Private Function search_desc(ByVal sheetname, ByVal cell_name, ByVal result_cell_name, ByVal row_from, ByVal row_to, ByVal value)
    result = 0
    If value > 0 And cell_name <> "" Then
        Do While row_from >= row_to
            norm_value = ActiveWorkbook.Sheets(sheetname).Cells(row_from, cell_name)
            If norm_value <> "-" Then
                If value >= norm_value Then
                    result = ActiveWorkbook.Sheets(sheetname).Cells(row_from, result_cell_name)
                Else
                    row_from = row_to 'break loop
                End If
            End If
            row_from = row_from - 1
        Loop
    End If
    search_desc = result
End Function


Private Function search_asc(ByVal sheetname, ByVal cell_name, ByVal result_cell_name, ByVal row_from, ByVal row_to, ByVal value)
    result = 0
    If value > 0 And cell_name <> "" Then
        Do While row_from <= row_to
            norm_value = ActiveWorkbook.Sheets(sheetname).Cells(row_from, cell_name)
            If norm_value <> "-" Then
                result = ActiveWorkbook.Sheets(sheetname).Cells(row_from, result_cell_name)
                If value <= norm_value Then
                    row_from = row_to 'break loop
                End If
            End If
            row_from = row_from + 1
        Loop
    End If
    search_asc = result
End Function

Function count_normativ_result(sex, age, result)
    If result > 0 Then
        sheetname = ""
        row_top = 0
        row_bottom = 0
        age_col = 0
        norm_col = 0
        If sex = "м" Then
            sheetname = "нормативы-мужчины"
            row_top = 6
            row_bottom = 13
            age_col = "K"
            norm_col = "N"
        ElseIf sex = "ж" Then
            sheetname = "нормативы-женщины"
            row_top = 6
            row_bottom = 11
            age_col = "H"
            norm_col = "K"
        End If
        
        norm_ball = search_asc(sheetname, age_col, norm_col, row_top, row_bottom, age)
        If result < norm_ball Then
            count_normativ_result = "неуд"
        Else
            count_normativ_result = "уд"
        End If
    Else
        count_normativ_result = ""
    End If
End Function


