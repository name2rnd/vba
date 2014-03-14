Attribute VB_Name = "normativs"
'sex - ��� (�/�)
'pullups - ������������
'pushups - ���������
'situps - �����
'gym - ��� ����
'ten_to_ten - ��������� ��� 10�10
'four_to_twenty - ��������� ��� 4�20

'Function count_normativ_strength - ������� ��� ������� ��������� ����.
'�� ���� ���� �������: ���, ������������, ���������, �����, ��� ����

'Function count_normativ_speed - ������� ��� ������� ��������� �������� � ��������
'�� ���� ���� �������: ���, ��������� ��� 10�10, ��������� ��� 4�20

'������� count_normativ_strength � count_normativ_strength �� �������� ������� ��� ���������� ����,
'� �������� ����� ����� ��������� � ����������� ���� �� �������� ������� �� ���������

Function count_normativ_strength(sex, Optional pullups, Optional pushups, Optional situps, Optional gym)
    sheetname = "" '��� �����, � �������� ����� ������� ���������
    pullups_col = "" '�������, ��� ��������� ��� ������������
    pushups_col = "" '�������, ��� ��������� ��� ���������
    situps_col = "" '�������, ��� ��������� ��� ������
    gym_col = "" '�������, ��� ��������� ��� ���� ����
    row_top = 0 '������ ������� � �����������
    row_bottom = 0 '����� ������� � �����������
    balls_col = "A" '�������, ��� �������� ���� �� ����������
    '���� ������� ���
    If sex = "�" Then
        sheetname = "���������-�������"
        pullups_col = "B"
        pushups_col = "C"
        gym_col = "D"
        row_top = 9
        row_bottom = 109
    '���� ������� ���
    ElseIf sex = "�" Then
        sheetname = "���������-�������"
        pushups_col = "B"
        situps_col = "C"
        row_top = 8
        row_bottom = 108
    End If
    result = 0
    If sheetname <> "" Then
        '����� �� ������������
        pullups_ball = search_desc(sheetname, pullups_col, balls_col, row_bottom, row_top, pullups)
        '����� �� ���������
        pushups_ball = search_desc(sheetname, pushups_col, balls_col, row_bottom, row_top, pushups)
        '������ �� �����
        situps_ball = search_desc(sheetname, situps_col, balls_col, row_bottom, row_top, situps)
        '����� �� ��� ����
        gym_ball = search_desc(sheetname, gym_col, balls_col, row_bottom, row_top, gym)
        
        result = pullups_ball + pushups_ball + situps_ball + gym_ball
    End If
    count_normativ_strength = result
End Function

Function count_normativ_speed(sex, Optional ten_to_ten, Optional four_to_twenty)
    sheetname = "" '��� �����, � �������� ����� ������� ���������
    ten_to_ten_col = "" '�������, ��� ��������� 10�10
    four_to_twenty_col = "" '�������, ��� ��������� 4�20
    row_top = 0 '������ ������� � �����������
    row_bottom = 0 '����� ������� � �����������
    balls_col = "A" '�������, ��� �������� ���� �� ����������
    '���� ������� ���
    If sex = "�" Then
        sheetname = "���������-�������"
        ten_to_ten_col = "E"
        four_to_twenty_col = "F"
        row_top = 9
        row_bottom = 109
    '���� ������� ���
    ElseIf sex = "�" Then
        sheetname = "���������-�������"
        ten_to_ten_col = "D"
        row_top = 8
        row_bottom = 108
    End If
    result = 0
    If sheetname <> "" Then
        '����� �� ��������� ��� 10�1
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
        If sex = "�" Then
            sheetname = "���������-�������"
            row_top = 6
            row_bottom = 13
            age_col = "K"
            norm_col = "N"
        ElseIf sex = "�" Then
            sheetname = "���������-�������"
            row_top = 6
            row_bottom = 11
            age_col = "H"
            norm_col = "K"
        End If
        
        norm_ball = search_asc(sheetname, age_col, norm_col, row_top, row_bottom, age)
        If result < norm_ball Then
            count_normativ_result = "����"
        Else
            count_normativ_result = "��"
        End If
    Else
        count_normativ_result = ""
    End If
End Function


