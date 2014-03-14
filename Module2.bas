Attribute VB_Name = "Module2"
Sub incoming_test_result()
    Set Rng = Application.InputBox("� ������� � �������� �������� ��� ������, �� ������� ������������ ���������", Type:=8)
    If Nothing Is Rng Then
        MsgBox "���-������ ����� �������."
    Else
        row_top = Rng.Cells(1).row
        row_bottom = row_top + Rng.Cells.Count - 1
        Call system_create_incoming_test_result(row_top, row_bottom)
    End If
End Sub

Sub system_create_incoming_test_result(ByVal row_top, ByVal row_bottom)
    '������, ������ ������
    '����� ������
    fam_col = "B"
    nam_col = "C"
    surnam_col = "D"
    otdel_col = "G"
    mvg_col = "F"
    
    '����������
    strengh_pullups_col = "X"
    strengh_pushups_col = "Y"
    strengh_situps_col = "Z"
    strengh_gym_col = "AA"
    strengh_ball_col = "AB"
    speed_10x10_col = "AC"
    speed_4x20_col = "AD"
    speed_ball_col = "AE"
    result_ball_col = "AF"
    grade_col = "AG"
    
    '������, ���� �����
    '����� ������
    list_num_col = "A"
    list_fio_col = "B"
    list_otdel_col = "C"
    list_mvg_col = "F"
    
    '����������
    list_strengh_name_col = "G"
    list_strengh_value_col = "H"
    list_strengh_ball_col = "I"
    list_speed_name_col = "J"
    list_speed_value_col = "K"
    list_speed_ball_col = "L"
    list_ball_col = "M"
    list_grade_col = "N"
    list_total_grade_col = "P"
    
    '����� ���� ���� �����
    list_row = 6
    list_name = "������� ������������"
    
    Set OurList = ActiveWorkbook.Sheets(list_name)
    Set SourceList = ActiveWorkbook.ActiveSheet
    
    counter = 0
    
    '��������� �������
    For row = row_top To row_bottom
        
        '���������� ����� ����
        strengh_res_name = ""
        strengh_res_value = ""
        strenght_res_ball = ""
        
        '���������� ����� ��������
        speed_res_name = ""
        speed_res_value = ""
        speed_res_ball = ""
    
        result_ball = ""
        grade = ""
        
        With SourceList
        
            fam = .Cells(row, fam_col)
            nam = .Cells(row, nam_col)
            surnam = .Cells(row, surnam_col)
            mvg = .Cells(row, mvg_col)
            
            '������� ����������
            If (.Cells(row, strengh_pullups_col)) Then
                strengh_res_name = "������"
                strengh_res_value = .Cells(row, strengh_pullups_col)
            ElseIf (.Cells(row, strengh_pushups_col)) Then
                strengh_res_name = "�����"
                strengh_res_value = .Cells(row, strengh_pushups_col)
            ElseIf (.Cells(row, strengh_situps_col)) Then
                strengh_res_name = "�����"
                strengh_res_value = .Cells(row, strengh_situps_col)
            ElseIf (.Cells(row, strengh_gym_col)) Then
                strengh_res_name = "��� ����"
                strengh_res_value = .Cells(row, strengh_gym_col)
            End If
            strenght_res_ball = .Cells(row, strengh_ball_col)
            
            If (.Cells(row, speed_10x10_col)) Then
                speed_res_name = "10x10"
                speed_res_value = .Cells(row, speed_10x10_col)
            ElseIf (.Cells(row, speed_4x20_col)) Then
                speed_res_name = "4x20"
                speed_res_value = .Cells(row, speed_4x20_col)
            End If
            speed_res_ball = .Cells(row, speed_ball_col)
            
            result_ball = .Cells(row, result_ball_col)
            grade = .Cells(row, grade_col)
            
        End With
        fio = get_fio(fam, nam, surnam)
        If fio <> "" Then
            counter = counter + 1
            With OurList
                '��� ��� ������ ����������� �� ����, �� list_row �� ��������
                .Rows(list_row).Insert
                
                'formats
                .Rows(list_row).RowHeight = 15
                .Range(.Cells(list_row, list_num_col), .Cells(list_row, list_total_grade_col)).Font.Size = 12
                .Range(.Cells(list_row, list_num_col), .Cells(list_row, list_total_grade_col)).Font.Bold = False
                .Range(.Cells(list_row, list_num_col), .Cells(list_row, list_total_grade_col)).Orientation = xlHorizontal
                .Range(.Cells(list_row, list_num_col), .Cells(list_row, list_fio_col)).HorizontalAlignment = xlHAlignLeft
                .Cells(list_row, list_total_grade_col).HorizontalAlignment = xlHAlignCenter
                .Range(.Cells(list_row, list_num_col), .Cells(list_row, list_total_grade_col)).Borders().LineStyle = xlContinuous
                
                
                '����� ������
                .Cells(list_row, list_num_col).value = counter
                .Cells(list_row, list_fio_col).value = fio
                .Cells(list_row, list_mvg_col).value = mvg
                
                '���������� ����� ����
                .Cells(list_row, list_strengh_name_col).value = strengh_res_name
                .Cells(list_row, list_strengh_value_col).value = strengh_res_value
                .Cells(list_row, list_strengh_ball_col).value = strenght_res_ball
                
                '���������� ����� ��������
                .Cells(list_row, list_speed_name_col).value = speed_res_name
                .Cells(list_row, list_speed_value_col).value = speed_res_value
                .Cells(list_row, list_speed_ball_col).value = speed_res_ball
                
                .Cells(list_row, list_ball_col).value = result_ball
                .Cells(list_row, list_grade_col).value = grade
                
                formula_res = "=IF(AND(C6<>" & Chr(34) & Chr(34) & ", D6<>" & Chr(34) & Chr(34) & ",E6<>" & Chr(34) & Chr(34) & ",N6<>" & Chr(34) & Chr(34) & ", O6<>" & Chr(34) & Chr(34) & "), IF(AND(C6=" & Chr(34) & "��" & Chr(34) & ",D6=" & Chr(34) & "��" & Chr(34) & ",E6=" & Chr(34) & "��" & Chr(34) & ",N6=" & Chr(34) & "��" & Chr(34) & ", O6=" & Chr(34) & "��" & Chr(34) & "), " & Chr(34) & "��" & Chr(34) & ", " & Chr(34) & "����" & Chr(34) & "), " & Chr(34) & "-" & Chr(34) & ")"
                .Cells(list_row, list_total_grade_col).Formula = formula_res
                    
            End With
        End If
    Next row
    
    With OurList
        .Range(.Cells(list_row, "A"), .Cells(list_row + counter - 1, "P")).Sort Key1:=.Cells(3, "A")
        .Activate
    End With
End Sub

