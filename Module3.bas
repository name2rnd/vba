Attribute VB_Name = "Module3"
'function fio - получает на вход фамилию, имя, отчество.
'если все три параметра заданы, то формирует строку Фамилия И.О.
Function get_fio(f, m, s)
    If f <> "" And m <> "" And s <> "" Then
        get_fio = f & " " & Left(m, 1) & "." & Left(s, 1) & "."
    Else
        get_fio = ""
    End If
End Function

Function select_range()
    Set Rng = Application.InputBox("Выберите ячейки", Type:=8)
    If Nothing Is Rng Then
        MsgBox "Что-нибудь нужно выбрать."
    Else
        Dim selected(1 To 2) As Integer
        selected(1) = Rng.Cells(1).row
        selected(2) = Rng.Cells(1).row + Rng.Cells.Count - 1
        select_range = selected
    End If
End Function
