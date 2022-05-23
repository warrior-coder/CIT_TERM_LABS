Attribute VB_Name = "Module1"
' ==================== лаб 5 VBA ====================

Sub zad_1_v4()

Set d = Range("A2").CurrentRegion
Dim kol As Integer, cena As Integer

m = d.Rows.Count
For i = 2 To m
    kol = d.Cells(i, 2).Value
    
    If kol <= 5 Then
        cena = 14
    ElseIf kol <= 10 Then
        cena = 12
    Else
        cena = 11
    End If
    
    d.Cells(i, 3).Value = cena
    d.Cells(i, 4).Value = cena * kol
Next i

End Sub



Sub zad_2_v4()

Set d = Selection
Dim imax As Byte, jmax As Byte, maxValue As Integer

' получаем размеры области
m = d.Rows.Count
n = d.Columns.Count

' поиск максимального элемента и его позиции
maxValue = d.Cells(1, 1).Value
For i = 1 To m
    For j = 1 To n
        If d.Cells(i, j).Value > maxValue Then
            maxValue = d.Cells(i, j).Value
            imax = i
            jmax = j
        End If
        
    Next j
Next i

MsgBox ("Позиция максимального элемента: [" & imax & ", " & jmax & "]")

' обмен 1 и jmax столбцов
For i = 1 To m
    tempValue = d.Cells(i, 1).Value
    d.Cells(i, 1).Value = d.Cells(i, jmax).Value
    d.Cells(i, jmax).Value = tempValue
Next i

End Sub


