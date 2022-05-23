Attribute VB_Name = "Module1"
' ==================== лаб 4 VBA ====================

Sub zadanie_1_4()

Dim a As Double, b As Double, s As Double

a = InputBox("Введите a: ")
b = InputBox("Введите b: ")
s = InputBox("Введите s: ")

For x = a To b Step s
    If x <= 5 Then
        y = (1 + x) / ((1 + x * x) ^ (1 / 3))
    ElseIf x < 7 Then
        y = -x + 2 * Exp(-2 * x)
    Else
        y = Abs(2 - x)
    End If
    
    MsgBox ("x = " & x & ", y = " & y)
Next x

End Sub



Sub zadanie_2_4()

Dim a(1 To 4, 1 To 3) As Integer, b(1 To 3) As Integer
Dim m As Byte, counter As Byte

' Ввод данных
For i = 1 To 4
    For j = 1 To 3
        a(i, j) = InputBox("a(" & i & "," & j & ") = ")
    Next j
Next i

For j = 1 To 3
    b(j) = InputBox("b(" & j & ") = ")
Next j

m = InputBox("Номер строки: ")

' Сравнение
counter = 0

For j = 1 To 3
    If a(m, j) > b(j) Then
        counter = counter + 1
    End If
Next j

' Вывод результата
MsgBox ("Результат: " & counter)

End Sub

