Attribute VB_Name = "Module1"
' ==================== пр 5 VBA ====================
Sub zadanie_1_4()

Dim x As Single, eps As Single

x = InputBox("Введите X:")
eps = InputBox("Введите точность:")

y = my_arctg(x, eps)

MsgBox ("arctg(X) = " & y)

End Sub



' функция расчета арктангенса
Function my_arctg(x, eps)

Dim sum As Single   ' сумма разложения ряда функции
Dim a As Single     ' значение члена ряда
Dim n As Integer    ' номер члена ряда

a = x / 1
sum = a
n = 2

' цикл пока значение члена ряда превышает точность
Do While Abs(a) > eps
    a = -a * x * x / (2 * n - 1) * (2 * (n - 1) - 1)
    
    sum = sum + a
    n = n + 1
Loop

my_arctg = sum

End Function



