Attribute VB_Name = "Module1"
' ==================== �� 5 VBA ====================
Sub zadanie_1_4()

Dim x As Single, eps As Single

x = InputBox("������� X:")
eps = InputBox("������� ��������:")

y = my_arctg(x, eps)

MsgBox ("arctg(X) = " & y)

End Sub



' ������� ������� �����������
Function my_arctg(x, eps)

Dim sum As Single   ' ����� ���������� ���� �������
Dim a As Single     ' �������� ����� ����
Dim n As Integer    ' ����� ����� ����

a = x / 1
sum = a
n = 2

' ���� ���� �������� ����� ���� ��������� ��������
Do While Abs(a) > eps
    a = -a * x * x / (2 * n - 1) * (2 * (n - 1) - 1)
    
    sum = sum + a
    n = n + 1
Loop

my_arctg = sum

End Function



