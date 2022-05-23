Attribute VB_Name = "Module1"
' ���������� ��������� �����
Private Type Otdel
    nomer As Integer
    zarp_koef As Single
    rab_kolich As Integer
    zarp_sum As Single
End Type


' ==================== ���8 VBA ������� 2 ====================
Sub lab8_1()

' �������� �����
Set fso = CreateObject("Scripting.FileSystemObject")
ChDrive ("D")
ChDir ("\Documents\KURS-2\���_2���\���_�������\���_���8_VBA_�������")
iflName = Application.GetOpenFilename()
Set ifl = fso.OpenTextFile(iflName, ForReading) ' ���� ��� �����

Set d = Range("A1").CurrentRegion
Dim line As String, newLine As String, otdely(1 To 10) As Otdel, tmp_otdel As Otdel

' ��������� ������ �� �����
m = d.Rows.Count
k = 0
Do While Not ifl.AtEndOfStream
    line = ifl.ReadLine ' ������ ������ �� �����
    line = line + " "

    ' ������� ������� �������
    Do While InStr(line, "  ") <> 0
        line = Replace(line, "  ", " ")
    Loop
    
    ' ��������� ������ ������� � ������������
    wrds = Split(line, " ") ' ��������� ������ �� ������ ����
    k = k + 1
    otdely(k).nomer = CInt(wrds(0))
    otdely(k).zarp_koef = CSng(wrds(1))
    
    ' ��������� ����� �������� � ��������� ������ �� �������
    otdely(k).rab_kolich = 0
    otdely(k).zarp_sum = 0
    For i = 1 To m
        If CInt(d.Cells(i, 2).Value) = otdely(k).nomer Then
            d.Cells(i, 3).Value = d.Cells(i, 3).Value * otdely(k).zarp_koef
            otdely(k).rab_kolich = otdely(k).rab_kolich + 1
            otdely(k).zarp_sum = otdely(k).zarp_sum + d.Cells(i, 3).Value
        End If
    Next i
Loop
ifl.Close

' ��������� ������ �� ������
For i = 1 To k - 1
    For j = i + 1 To k
        If (otdely(i).nomer > otdely(j).nomer) Then
            tmp_otdel = otdely(i)
            otdely(i) = otdely(j)
            otdely(j) = tmp_otdel
        End If
    Next j
Next i

' ������� ������ �� �������
Set ofl = fso.OpenTextFile("D:\Documents\KURS-2\���_2���\���_�������\���_���8_VBA_�������\otdely.txt", ForWriting, True) ' ���� ��� ������
newLine = "���.  ���.  ���."
ofl.WriteLine (newLine)
For i = 1 To k
    newLine = CStr(otdely(i).nomer) + "      " + CStr(otdely(i).rab_kolich) + "     " + CStr(otdely(i).zarp_sum)
    ofl.WriteLine (newLine)
Next i
ofl.Close

End Sub

