Attribute VB_Name = "Module1"
' ==================== ���8 VBA ������� 2 ====================
Sub lab8_2()

Set fso = CreateObject("Scripting.FileSystemObject")
Set ifl = fso.OpenTextFile("\Documents\KURS-2\���_2���\���_�������\���_���8_VBA_�������\input.txt", ForReading) ' ���� ��� �����
Set ofl = fso.OpenTextFile("\Documents\KURS-2\���_2���\���_�������\���_���8_VBA_�������\output.txt", ForWriting, True) ' ���� ��� ������
Dim newLine As String, line As String

Do While Not ifl.AtEndOfStream
    newLine = ""
    line = ifl.ReadLine ' ������ ������ �� �����
    line = line + " "
        
    ' ������� ������� �������
    Do While InStr(line, "  ") <> 0
        line = Replace(line, "  ", " ")
    Loop
    
    wrds = Split(line, " ") ' ��������� ������ �� ������ ����
    m = UBound(wrds) ' �������� ������ ���������� ��������� �������
    
    ' ������� ����� � ������ � �������� �������
    For i = m - 1 To 0 Step -1
        newLine = newLine + wrds(i) + " "
    Next i
    ofl.WriteLine (newLine)
Loop

ifl.Close

End Sub
