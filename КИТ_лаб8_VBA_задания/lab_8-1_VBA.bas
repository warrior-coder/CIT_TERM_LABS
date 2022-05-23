Attribute VB_Name = "Module1"
' определяем структуру Отдел
Private Type Otdel
    nomer As Integer
    zarp_koef As Single
    rab_kolich As Integer
    zarp_sum As Single
End Type


' ==================== лаб8 VBA задание 2 ====================
Sub lab8_1()

' открытие файла
Set fso = CreateObject("Scripting.FileSystemObject")
ChDrive ("D")
ChDir ("\Documents\KURS-2\КИТ_2сем\КИТ_пособие\КИТ_лаб8_VBA_задания")
iflName = Application.GetOpenFilename()
Set ifl = fso.OpenTextFile(iflName, ForReading) ' файл для ввода

Set d = Range("A1").CurrentRegion
Dim line As String, newLine As String, otdely(1 To 10) As Otdel, tmp_otdel As Otdel

' считываем данные из файла
m = d.Rows.Count
k = 0
Do While Not ifl.AtEndOfStream
    line = ifl.ReadLine ' читаем строку из файла
    line = line + " "

    ' убираем двойные пробелы
    Do While InStr(line, "  ") <> 0
        line = Replace(line, "  ", " ")
    Loop
    
    ' считываем номера отделов и коэффициенты
    wrds = Split(line, " ") ' разбиваем строку на массив слов
    k = k + 1
    otdely(k).nomer = CInt(wrds(0))
    otdely(k).zarp_koef = CSng(wrds(1))
    
    ' вычисляем новые зарплаты и зааолняем данные по отделам
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

' сортируем отделы по номеру
For i = 1 To k - 1
    For j = i + 1 To k
        If (otdely(i).nomer > otdely(j).nomer) Then
            tmp_otdel = otdely(i)
            otdely(i) = otdely(j)
            otdely(j) = tmp_otdel
        End If
    Next j
Next i

' выводим данные об отделах
Set ofl = fso.OpenTextFile("D:\Documents\KURS-2\КИТ_2сем\КИТ_пособие\КИТ_лаб8_VBA_задания\otdely.txt", ForWriting, True) ' файл для вывода
newLine = "ном.  кол.  сум."
ofl.WriteLine (newLine)
For i = 1 To k
    newLine = CStr(otdely(i).nomer) + "      " + CStr(otdely(i).rab_kolich) + "     " + CStr(otdely(i).zarp_sum)
    ofl.WriteLine (newLine)
Next i
ofl.Close

End Sub

