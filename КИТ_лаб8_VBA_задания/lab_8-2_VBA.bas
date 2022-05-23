Attribute VB_Name = "Module1"
' ==================== лаб8 VBA задание 2 ====================
Sub lab8_2()

Set fso = CreateObject("Scripting.FileSystemObject")
Set ifl = fso.OpenTextFile("\Documents\KURS-2\КИТ_2сем\КИТ_пособие\КИТ_лаб8_VBA_задания\input.txt", ForReading) ' файл для ввода
Set ofl = fso.OpenTextFile("\Documents\KURS-2\КИТ_2сем\КИТ_пособие\КИТ_лаб8_VBA_задания\output.txt", ForWriting, True) ' файл для вывода
Dim newLine As String, line As String

Do While Not ifl.AtEndOfStream
    newLine = ""
    line = ifl.ReadLine ' читаем строку из файла
    line = line + " "
        
    ' убираем двойные пробелы
    Do While InStr(line, "  ") <> 0
        line = Replace(line, "  ", " ")
    Loop
    
    wrds = Split(line, " ") ' разбиваем строку на массив слов
    m = UBound(wrds) ' получаем индекс последнего элекемнта массива
    
    ' выводим слова в строку в обратном порядке
    For i = m - 1 To 0 Step -1
        newLine = newLine + wrds(i) + " "
    Next i
    ofl.WriteLine (newLine)
Loop

ifl.Close

End Sub
