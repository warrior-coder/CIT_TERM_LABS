VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ==================== пр7 VBA ====================
Private Sub CommandButton1_Click()

Set Result = CreateObject("word.application") 'Создаем объект в MS Word

'Добавление документа
Result.Visible = True
Result.Documents.Add

' выводим заголовок в Word
With Result.Selection
    .Font.Bold = True
    .Font.Size = 16
    .ParagraphFormat.Alignment = wdAlignParagraphCenter
    .TypeText Text:="Лучшие студенты"
    
    .Font.Bold = False
    .Font.Size = 14
    .TypeParagraph
    .TypeParagraph
End With

' определяем диапазон данных в Excel
Set d = Worksheets("Лист1").Range("A1").CurrentRegion
m = d.Rows.Count
n = d.Columns.Count
Dim Sum As Integer, Sredniy As Single, mGreen As Byte

mGreen = 0
For i = 1 To m
    If d.Cells(i, 1).Font.Color = vbGreen Then mGreen = mGreen + 1
Next i

' добавляем таблицу в Word
Result.ActiveDocument.Tables.Add Range:=Result.Selection.Range, numrows:=mGreen, numcolumns:=2

' выводим данные из Excel в Word
With Result.Selection
    For i = 1 To m
        If d.Cells(i, 1).Font.Color = vbGreen Then
            ' подсчитываем средний бал
            Sum = 0
            For j = 2 To n
                Sum = Sum + d.Cells(i, j).Value
            Next j
            Sredniy = Sum / (n - 1)
            
            ' выводим фамилию и бал
            .TypeText Text:=d.Cells(i, 1).Value
            .MoveRight unit:=wdCell ' переход в следующую ячейку таблицы Word
            .TypeText Text:=CStr(Sredniy)
            .MoveRight unit:=wdCell ' переход в следующую ячейку таблицы Word
        End If
    Next i
    
    ' удаляем строку в конце таблицы
    .Cells.Delete shiftcells:=wdDeleteCellsEntireRow
End With



End Sub

Sub Worksheet_Change(ByVal target As Range)

Set d = Range("A1").CurrentRegion
m = d.Rows.Count
n = d.Columns.Count
Dim peresech As Range
Dim Sum As Integer, nize5 As Boolean, Sredniy As Single

Set peresech = Application.Intersect(target, d)
If Not (peresech Is Nothing) Then ' проверяем на наличие изменений
    For i = 1 To m
        ' вычисляем среднее
        Sum = 0
        nize5 = False
        
        For j = 2 To n
            Sum = Sum + d.Cells(i, j).Value
            If d.Cells(i, j).Value < 5 Then nize5 = True
        Next j
        Sredniy = Sum / (n - 1)
        
        If (Not nize5) And (Sredniy >= 7) Then
            ' выделяем зеленым
            For j = 1 To n
                d.Cells(i, j).Font.Color = vbGreen
                d.Cells(i, j).Font.Bold = True
            Next j
        Else
            ' не выделяем
            For j = 1 To n
                d.Cells(i, j).Font.Color = vbBlack
                d.Cells(i, j).Font.Bold = False
            Next j
  
        End If
    Next i
End If

End Sub
