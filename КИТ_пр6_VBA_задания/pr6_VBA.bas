Attribute VB_Name = "Module1"
' ==================== пр6 VBA ====================
Sub pr6_4()

Set d1 = Range("A1:C4")
Set d2 = Range("G1:I4")

m = d1.Rows.count
n = d1.Columns.count

Dim sum1 As Integer, sum2 As Integer
Dim count As Byte

count = 0
For i = 1 To m
    ' вычисляем сумму i строки 1 и 2 областей
    sum1 = 0
    sum2 = 0
    For j = 1 To n
        sum1 = sum1 + d1.Cells(i, j).Value
        sum2 = sum2 + d2.Cells(i, j).Value
    Next j
    
    If sum1 < sum2 Then
        ' обмен строк
        For j = 1 To n
            tempValue = d1.Cells(i, j).Value
            d1.Cells(i, j).Value = d2.Cells(i, j).Value
            d2.Cells(i, j).Value = tempValue
        Next j
        
        count = count + 1
    End If
Next i

MsgBox ("Строк поменялось: " & count)

End Sub
