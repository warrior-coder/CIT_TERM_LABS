VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ==================== лаб6 VBA часть 1 ====================

' нажатие кнопки замена
Private Sub Zamena_Click()
    Set d = Selection
    m = d.Rows.Count
    n = d.Columns.Count
   
    Dim zamen As Integer
    zamen = 0
    
    If PoVsemyDiapazony.Value = True Then
        ' Замена по всему диапазону
        For i = 1 To m
            For j = 1 To n
                If d.Cells(i, j).Value = CInt(ZamenitChislo.Value) Then
                    d.Cells(i, j).Value = CInt(ZamenitNa.Value)
                    zamen = zamen + 1
                End If
                
            Next j
        Next i
    Else
        ' Замена по строке
        i = CInt(NomerStroki.Value) ' номер строки для замены
        
        If i > m Then
            ' если номер строки больше количества строк
            MsgBox ("Ошибка: строка вне диапазона")
        Else
            ' замена по строке
            For j = 1 To n
                If d.Cells(i, j).Value = CInt(ZamenitChislo.Value) Then
                    d.Cells(i, j).Value = CInt(ZamenitNa.Value)
                    zamen = zamen + 1
                End If
            Next j
        End If
        
    End If
    
    ' Вывод количества замен
    If Podschet.Value = True Then
        KolichestvoZamen.Value = zamen
    Else
        KolichestvoZamen.Value = ""
    End If
    
    
End Sub

