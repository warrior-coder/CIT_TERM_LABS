VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6612
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6612
   OleObjectBlob   =   "lab7_VBA-2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==================== лаб7 VBA задание 2 ====================

Private Sub UserForm_Initialize()
    Set d = Range("A1").CurrentRegion
    m = d.Rows.Count
    n = d.Columns.Count
    
    Dim names() As String
    ReDim names(1 To m)
    kol = 1
    Dim bylo As Boolean
    
    ' создаем массив уникальных имен
    names(1) = d.Cells(1, 2).Value
    For i = 2 To m
        tempName = d.Cells(i, 2)
        
        ' проверка наличия элемента
        bylo = False
        For j = 1 To kol
            If names(j) = tempName Then bylo = True
        Next j
        
        If bylo = False Then
            kol = kol + 1
            names(kol) = tempName
        End If
    Next i
    
    For k = 1 To kol
        Spisok.AddItem (names(k))
    Next k
    
    Spisok.MultiSelect = 2
    
End Sub

' нажатие кнопки [Выход]
Private Sub Zakrit_Click()
    Unload UserForm1
End Sub

' нажатие кнопки [Вычислить]
Private Sub Vichislit_Click()
    Set d = Range("A1").CurrentRegion
    m = d.Rows.Count
    n = d.Columns.Count
    
    Dim tovarySoSkidkoy() As String, kol As Byte
    ReDim tovarySoSkidkoy(1 To m)
    kol = 0
    
    ' перебор элементов списка
    For k = 0 To Spisok.ListCount - 1
    If Spisok.Selected(k) Then ' если товар выбран
        
        For i = 1 To m
        If d.Cells(i, 2).Value = Spisok.List(k) Then ' попали на выбранный товар
            If CInt(d.Cells(i, 4).Value) >= CInt(Granitsa.Value) Then ' если количество элементов > границы, то применяется скидка
                primenSkidka = True
                
                If SkidkaKChasti.Value Then ' скидка к части
                    d.Cells(i, 5).Value = d.Cells(i, 3).Value * (Granitsa.Value - 1) + d.Cells(i, 3).Value * (d.Cells(i, 4).Value - Granitsa.Value + 1) * (1 - Skidka.Value / 100)
                Else ' скидка ко всем
                    d.Cells(i, 5).Value = d.Cells(i, 3).Value * d.Cells(i, 4).Value * (100 - Skidka.Value) / 100
                End If
            Else ' количество элементов < границы => без скидки
                d.Cells(i, 5).Value = d.Cells(i, 3).Value * d.Cells(i, 4).Value
            End If
        End If
        Next i
        
        ' добавяем элемент в массив товаров с примененной скидкой
        If primenSkidka Then
            kol = kol + 1
            tovarySoSkidkoy(kol) = Spisok.List(k)
        End If
    Else
        ' товар не выбран => без скидки
        For i = 1 To m
            If Spisok.List(k) = d.Cells(i, 2).Value Then d.Cells(i, 5).Value = d.Cells(i, 3).Value * d.Cells(i, 4).Value
        Next i
    End If
    Next k
    
    ' вывод списка товаров с примененной скидкой
    Set d2 = Range(vivodDiapazon)
    d2.Cells(1, 1).Value = "товары со скодкой"
    For i = 1 To kol
        d2.Cells(i + 1, 1).Value = tovarySoSkidkoy(i)
    Next i
    
End Sub

