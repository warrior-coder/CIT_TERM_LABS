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
' ==================== ���7 VBA ������� 2 ====================

Private Sub UserForm_Initialize()
    Set d = Range("A1").CurrentRegion
    m = d.Rows.Count
    n = d.Columns.Count
    
    Dim names() As String
    ReDim names(1 To m)
    kol = 1
    Dim bylo As Boolean
    
    ' ������� ������ ���������� ����
    names(1) = d.Cells(1, 2).Value
    For i = 2 To m
        tempName = d.Cells(i, 2)
        
        ' �������� ������� ��������
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

' ������� ������ [�����]
Private Sub Zakrit_Click()
    Unload UserForm1
End Sub

' ������� ������ [���������]
Private Sub Vichislit_Click()
    Set d = Range("A1").CurrentRegion
    m = d.Rows.Count
    n = d.Columns.Count
    
    Dim tovarySoSkidkoy() As String, kol As Byte
    ReDim tovarySoSkidkoy(1 To m)
    kol = 0
    
    ' ������� ��������� ������
    For k = 0 To Spisok.ListCount - 1
    If Spisok.Selected(k) Then ' ���� ����� ������
        
        For i = 1 To m
        If d.Cells(i, 2).Value = Spisok.List(k) Then ' ������ �� ��������� �����
            If CInt(d.Cells(i, 4).Value) >= CInt(Granitsa.Value) Then ' ���� ���������� ��������� > �������, �� ����������� ������
                primenSkidka = True
                
                If SkidkaKChasti.Value Then ' ������ � �����
                    d.Cells(i, 5).Value = d.Cells(i, 3).Value * (Granitsa.Value - 1) + d.Cells(i, 3).Value * (d.Cells(i, 4).Value - Granitsa.Value + 1) * (1 - Skidka.Value / 100)
                Else ' ������ �� ����
                    d.Cells(i, 5).Value = d.Cells(i, 3).Value * d.Cells(i, 4).Value * (100 - Skidka.Value) / 100
                End If
            Else ' ���������� ��������� < ������� => ��� ������
                d.Cells(i, 5).Value = d.Cells(i, 3).Value * d.Cells(i, 4).Value
            End If
        End If
        Next i
        
        ' �������� ������� � ������ ������� � ����������� �������
        If primenSkidka Then
            kol = kol + 1
            tovarySoSkidkoy(kol) = Spisok.List(k)
        End If
    Else
        ' ����� �� ������ => ��� ������
        For i = 1 To m
            If Spisok.List(k) = d.Cells(i, 2).Value Then d.Cells(i, 5).Value = d.Cells(i, 3).Value * d.Cells(i, 4).Value
        Next i
    End If
    Next k
    
    ' ����� ������ ������� � ����������� �������
    Set d2 = Range(vivodDiapazon)
    d2.Cells(1, 1).Value = "������ �� �������"
    For i = 1 To kol
        d2.Cells(i + 1, 1).Value = tovarySoSkidkoy(i)
    Next i
    
End Sub

