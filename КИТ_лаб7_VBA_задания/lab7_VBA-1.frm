VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8388.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10008
   OleObjectBlob   =   "lab7_VBA-1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==================== ���7 VBA ������� 1 ====================

Private Sub UserForm1_Initialize()
    Chislo.Value = 0
End Sub

' �������� �������
Private Sub ChisloStchetchik_Change()
    Chislo.Value = ChisloStchetchik.Value
End Sub

' ������� ������ [�����]
Private Sub Vihod_Click()
    Unload UserForm1
End Sub

' ������� ������ [���������]
Private Sub Vipolnit_Click()
    
    If TipPodscheta.ListIndex = 0 Then
        ' ��� ��������: �� �������
        
        Set d = Range(IshodnieDannye.Value)
        m = d.Rows.Count
        n = d.Columns.Count
        chisloStrok = 0
        Dim estChislo As Boolean
        
        ' ������� ������ � �������� �����
        Dim nomeraStrok() As Byte, k As Byte
        ReDim nomeraStrok(1 To m)
        k = 0
        
        ' ������� ���������
        For i = 1 To m
            estChislo = False
            For j = 1 To n
                If d.Cells(i, j).Value = CInt(Chislo.Value) Then estChislo = True
            Next j
            
            If estChislo Then
                chisloStrok = chisloStrok + 1
                k = k + 1
                nomeraStrok(k) = i
            End If
        Next i
        
        MsgBox ("����� �����: " & chisloStrok)
        
        If VivodNomerov.Value = True Then
            ' ����� ������� ����� � ��������� ��������
            Set d2 = Range(Resultati.Value)
            
            d2.Cells(1, 1).Value = "������ �����:"
            For i = 1 To k
                d2.Cells(i + 1, 1).Value = nomeraStrok(i)
            Next i
        End If
    Else
        ' ��� ��������: ����� ���������
        Set d = Range(IshodnieDannye.Value)
        m = d.Rows.Count
        n = d.Columns.Count
        vsegoVhozdeniy = 0
        
        ' ������� ���������
        For i = 1 To m
            For j = 1 To n
                If d.Cells(i, j).Value = CInt(Chislo.Value) Then
                    vsegoVhozdeniy = vsegoVhozdeniy + 1
                End If
            Next j
        Next i
            
        MsgBox ("����� ���������: " & vsegoVhozdeniy)
    End If

End Sub
