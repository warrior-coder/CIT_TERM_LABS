VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "����1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ==================== ���6 VBA ����� 2 ====================

Private Sub TipZameni_Click()

End Sub

' ������� ������ ������
Private Sub Zamena_Click()
    Set d = Selection
    m = d.Rows.Count
    n = d.Columns.Count

    Dim zamen As Integer
    zamen = 0
    
    If TipZameni.ListIndex = 1 Then
        ' ������ �� ����� ���������
        For i = 1 To m
            For j = 1 To n
                If d.Cells(i, j).Value = CInt(ZamenitChislo.Value) Then
                    d.Cells(i, j).Value = CInt(ZamenitNa.Value)
                    zamen = zamen + 1
                End If
                
            Next j
        Next i
    Else
        ' ������ �� ������
        i = CInt(NomerStroki.Value) ' ����� ������ ��� ������
        
        If i > m Then
            ' ���� ����� ������ ������ ���������� �����
            MsgBox ("������: ������ ��� ���������")
        Else
            ' ������ �� ������
            For j = 1 To n
                If d.Cells(i, j).Value = CInt(ZamenitChislo.Value) Then
                    d.Cells(i, j).Value = CInt(ZamenitNa.Value)
                    zamen = zamen + 1
                End If
            Next j
        End If
        
    End If
    
    ' ����� ���������� �����
    If Podschet.Value = True Then
        KolichestvoZamen.Value = zamen
    Else
        KolichestvoZamen.Value = ""
    End If
    
    
End Sub

