Attribute VB_Name = "Module2"
Sub ����������()
Attribute ����������.VB_Description = "������ ������� 22.09.2020 (MishOK)"
Attribute ����������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ���������� ������

' ������ ������� � ������� ���������� �������� � ������� ������������� � ���������� ��
' ��������� ������: Ctrl+�

    Dim stroka As Integer
    Dim kol As Integer
    Dim nazv As String
    Dim oboz As String
    Dim List As String
    Dim Razmer As Integer
    
    List1 = "���"
    List2 = "��������"
     
    ' ����������� ������� �������
    For stroka = 2 To 2000
       If Worksheets(List1).Cells(stroka, 1).Value = 0 Then
       Razmer = stroka
       Exit For
       End If
    Next stroka
    
    For stroka = 3 To Razmer
    
     Worksheets(List2).Cells(stroka, 1).Value = Worksheets(List1).Cells(stroka - 1, 2).Value '������������ � ��� ��������
     Worksheets(List2).Cells(stroka, 2).Value = Worksheets(List1).Cells(stroka - 1, 1).Value '����������� ��������
     Worksheets(List2).Cells(stroka, 4).Value = Worksheets(List1).Cells(stroka - 1, 3).Value '������� ������������� �������
     Worksheets(List2).Cells(stroka, 5).Value = Worksheets(List1).Cells(stroka - 1, 4).Value '�1
     Worksheets(List2).Cells(stroka, 6).Value = Worksheets(List1).Cells(stroka - 1, 5).Value '�2
     Worksheets(List2).Cells(stroka, 7).Value = Worksheets(List1).Cells(stroka - 1, 6).Value '�3
     Worksheets(List2).Cells(stroka, 8).Value = Worksheets(List1).Cells(stroka - 1, 7).Value '�4
     Worksheets(List2).Cells(stroka, 12).Value = Worksheets(List1).Cells(stroka - 1, 10).Value '���
     Worksheets(List2).Cells(stroka, 11).Value = Worksheets(List1).Cells(stroka - 1, 11).Value '��
     Worksheets(List2).Cells(stroka, 9).Value = Worksheets(List1).Cells(stroka - 1, 8).Value '�5
     Worksheets(List2).Cells(stroka, 10).Value = Worksheets(List1).Cells(stroka - 1, 9).Value '�6
     Worksheets(List2).Cells(stroka, 13).Value = Worksheets(List1).Cells(stroka - 1, 12).Value '������������� �������
     Worksheets(List2).Cells(stroka, 15).Value = Worksheets(List1).Cells(stroka - 1, 14).Value '������������� ��������

    Next stroka
    
    kol = 0
    For stroka = 3 To Razmer
     nazv = Worksheets(List2).Cells(stroka, 1).Value '��������� �������� ��������
        For strokain = 3 To Razmer
        '����� ����� �� ���������
        Check = Worksheets(List2).Cells(strokain, 1).Value Like nazv '���� ����� �� ������� ������ ��
            If Check = True And Worksheets(List2).Cells(stroka, 13).Value = Worksheets(List2).Cells(strokain, 13).Value Then
            kol = kol + 1 '���������� �������� ���������� ���������
            oboz = oboz + ", " + Worksheets(List2).Cells(strokain, 2).Value '������������ ������ � �������������
            Worksheets(List2).Cells(strokain, 1).ClearContents '������� �������� ������������ ��������
            End If
        Next strokain
        Worksheets(List2).Cells(stroka + Razmer + 2, 1).Value = nazv '������������ � ��� ��������
        Worksheets(List2).Cells(stroka, 1).ClearContents '������� ��������� ������� ��� ������
        oboz = Right(oboz, Len(oboz) - 2) '������ 2 ������ ������� � �����������
        Worksheets(List2).Cells(stroka + Razmer + 2, 2).Value = oboz '�����������
        Worksheets(List2).Cells(stroka + Razmer + 2, 3).Value = kol '����������
        Worksheets(List2).Cells(stroka + Razmer + 2, 4).Value = Worksheets(List2).Cells(stroka, 4).Value '������������� �������
        Worksheets(List2).Cells(stroka + Razmer + 2, 5).Value = Worksheets(List2).Cells(stroka, 5).Value '�1
        Worksheets(List2).Cells(stroka + Razmer + 2, 6).Value = Worksheets(List2).Cells(stroka, 6).Value '�2
        Worksheets(List2).Cells(stroka + Razmer + 2, 7).Value = Worksheets(List2).Cells(stroka, 7).Value '�3
        Worksheets(List2).Cells(stroka + Razmer + 2, 8).Value = Worksheets(List2).Cells(stroka, 8).Value '�4
        Worksheets(List2).Cells(stroka + Razmer + 2, 9).Value = Worksheets(List2).Cells(stroka, 9).Value '�5
        Worksheets(List2).Cells(stroka + Razmer + 2, 10).Value = Worksheets(List2).Cells(stroka, 10).Value '�6
        Worksheets(List2).Cells(stroka + Razmer + 2, 11).Value = Worksheets(List2).Cells(stroka, 11).Value '��
        Worksheets(List2).Cells(stroka + Razmer + 2, 12).Value = Worksheets(List2).Cells(stroka, 12).Value '���
        Worksheets(List2).Cells(stroka + Razmer + 2, 13).Value = Worksheets(List2).Cells(stroka, 13).Value '�������������
        Worksheets(List2).Cells(stroka + Razmer + 2, 14).Value = Worksheets(List2).Cells(stroka, 13).Value * kol '������������� �� ����������
        Worksheets(List2).Cells(stroka + Razmer + 2, 15).Value = Worksheets(List2).Cells(stroka, 15).Value '������������� ��������
        Worksheets(List2).Cells(stroka + Razmer + 2, 16).Value = Worksheets(List2).Cells(stroka, 15).Value * kol '������������� ������� �� ����������

        kol = 0
        oboz = Clear
    Next stroka
    
    ' ������ �� ���� ������� � ������� ������ � ������� ��������������
    For stroka = Razmer * 3 To 3 Step -1
       If Worksheets(List2).Cells(stroka, 1).Value = 0 Then
       Worksheets(List2).Rows(stroka).Delete
       End If
    Next stroka


End Sub


