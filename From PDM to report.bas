Attribute VB_Name = "Module1"
Sub ����������������()
Attribute ����������������.VB_Description = "���������� �������� ��������� � ���������� ������, �������� ����� ����"
Attribute ����������������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ���������������� ������
' ���������� �������� ��������� � ���������� ������, �������� ����� ����

'������� ����� � ������ �����
Dim Diap As String

List = ActiveSheet.Name
Nlist = Worksheets(List).Cells(2, 7).Value
SList = "����������"
Numb = 0
   'Application.DisplayAlerts = False
   'Sheets(Nlist).Delete
   'Application.DisplayAlerts = True

Sheets.Add.Name = Nlist '��� ����� ������������ ����������� ������ F2

'�������� ����� �������
     
     With Range("A1:N1")
     .Font.Name = "Arial"
     .HorizontalAlignment = -4108
     .VerticalAlignment = xlCenter
     .WrapText = True
     End With
     
     Rows(1).RowHeight = 42
     Columns(1).ColumnWidth = 9
     Worksheets(Nlist).Cells(1, 1).Value = "�������"
     Columns(2).ColumnWidth = 42
     Worksheets(Nlist).Cells(1, 2).Value = "����������"
     Columns(3).ColumnWidth = 16
     Worksheets(Nlist).Cells(1, 3).Value = "������� �������������"
     Columns(4).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 4).Value = "�1"
     Columns(5).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 5).Value = "�2"
     Columns(6).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 6).Value = "�3"
     Columns(7).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 7).Value = "�4"
     Columns(8).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 8).Value = "�5"
     Columns(9).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 9).Value = "�6"
     Columns(10).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 10).Value = "���"
     Columns(11).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 11).Value = "��"
     Columns(12).ColumnWidth = 19
     Worksheets(Nlist).Cells(1, 12).Value = "���������������� �������������"
     Columns(13).ColumnWidth = 19
     Worksheets(Nlist).Cells(1, 13).Value = "����������� ����������� ������"
     Columns(14).ColumnWidth = 15
     Worksheets(Nlist).Cells(1, 14).Value = "������������� � ������ ��������"
      
    '����������� ������ ������� ���������
    For strokaP = 1 To 10000
       If Worksheets(List).Cells(strokaP + 2, 8).Value = 0 Then
            RazmerP = strokaP - 1
       Exit For
       End If
    Next strokaP
        
    '����������� �������������� ������ �������
    Diap = 0
    For strokaR = 1 To RazmerP
        Diap = Diap + Worksheets(List).Cells(strokaR + 2, 9).Value * 1
    Next strokaR
    Range("A1:N" & Diap + 1).Borders.LineStyle = True '����� ��� �������

   
    '����������� ������ �����������
    For strokaS = 1 To 2000
       If Worksheets(SList).Cells(strokaS, 2).Value = 0 Then
       RazmerS = strokaS - 1
       Exit For
       End If
    Next strokaS
      
    Numb = 0

    For stroka = 1 To RazmerP
    
    Obozn1 = 0
    obozn3 = 0
    obozn2 = 0
    
    Kolv = Worksheets(List).Cells(stroka + 2, 9).Value * 1 '���������� ��������� ������� ����
   
    Obozn1 = Worksheets(List).Cells(stroka + 2, 13).Value '����������� ������� �� ������� ���������
     If InStr(1, Obozn1, "-") > 0 Then '���� ����������� �������� ����, �� ���������� ����� ������ ��, ��� ����� �� ����
        obozn2 = Left(Obozn1, InStr(1, Obozn1, "-") - 1)
     Else    '���� ����������� �� �������� ����, ��, ��������, ��� �������� �������
        If InStr(1, Obozn1, ",") > 0 Then
           obozn2 = Left(Obozn1, InStr(1, Obozn1, ",") - 1) '���� ��� �������� ������� - ���������� ����� ��, ��� ����� �� ���
           Else
           obozn2 = Obozn1 '���� ����������� �� �������� �� ������� �� ����, �� ��� ����������� ��� ���������� ��������� ��� ���������
           End If
     End If

     obozn3 = Right(obozn2, Len(obozn2) - 1)
     If Val(obozn3) = 0 Then '���� ����������� �������� ��� �����, �� ���������� ����� ��� ������ �������
        i = 2
     Else: i = 1   '���� ����������� �������� ���� �����, �� ���������� ����� ���� ������
     End If
     zifir = Right(obozn2, Len(obozn2) - i) * 1 '�� ����� � ����������� ����������� ��� ����� ����
     Bukva = Left(obozn2, i) ' ��� ������������ ������������ ����������� �������� ����� ������������ �� ��������� ����� + ����� ������
     
         
    '����� �� �����������
      For strokaSS = 2 To RazmerS ' ���� �� ������� ������������ ������� ����������� ����� ��������
        '�������� ��������� �������� �������� �� ����������� � �������� �������� � �������:
        If InStr(1, Worksheets(List).Cells(stroka + 2, 8).Value, Worksheets(SList).Cells(strokaSS, 2).Value) > 0 Then
        Pos = strokaSS '����������� ������ ������ �����������, � ������� ������ ������� �������
        Exit For
        Else: Pos = 2 '���� ������� �� ������, �� ����� ������ ����������� "2"
        End If
      Next strokaSS
      
      ' ������ ��������
      For strokain = Numb + 1 To Kolv + Numb
       oboznach = Bukva & strokain + zifir - Numb - 1 '������������ �����������
       Worksheets(Nlist).Cells(strokain + 1, 1).Value = oboznach ' ����������� ��������
      If Pos = 2 Then
      Worksheets(Nlist).Cells(strokain + 1, 2).Value = Worksheets(List).Cells(stroka + 2, 8).Value '������������ ������������ ��������
      Else
       Worksheets(Nlist).Cells(strokain + 1, 2).Value = Worksheets(SList).Cells(Pos, 2).Value '�������� ��������
       Worksheets(Nlist).Cells(strokain + 1, 3).Value = Worksheets(SList).Cells(Pos, 3).Value '������� �������������
       Worksheets(Nlist).Cells(strokain + 1, 4).Value = Worksheets(SList).Cells(Pos, 4).Value '�1
       Worksheets(Nlist).Cells(strokain + 1, 5).Value = Worksheets(SList).Cells(Pos, 5).Value '�2
       Worksheets(Nlist).Cells(strokain + 1, 6).Value = Worksheets(SList).Cells(Pos, 6).Value '�3
       Worksheets(Nlist).Cells(strokain + 1, 7).Value = Worksheets(SList).Cells(Pos, 7).Value '�4
       Worksheets(Nlist).Cells(strokain + 1, 8).Value = Worksheets(SList).Cells(Pos, 8).Value '�5
       Worksheets(Nlist).Cells(strokain + 1, 9).Value = Worksheets(SList).Cells(Pos, 9).Value '�6
       Worksheets(Nlist).Cells(strokain + 1, 10).Value = Worksheets(SList).Cells(Pos, 10).Value '�7
       Worksheets(Nlist).Cells(strokain + 1, 11).Value = Worksheets(SList).Cells(Pos, 11).Value '�8
       Worksheets(Nlist).Cells(strokain + 1, 14).Value = Worksheets(SList).Cells(Pos, 14).Value '������������� ��������
       End If
      Next strokain
     Numb = Numb + Kolv '������� �����, ����������� ��� ����������� ������ ��������� ��������� ������
    
    Next stroka

End Sub

   

