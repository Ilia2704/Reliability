Attribute VB_Name = "Module1"
Sub ОбработкаПеречня()
Attribute ОбработкаПеречня.VB_Description = "Перегоняет перечень элементов в нормальный формат, создавая новый лист"
Attribute ОбработкаПеречня.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ОбработкаПеречня Макрос
' Перегоняет перечень элементов в нормальный формат, создавая новый лист

'Создние листа с именем блока
Dim Diap As String

List = ActiveSheet.Name
Nlist = Worksheets(List).Cells(2, 7).Value
SList = "Справочник"
Numb = 0
   'Application.DisplayAlerts = False
   'Sheets(Nlist).Delete
   'Application.DisplayAlerts = True

Sheets.Add.Name = Nlist 'имя листа определяется содержанием ячейки F2

'Создание шапки таблицы
     
     With Range("A1:N1")
     .Font.Name = "Arial"
     .HorizontalAlignment = -4108
     .VerticalAlignment = xlCenter
     .WrapText = True
     End With
     
     Rows(1).RowHeight = 42
     Columns(1).ColumnWidth = 9
     Worksheets(Nlist).Cells(1, 1).Value = "Элемент"
     Columns(2).ColumnWidth = 42
     Worksheets(Nlist).Cells(1, 2).Value = "Примечание"
     Columns(3).ColumnWidth = 16
     Worksheets(Nlist).Cells(1, 3).Value = "Базовая интенсивность"
     Columns(4).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 4).Value = "К1"
     Columns(5).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 5).Value = "К2"
     Columns(6).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 6).Value = "К3"
     Columns(7).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 7).Value = "К4"
     Columns(8).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 8).Value = "К5"
     Columns(9).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 9).Value = "К6"
     Columns(10).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 10).Value = "Кпр"
     Columns(11).ColumnWidth = 7
     Worksheets(Nlist).Cells(1, 11).Value = "Кэ"
     Columns(12).ColumnWidth = 19
     Worksheets(Nlist).Cells(1, 12).Value = "Эксплуатационная интенсивность"
     Columns(13).ColumnWidth = 19
     Worksheets(Nlist).Cells(1, 13).Value = "Вероятность безотказной работы"
     Columns(14).ColumnWidth = 15
     Worksheets(Nlist).Cells(1, 14).Value = "Интенсивность в режиме ожидания"
      
    'Определение длинны перечня элементов
    For strokaP = 1 To 10000
       If Worksheets(List).Cells(strokaP + 2, 8).Value = 0 Then
            RazmerP = strokaP - 1
       Exit For
       End If
    Next strokaP
        
    'Определение результирующей длинны таблицы
    Diap = 0
    For strokaR = 1 To RazmerP
        Diap = Diap + Worksheets(List).Cells(strokaR + 2, 9).Value * 1
    Next strokaR
    Range("A1:N" & Diap + 1).Borders.LineStyle = True 'Рамка для перечня

   
    'Определение длинны справочника
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
    
    Kolv = Worksheets(List).Cells(stroka + 2, 9).Value * 1 'Количество элементов данного типа
   
    Obozn1 = Worksheets(List).Cells(stroka + 2, 13).Value 'Обозначение берется из перечня элементов
     If InStr(1, Obozn1, "-") > 0 Then 'Если обозначение содержит тире, то необходимо взять только то, что слева от тире
        obozn2 = Left(Obozn1, InStr(1, Obozn1, "-") - 1)
     Else    'Если обозначение не содержит тире, то, возможно, оно содержит запятую
        If InStr(1, Obozn1, ",") > 0 Then
           obozn2 = Left(Obozn1, InStr(1, Obozn1, ",") - 1) 'Если оно содержит запятую - необходимо взять то, что слева от нее
           Else
           obozn2 = Obozn1 'Если обозначение не содержит ни запятой ни тире, то оно принимается для дальнейшей обработки без изменений
           End If
     End If

     obozn3 = Right(obozn2, Len(obozn2) - 1)
     If Val(obozn3) = 0 Then 'Если обозначение содержит две буквы, то необходимо взять два первых символа
        i = 2
     Else: i = 1   'Если обозначение содержит одну букву, то необходимо взять один символ
     End If
     zifir = Right(obozn2, Len(obozn2) - i) * 1 'За число в обозначении принимается все кроме букв
     Bukva = Left(obozn2, i) ' При формировании позиционного обозначения значения буквы объединяются со значением цифры + номер строки
     
         
    'Поиск по справочнику
      For strokaSS = 2 To RazmerS ' Цикл на заранее определенное длинной справочника число итераций
        'Проверка вхождения названия элемента из справочника в название элемента в перечне:
        If InStr(1, Worksheets(List).Cells(stroka + 2, 8).Value, Worksheets(SList).Cells(strokaSS, 2).Value) > 0 Then
        Pos = strokaSS 'Запоминание номера строки справочника, в которой найден искомый элемент
        Exit For
        Else: Pos = 2 'Если элемент не найден, то номер строки справочника "2"
        End If
      Next strokaSS
      
      ' Запись элемента
      For strokain = Numb + 1 To Kolv + Numb
       oboznach = Bukva & strokain + zifir - Numb - 1 'Формирование обозначения
       Worksheets(Nlist).Cells(strokain + 1, 1).Value = oboznach ' Обозначение элемента
      If Pos = 2 Then
      Worksheets(Nlist).Cells(strokain + 1, 2).Value = Worksheets(List).Cells(stroka + 2, 8).Value 'Наименование ненайденного элемента
      Else
       Worksheets(Nlist).Cells(strokain + 1, 2).Value = Worksheets(SList).Cells(Pos, 2).Value 'Название элемента
       Worksheets(Nlist).Cells(strokain + 1, 3).Value = Worksheets(SList).Cells(Pos, 3).Value 'Базовая интенсивность
       Worksheets(Nlist).Cells(strokain + 1, 4).Value = Worksheets(SList).Cells(Pos, 4).Value 'К1
       Worksheets(Nlist).Cells(strokain + 1, 5).Value = Worksheets(SList).Cells(Pos, 5).Value 'К2
       Worksheets(Nlist).Cells(strokain + 1, 6).Value = Worksheets(SList).Cells(Pos, 6).Value 'К3
       Worksheets(Nlist).Cells(strokain + 1, 7).Value = Worksheets(SList).Cells(Pos, 7).Value 'К4
       Worksheets(Nlist).Cells(strokain + 1, 8).Value = Worksheets(SList).Cells(Pos, 8).Value 'К5
       Worksheets(Nlist).Cells(strokain + 1, 9).Value = Worksheets(SList).Cells(Pos, 9).Value 'К6
       Worksheets(Nlist).Cells(strokain + 1, 10).Value = Worksheets(SList).Cells(Pos, 10).Value 'К7
       Worksheets(Nlist).Cells(strokain + 1, 11).Value = Worksheets(SList).Cells(Pos, 11).Value 'К8
       Worksheets(Nlist).Cells(strokain + 1, 14).Value = Worksheets(SList).Cells(Pos, 14).Value 'Интенсивность ожидания
       End If
      Next strokain
     Numb = Numb + Kolv 'Счетчик строк, необходимый для запоминания номера последней свободной строки
    
    Next stroka

End Sub

   

