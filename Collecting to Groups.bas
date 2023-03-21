Attribute VB_Name = "Module2"
Sub Сортировка()
Attribute Сортировка.VB_Description = "Макрос записан 22.09.2020 (MishOK)"
Attribute Сортировка.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Сортировка Макрос

' Макрос находит в таблице одинаковые элементы с разными обозначениями и группирует их
' Сочетание клавиш: Ctrl+ы

    Dim stroka As Integer
    Dim kol As Integer
    Dim nazv As String
    Dim oboz As String
    Dim List As String
    Dim Razmer As Integer
    
    List1 = "БКП"
    List2 = "Перечень"
     
    ' Определение размера таблицы
    For stroka = 2 To 2000
       If Worksheets(List1).Cells(stroka, 1).Value = 0 Then
       Razmer = stroka
       Exit For
       End If
    Next stroka
    
    For stroka = 3 To Razmer
    
     Worksheets(List2).Cells(stroka, 1).Value = Worksheets(List1).Cells(stroka - 1, 2).Value 'Наименование и тип элемента
     Worksheets(List2).Cells(stroka, 2).Value = Worksheets(List1).Cells(stroka - 1, 1).Value 'Обозначение элемента
     Worksheets(List2).Cells(stroka, 4).Value = Worksheets(List1).Cells(stroka - 1, 3).Value 'Базовая интенсивность отказов
     Worksheets(List2).Cells(stroka, 5).Value = Worksheets(List1).Cells(stroka - 1, 4).Value 'К1
     Worksheets(List2).Cells(stroka, 6).Value = Worksheets(List1).Cells(stroka - 1, 5).Value 'К2
     Worksheets(List2).Cells(stroka, 7).Value = Worksheets(List1).Cells(stroka - 1, 6).Value 'К3
     Worksheets(List2).Cells(stroka, 8).Value = Worksheets(List1).Cells(stroka - 1, 7).Value 'К4
     Worksheets(List2).Cells(stroka, 12).Value = Worksheets(List1).Cells(stroka - 1, 10).Value 'Кпр
     Worksheets(List2).Cells(stroka, 11).Value = Worksheets(List1).Cells(stroka - 1, 11).Value 'КЭ
     Worksheets(List2).Cells(stroka, 9).Value = Worksheets(List1).Cells(stroka - 1, 8).Value 'К5
     Worksheets(List2).Cells(stroka, 10).Value = Worksheets(List1).Cells(stroka - 1, 9).Value 'К6
     Worksheets(List2).Cells(stroka, 13).Value = Worksheets(List1).Cells(stroka - 1, 12).Value 'Интенсивность готовая
     Worksheets(List2).Cells(stroka, 15).Value = Worksheets(List1).Cells(stroka - 1, 14).Value 'Интенсивность хранения

    Next stroka
    
    kol = 0
    For stroka = 3 To Razmer
     nazv = Worksheets(List2).Cells(stroka, 1).Value 'Запомнить название элемента
        For strokain = 3 To Razmer
        'поиск таких же элементов
        Check = Worksheets(List2).Cells(strokain, 1).Value Like nazv 'если такой же элемент найден то
            If Check = True And Worksheets(List2).Cells(stroka, 13).Value = Worksheets(List2).Cells(strokain, 13).Value Then
            kol = kol + 1 'Увеличение счетчика количества элементов
            oboz = oboz + ", " + Worksheets(List2).Cells(strokain, 2).Value 'Формирование строки с обозначениями
            Worksheets(List2).Cells(strokain, 1).ClearContents 'Очистка значения наименования элемента
            End If
        Next strokain
        Worksheets(List2).Cells(stroka + Razmer + 2, 1).Value = nazv 'Наименование и тип элемента
        Worksheets(List2).Cells(stroka, 1).ClearContents 'Стираем найденный элемент всю строку
        oboz = Right(oboz, Len(oboz) - 2) 'Убрать 2 первых символа у обозначения
        Worksheets(List2).Cells(stroka + Razmer + 2, 2).Value = oboz 'Обозначение
        Worksheets(List2).Cells(stroka + Razmer + 2, 3).Value = kol 'Количество
        Worksheets(List2).Cells(stroka + Razmer + 2, 4).Value = Worksheets(List2).Cells(stroka, 4).Value 'Интенсивность базовая
        Worksheets(List2).Cells(stroka + Razmer + 2, 5).Value = Worksheets(List2).Cells(stroka, 5).Value 'К1
        Worksheets(List2).Cells(stroka + Razmer + 2, 6).Value = Worksheets(List2).Cells(stroka, 6).Value 'К2
        Worksheets(List2).Cells(stroka + Razmer + 2, 7).Value = Worksheets(List2).Cells(stroka, 7).Value 'К3
        Worksheets(List2).Cells(stroka + Razmer + 2, 8).Value = Worksheets(List2).Cells(stroka, 8).Value 'К4
        Worksheets(List2).Cells(stroka + Razmer + 2, 9).Value = Worksheets(List2).Cells(stroka, 9).Value 'К5
        Worksheets(List2).Cells(stroka + Razmer + 2, 10).Value = Worksheets(List2).Cells(stroka, 10).Value 'К6
        Worksheets(List2).Cells(stroka + Razmer + 2, 11).Value = Worksheets(List2).Cells(stroka, 11).Value 'КЭ
        Worksheets(List2).Cells(stroka + Razmer + 2, 12).Value = Worksheets(List2).Cells(stroka, 12).Value 'Кпр
        Worksheets(List2).Cells(stroka + Razmer + 2, 13).Value = Worksheets(List2).Cells(stroka, 13).Value 'Интенсивность
        Worksheets(List2).Cells(stroka + Razmer + 2, 14).Value = Worksheets(List2).Cells(stroka, 13).Value * kol 'Интенсивность по количеству
        Worksheets(List2).Cells(stroka + Razmer + 2, 15).Value = Worksheets(List2).Cells(stroka, 15).Value 'Интенсивность хранения
        Worksheets(List2).Cells(stroka + Razmer + 2, 16).Value = Worksheets(List2).Cells(stroka, 15).Value * kol 'Интенсивность хранени по количеству

        kol = 0
        oboz = Clear
    Next stroka
    
    ' Пройти по всем строкам и удалить строки с пустыми наименованиями
    For stroka = Razmer * 3 To 3 Step -1
       If Worksheets(List2).Cells(stroka, 1).Value = 0 Then
       Worksheets(List2).Rows(stroka).Delete
       End If
    Next stroka


End Sub


