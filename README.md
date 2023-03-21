# Reliability

  The algorithm assumes the presence of a list of elements of a separate block formed by the designer, which can be obtained from the PDM system; and a reference book containing a list of elements, reliability indicators and correction coefficients for which were previously obtained. The format of the list of elements and the format of the reference book are shown
  
  ![image](https://user-images.githubusercontent.com/111389991/226730845-b57832e5-8acb-4075-a40d-70bcb7c40baf.png)
  
The list of elements includes the necessary fields: positional designation, element name and quantity.

![image](https://user-images.githubusercontent.com/111389991/226731323-bb2b45b9-cd14-4ad7-8eb9-8fc3a88ccfdb.png)

The algorithm for forming a table with the results of calculations is implemented in the MS VBA programming environment and includes the following steps:

1. Creating a template table with the results of the reliability calculation.
 
2. Determination of the length of the initial list of elements, the length of the reference book and the length of the results table.

3. Starting the cycle of iterating through the list of elements. This loop processes one row of the list of elements in one iteration. The length of the list was previously determined.

3.1 Determination of the number of elements of this current name. The quantity value is taken from the list of elements.

3.2 Processing of the positional designation.


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
     

3.3 Starting the directory search cycle. The number of iterations of this cycle is determined by the length of the directory.

3.3.1 Comparison of the directory line with the name of the list item.

3.3.2 Memorizing the line number in which the elements match. If the element is not found in the directory, then the value of the variable line number is assigned to the value of the second line, on which the value of the string variable "Element is not found" is placed in the directory in advance.

    For strokaSS = 2 To RazmerS ' Цикл на заранее определенное длинной справочника число итераций
        'Проверка вхождения названия элемента из справочника в название элемента в перечне:
        If InStr(1, Worksheets(List).Cells(stroka + 2, 8).Value, Worksheets(SList).Cells(strokaSS, 2).Value) > 0 Then
        Pos = strokaSS 'Запоминание номера строки справочника, в которой найден искомый элемент
        Exit For
        Else: Pos = 2 'Если элемент не найден, то номер строки справочника "2"
        End If
      Next strokaSS

3.4 Starting the cycle of recording the found element in the table with the results of calculations.

It should be noted here that the found match in the directory should be recorded as many times as the elements of the current given type are listed in the list of elements. This quantity is expressed as the number of elements of this type and is indicated by a specific number. In addition, the next element should be recorded on the last free line. To do this, a counter of "free rows" is entered.

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
     
 After completing the main cycle of sorting through the list of elements, the program stops working. All the attributes necessary for further calculations are placed in the created table of the results of the calculation of reliable news
 
 ![image](https://user-images.githubusercontent.com/111389991/226735212-75daa82a-8aa1-4e09-acb7-7b34c8c6d31d.png)

Further calculation involves taking into account the possible redundancy of individual elements into complex structures. The described approach to the automation of "routine" reliability calculation procedures is also suitable for the formation of accounting documentation. We are talking about an appendix to the calculation of reliability news, in the table of which it is necessary to reflect the composition of each block element by element.

![image](https://user-images.githubusercontent.com/111389991/226735362-e71bf4f6-49d5-48ae-9c68-c3a1a9b6c40d.png)

To form such a table from the reliability calculation results table , it is necessary: 

1. Create a template for the application table to the report. 

2. Transfer the values of the elements of the reliability calculation results table to the fields of the application table

3. Start the cycle of iterating through the application table. 

3.1 Compare the current element with each of the following elements and write such an element with all the values of the fields of the current row to the end of the table. 

3.2 Write the positions of the found elements to the service string variable.

3.3 Increase the counter of the number of elements of the found type by one point. 

3.4 Delete the values of the names of elements not only from the found elements, but also from the current one. At the same time, the line itself remains. 

3.5 Assign the values of service variables to the quantity and positional designation fields for each element, respectively. 

4. Start a cycle during which every line containing an empty name is deleted
