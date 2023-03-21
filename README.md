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
![image](https://user-images.githubusercontent.com/111389991/226732274-ee03c7f9-87a6-4be4-9556-a279883c34a7.png)

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
