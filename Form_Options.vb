Private Sub CommandButton1_Click()
    Dim s4 As Worksheet 'Вспомог. перем., хран. часть пути до ячейки на четвертом листе файла парсера.
    Dim Sep(4) As String 'Массив с разделителями
    Dim i, j As Integer 'Счетчики цикла
    Set s4 = ActiveWorkbook.Worksheets(3)
    
    'Запись в массив значений разделителей для проверки совпадений
    Sep(0) = TextBox5.Value 'Разделитель столбцов
    Sep(1) = TextBox6.Value 'Разделитель строк
    Sep(2) = TextBox7.Value 'Левый символ границ столбца
    Sep(3) = TextBox1.Value 'Разделитель значений в столбцах при сжатии
    Sep(4) = TextBox9.Value 'Разделитель повторяющихся цветов
    
    'Проверка совпадений
    For i = 0 To 4
        If Sep(i) <> "" Then
        For j = 0 To 4
            If i <> j And Sep(i) = Sep(j) Then
                MsgBox "Ошибка! Найдено совпадение разделителей!"
                Exit Sub
            End If
        Next j
        End If
    Next i
    
    'Сохранение значений чекбоксов.
    s4.Cells(1, 2).Value = CBool(CheckBox1.Value) 'Многострочный вывод
    s4.Cells(2, 2).Value = CBool(CheckBox2.Value) 'Опции в одну строку
    s4.Cells(3, 2).Value = CBool(CheckBox3.Value) 'Кодировка в UTF-8
    s4.Cells(5, 2).Value = TextBox5.Value 'Разделитель столбцов
    s4.Cells(6, 2).Value = TextBox6.Value 'Разделитель строк
    s4.Cells(7, 2).Value = CBool(CheckBox9.Value) 'Границы столбца включены
    s4.Cells(8, 2).Value = TextBox7.Value 'Левый символ границ столбца
    s4.Cells(8, 3).Value = TextBox8.Value 'Правый символ границ столбца
    s4.Cells(9, 2).Value = TextBox1.Value 'Разделитель значений в столбцах при сжатии
    s4.Cells(10, 2).Value = CBool(CheckBox8.Value) 'Заполнение, если нету опции
    s4.Cells(11, 2).Value = TextBox2.Value 'Если товар есть в наличии
    s4.Cells(12, 2).Value = TextBox3.Value 'Если товар отсутствует
    s4.Cells(13, 2).Value = TextBox4.Value 'Если нету кол-ва, писать
    s4.Cells(16, 2).Value = CBool(CheckBox7.Value) 'Хитэк: Выводить цену
    s4.Cells(17, 2).Value = CBool(CheckBox6.Value) 'Хитэк: Вывод отсутствующих товаров
    s4.Cells(18, 2).Value = CBool(CheckBox11.Value) 'Хитэк: округление цены в большую сторону
    s4.Cells(19, 2).Value = TextBox9.Value 'Разделитель повторяющихся цветов

    Unload Options 'Закрытие формы и выгрузка из памяти.
End Sub

Private Sub CommandButton2_Click()
    Unload Options 'Закрытие формы.
End Sub

Private Sub TextBox7_Change()
    Select Case TextBox7.Value
        'При вводе в поле левого символа парного знака в поле ввода правого появляется его пара.
        Case "("
            TextBox8.Value = ")"
        Case "{"
            TextBox8.Value = "}"
        Case "["
            TextBox8.Value = "]"
        Case "<"
            TextBox8.Value = ">"
        Case "«"
            TextBox8.Value = "»"
        Case Else
            'В остальных случаях при изменении левого символа границы столбца правый принимает его значение.
            TextBox8.Value = TextBox7.Value
    End Select
End Sub

Private Sub UserForm_Initialize()
    Dim s4 As Worksheet 'Вспомог. перем., хран. часть пути до ячейки на четвертом листе файла парсера.
    Set s4 = ActiveWorkbook.Worksheets(3)
    
    'Считывание значений чекбоксов.
    CheckBox1.Value = s4.Cells(1, 2).Value 'Многострочный вывод
    CheckBox2.Value = s4.Cells(2, 2).Value 'Опции в одну строку
    CheckBox3.Value = s4.Cells(3, 2).Value 'Кодировка в UTF-8
    TextBox5.Value = s4.Cells(5, 2).Value 'Разделитель столбцов
    TextBox6.Value = s4.Cells(6, 2).Value 'Разделитель строк
    CheckBox9.Value = s4.Cells(7, 2).Value 'Границы столбца включены
    TextBox7.Value = s4.Cells(8, 2).Value 'Левый символ границ столбца
    TextBox8.Value = s4.Cells(8, 3).Value 'Правый символ границ столбца
    TextBox1.Value = s4.Cells(9, 2).Value 'Разделитель значений в столбцах при сжатии
    CheckBox8.Value = s4.Cells(10, 2).Value 'Заполнение, если нету опции
    TextBox2.Value = s4.Cells(11, 2).Value 'Если товар есть в наличии
    TextBox3.Value = s4.Cells(12, 2).Value 'Если товар отсутствует
    TextBox4.Value = s4.Cells(13, 2).Value 'Если нету кол-ва, писать
    CheckBox7.Value = s4.Cells(16, 2).Value 'Хитэк: Выводить цену
    CheckBox6.Value = s4.Cells(17, 2).Value 'Хитэк: Вывод отсутствующих товаров
    CheckBox11.Value = s4.Cells(18, 2).Value 'Хитэк: округление цены в большую сторону
    TextBox9.Value = s4.Cells(19, 2).Value 'Разделитель повторяющихся цветов
End Sub

