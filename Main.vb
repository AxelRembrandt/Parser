Option Explicit
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Sub CeroMetralleta()

    'Application.Visible = 0
    
    'Внезапно переменные! Тысячи их!
    Dim i, j, k, o As Integer 'i,k,o - счетчики циклов; j - счетчик для записи в новый файл(по сути =№строки, в которую записываются значения)
    Dim dt As String 'для даты или псевдослуч. числа, записыв. в имя сохраняемого csv
    Dim cdc, name As String 'для долгого хранения значения ячейки с наименованием товара (для случая с несколькими цветами/опциями)
    Dim cdc2 As String 'Вспомогательная переменная для названий товаров. Все, что связано с переменной cdc2, - Заплатка Костылевна к 16-ти строкам прайса Мэдисон. (Мойки Елена, Дасти)
    Dim color As String 'Вспомогательная переменная для цветов Мэдисон. Используется в FindMadOptions.
    Dim s1, s2, s4 As Worksheet 'Вспомог. перем., хран. часть пути до ячейки на первом, втором, третьем листах файла парсера.
    Dim cemp As Boolean '"показатель опциональности" - соответствия ячейки с наимен.товара с возможными опциями след.усл.: не пустая; текст не выделен полужирным.
    Dim cemp2, cemp3, cha As Long 'Для двойных цен. Используется для записи первого вхождения подстроки в строку. (Мэдисон)| cha исп в вырезке цвета из арт-ов Хитэка.
    Dim pr As String 'Вспомогательная переменная для цен. Все, что связано с переменной pr, - Заплатка Костылевна к двойным ценам (Мэдисон).
    Dim hpr As Double 'используется для хранения и обработки цены в блоке Хитэк.
    Dim rname As String 'для объединенных частей общего названия (Мэдисон)
    Dim rcolor As String 'Исп. в FindMadOptions для сжатия опций в одну строку (Мэдисон).
    Dim dname As String 'Для частей названия в случае двойной цены. (Мэдисон)
    Dim q, w, e As Long 'вспомогательные переменные для аналогов и не только (Мэдисон)
    Dim analogue As String 'для вырезанных аналогов (Мэдисон) (Поиск и обработка аналогов не распространяется на товары с двойными ценами.)
    Dim art As String 'для артикулов Хитэк
    Dim kmaxhitek, kmaxmad, oexcept As Integer 'максимальное кол-во цветов, исключений (Хитэк) - используется как верхняя граница цикла
    Dim artwocolor As String 'Переменная для артикула с вырезанным цветом. Для вырезки цвета из арт-ов Хитэка.
    Dim artsamecol As Boolean 'Показывает, является ли артикул и цвет одной и той же строкой. (Хитэк)
    Dim qtyhit, qtymad As String 'Переменные для хранения количества товара Хитэк и Мэдисон соответственно.
    Dim kni As Integer 'Индикатор, показывающий, сколько раз был найден цвет в названии.
    Dim st As Integer 'Количество столбцов для выгрузки в .csv
        
    Dim lffpath, parsfn, tmprofifn, hitekfn, madisonfn, euromedfn As String 'lffpath - вспомогательная переменная, хранящая нужный путь до файлов; остальные - хранят названия файлов
    Dim np As Integer 'np - вспомогательная переменная, хранящая значение-индикатор для идентификации поставшика в ф-ии поиска файла и правильного выбора маски поиска.
        
    parsfn = ThisWorkbook.name 'Запись имени файла парсера
        
    'Блок поиска и открытия последних файлов
        
    lffpath = ThisWorkbook.Path & "\in_xls\hitek\"
    hitekfn = findfile(lffpath, 2)
    Workbooks.Open Filename:=lffpath & hitekfn
                
    'Конец блока
        
    Workbooks(parsfn).Activate
    
    cemp = 0
    artsamecol = 0
    j = 2
    kmaxhitek = 346
    oexcept = 5
    Set s1 = Workbooks(parsfn).Worksheets(1)
    Set s2 = Workbooks(parsfn).Worksheets(2)
    Set s4 = Workbooks(parsfn).Worksheets(3)
    
    'очистка старых записей
    For i = 1 To 12
        s1.Columns(i).Clear
    Next i
            
            
    '---------- Блок обработки Хитэк ----------

    For i = 13 To 15000
        'проверка наличия кода
        If (Workbooks(hitekfn).Worksheets(1).Cells(i, 5).Value <> "") Then
            
            'проверка заказного товара, наличия товара по цвету заливки
            If (Workbooks(hitekfn).Worksheets(1).Cells(i, 3).Interior.color = RGB(255, 255, 255) _
                And Workbooks(hitekfn).Worksheets(1).Cells(i, 3).Value <> "" _
                And InStr(1, LCase(Workbooks(hitekfn).Worksheets(1).Cells(i, 6).Value), "Товар под заказ", 0) = 0) Then
                
                art = Workbooks(hitekfn).Worksheets(1).Cells(i, 4).Value 'копирование арт-а в переменную
                name = Workbooks(hitekfn).Worksheets(1).Cells(i, 2).Value  'запись названия в переменную
                
                's1.Cells(j, 6).Value = name 'для проверки - по необходимости раскомментировать
                's1.Cells(j, 7).Value = cdc
                
                name = Trim(name): cdc = Trim(cdc) 'Очистка от лишних пробелов
                
                'Обработка цены
                If s4.Cells(16, 2).Value = True Then 'Проверка вкл/выкл вывода цены
                    hpr = Workbooks(hitekfn).Worksheets(1).Cells(i, 8).Value 'Запись цены в переменную.
                    If s4.Cells(18, 2).Value = True Then 'Проверка вкл/выкл округления в большую сторону
                        hpr = Application.WorksheetFunction.RoundUp(hpr, 0) 'Собственно округление
                    End If
                End If
                
                If (art <> "" _
                    And InStr(1, LCase(art), "экспер", 0) = 0) Then 'проверка наличия артикула и экспериментальных товаров
                    'Если включена запись опций товаров в 1 строку и арт совпадает с артом в предыдущей строке
                    If (s4.Cells(2, 2).Value = True And art = s1.Cells(j - 1, 1).Value) Then
                        j = j - 1 'Счетчик крутится назад, чтобы запись шла в ту же строку.
                    Else
                        s1.Cells(j, 1).Value = art 'вставка арт-а в новый файл
                    End If
                    'Проверка на повтор в цвете
                    If (InStr(name, "----//----") <> 0 And cdc <> "") Then 'проверка на случай, если цвет надо найти из "повтора" и этот "повтор" - опция (заплатка костылевна)
                        name = cdc 'название товара присваивается названию опции (используется вместо него, другими словами)
                    End If
                    
                Else 'если арт-а нету - "переход" на название
                    
                    'проверка на "повтор значения"
                    If (InStr(name, "----//----") <> 0) Then
                        If (cdc <> Workbooks(hitekfn).Worksheets(1).Cells(i - 1, 3).Value) Then 'проверка на совпадение с названием товара
                            name = Workbooks(hitekfn).Worksheets(1).Cells(i - 1, 3).Value 'запись в переменную названия из предыдущей строки
                            name = Trim(name) 'Очистка от лишних пробелов
                        Else
                            name = ""
                        End If
                    End If
                
                    If (cemp) Then 'проверка на опциональность товара: если у товара есть опции =>
                    
                        'проверка отсутствия кода товара в предыдущей строке или его соответствие коду в текущей строке
                        If (Workbooks(hitekfn).Worksheets(1).Cells(i - 1, 2).Value = "" _
                            Or Workbooks(hitekfn).Worksheets(1).Cells(i - 1, 2).Value = Workbooks(hitekfn).Worksheets(1).Cells(i, 5).Value) Then
                            'Если включена запись опций товаров в 1 строку и название совпадает с названием в предыдущей строке
                            If (s4.Cells(2, 2).Value = True And cdc = s1.Cells(j - 1, 1).Value) Then
                                j = j - 1 'Счетчик крутится назад, чтобы запись шла в ту же строку.
                                If (s1.Cells(j, 4).Value <> "") Then 'Проверка на пустую строку.
                                    s1.Cells(j, 4).Value = s1.Cells(j, 4).Value & s4.Cells(9, 2).Value & name 'Дозапись в ячейку наименования опции через разделитель.
                                Else
                                    s1.Cells(j, 4).Value = name 'Дозапись в ячейку наименования опции без разделителя.
                                End If
                            Else
                                s1.Cells(j, 1).Value = cdc 'запись наименования товара в новый файл
                                s1.Cells(j, 4).Value = name 'запись наименования опции
                            End If
                        Else 'Е ни одно из условий не удовл., сменился товар =>
                            cemp = 0 'обнуление показателя опциональности
                            cdc = "" 'очистка переменной с наименованием товара
                            s1.Cells(j, 1).Value = name 'запись наименования товара
                        End If
                        
                    Else 'если у товара нету опций =>
                        s1.Cells(j, 1).Value = name 'вставка названия в новый файл
                    End If
                    
                End If
                
                'Обработка и запись кол-ва / Проверка на наличие переносится на верхние уровни, т.к. появилось требование, по которому товары с кол-вом <= 0 не надо записывать.
                'Если включена запись опций товаров в 1 строку и ячейка с кол-вом не пустая
                If (s4.Cells(2, 2).Value = True And s1.Cells(j, 2).Value <> "") Then
                    'Дозапись в ячейку количества через разделитель.
                    s1.Cells(j, 2).Value = s1.Cells(j, 2).Value & s4.Cells(9, 2).Value & Workbooks(hitekfn).Worksheets(1).Cells(i, 3).Value
                Else
                    s1.Cells(j, 2).Value = Workbooks(hitekfn).Worksheets(1).Cells(i, 3).Value 'Копипаст кол-ва в новый файл
                End If
                
                'Запись цены
                'Если включена запись опций товаров в 1 строку и ячейка с ценой не пустая
                If (s4.Cells(2, 2).Value = True And s1.Cells(j, 3).Value <> "") Then
                    'Дозапись цены в ячейку через разделитель.
                    s1.Cells(j, 3).Value = s1.Cells(j, 3).Value & s4.Cells(9, 2).Value & hpr
                Else
                    s1.Cells(j, 3).Value = hpr 'Копипаст кол-ва в новый файл
                End If
                                                
                'Еще одна проверка на знак повтора уже записанного артикула/названия: если найден знак повтора, в ячейку записывается значение предыдущей ячейки
                If (InStr(s1.Cells(j, 1).Value, "----//----") <> 0) Then
                    s1.Cells(j, 1).Value = s1.Cells(j - 1, 1)
                End If
                
                'Блок поиска и записи цвета (вне зависимости от того, одинаковые ли артикулы)
                kni = 0 'Обнуление индикатора, показывающего, сколько раз был найден цвет в названии.
                For k = 1 To kmaxhitek
                    'поиск цвета из заданного списка в названии
                    cemp2 = InStr(1, LCase(name), LCase(s2.Cells(k, 1).Value), 0)
                    If (cemp2 <> 0) Then 'если цвет в названии найден
                        kni = kni + 1
                        
                        'Обработка исключений
                        For o = 1 To oexcept
                            'Проверяется, является ли найденный цвет возможным исключением и
                            'содержит ли название строку-исключение.
                            If (InStr(1, LCase(s2.Cells(o, 5).Value), LCase(s2.Cells(k, 1).Value), 0) <> 0 _
                                And InStr(1, LCase(name), LCase(s2.Cells(o, 6).Value), 0) <> 0) Then
                                Exit For 'Цикл прерывается
                            End If
                        Next o
                        If o <= oexcept Then
                            'В случае прерывания цикла пропускается блок записи цвета
                            GoTo ContinueToArtCol
                        End If
                        'Конец обработки исключений
                        
                        'проверяется, был ли какой-то цвет уже вписан
                        If (s1.Cells(j, 4).Value = "") Then 'если цвет еще не вписывался
                            'строка, равная по кол-ву символов строке с цветом, вытаскивается из названия и записывается
                            s1.Cells(j, 4).Value = _
                            Mid(name, cemp2, Len(s2.Cells(k, 1).Value))
                        Else 'если какой-то цвет уже был вписан,
                            rname = s1.Cells(j, 4).Value
                            'он сравнивается с выбранным
                            q = InStr(1, LCase(rname), LCase(s2.Cells(k, 1).Value), 0)
                            If (q = 0) Then 'если такой цвет еще не вписан, он дописывается к имеющемуся(имся) через разделитель
                                'Проверяется, сколько раз был найден к.-л. цвет в названии.
                                If (kni <= 1) Then
                                    s1.Cells(j, 4).Value = _
                                    rname & s4.Cells(9, 2).Value & Mid(name, cemp2, Len(s2.Cells(k, 1).Value)) 'вставка цвета в новый файл (возможно надо через 47 или 124)
                                Else
                                    s1.Cells(j, 4).Value = _
                                    rname & s4.Cells(19, 2).Value & Mid(name, cemp2, Len(s2.Cells(k, 1).Value)) 'вставка цвета в новый файл
                                End If
                            End If
                        End If
                    End If

ContinueToArtCol:
                    'Блок перезаписи арт-а/названия без цвета
                    If artsamecol = 0 Then
                        cemp2 = InStr(1, LCase(s1.Cells(j, 1).Value), LCase(s2.Cells(k, 1).Value), 0)
                        If cemp2 > 1 Then 'если цвет не с первого символа названия
                            cha = 1 'переменной значение передавай - ложных срабатываний не допускай!
                            'Суть переменной cha и проверок ниже в том, чтобы определить, стоят ли перед цветом какие-то лишние знаки, и убрать их вместе с цветом во время перезаписи
                            If (InStr(cemp2 - 1, s1.Cells(j, 1).Value, " ", 0) = (cemp2 - 1) Or InStr(cemp2 - 1, s1.Cells(j, 1).Value, "/", 0) = (cemp2 - 1) Or InStr(cemp2 - 1, s1.Cells(j, 1).Value, "-", 0) = (cemp2 - 1)) Then
                                cha = 2
                            End If
                            If (InStr(cemp2 - 2, s1.Cells(j, 1).Value, "/ ", 0) = (cemp2 - 2) Or InStr(cemp2 - 2, s1.Cells(j, 1).Value, "- ", 0) = (cemp2 - 2)) Then
                                cha = 3
                            End If
                            If (InStr(cemp2 - 3, s1.Cells(j, 1).Value, " / ", 0) = (cemp2 - 3) Or InStr(cemp2 - 3, s1.Cells(j, 1).Value, " - ", 0) = (cemp2 - 3)) Then
                                cha = 4
                            End If
                            'Проверка на отсутствие отдельно записанного цвета. В случае пустой строки туда пишется вырезаемый из арт-а/названия цвет.
                            If s1.Cells(j, 4) = "" Then
                                s1.Cells(j, 4).Value = Mid(s1.Cells(j, 1).Value, cemp2, Len(s2.Cells(k, 1).Value))
                            End If
                            'Запись артикула без цвета в переменную для удобства
                            artwocolor = Left(s1.Cells(j, 1).Value, cemp2 - cha) & Mid(s1.Cells(j, 1).Value, cemp2 & Len(s2.Cells(k, 1).Value))
                            'После вырезки цвета show must go on, т.е. должен остаться либо артикул целиком, либо его часть без цвета. Однако, некоторые арт-ы целиком состоят из цвета. В таком случае вырезать нельзя.
                            'Проверка длины строки после вырезки цвета и ее содержимого
                            If Len(artwocolor) > 1 Or artwocolor <> " " Then
                                'Собственно, перезапись названия уже без цвета и, предположительно, без лишних знаков перед ним
                                s1.Cells(j, 1).Value = artwocolor
                            End If
                        End If
                        If cemp2 = 1 Then 'Если цвет с первого символа арт-а, => арт=цвет и вырезать его не надо, а надо скопировать.
                            'Проверка на отсутствие отдельно записанного цвета. В случае пустой строки туда пишется копируемый из арт-а/названия цвет.
                            If s1.Cells(j, 4) = "" Then
                                s1.Cells(j, 4).Value = Mid(s1.Cells(j, 1).Value, cemp2, Len(s2.Cells(k, 1).Value))
                            End If
                            artsamecol = 1 'предотвращение перезаписи в след. итерациях цикла
                        End If
                    End If
                    'Конец блока перезаписи арт-а/названия без цвета
                    
                Next k
                                
                'Блок сжатия опций для артикулов, перезаписанных без цвета (очень топорно!! >=[ )
                'Если сжатие включено и артикулы стали одинаковые, то по сути строки объединяются путем добавления нижней к верхней.
                If (s4.Cells(2, 2).Value = True And s1.Cells(j, 1).Value = s1.Cells(j - 1, 1).Value) Then
                    s1.Cells(j - 1, 2).Value = s1.Cells(j - 1, 2).Value & s4.Cells(9, 2).Value & s1.Cells(j, 2).Value 'вставка кол-ва через разделитель
                    s1.Cells(j - 1, 3).Value = s1.Cells(j - 1, 3).Value & s4.Cells(9, 2).Value & s1.Cells(j, 3).Value 'вставка цены через разделитель
                    s1.Cells(j - 1, 4).Value = s1.Cells(j - 1, 4).Value & s4.Cells(9, 2).Value & s1.Cells(j, 4).Value 'Вставка цвета через разделитель.
                    s1.Cells(j, 1).Value = "": s1.Cells(j, 2).Value = "": s1.Cells(j, 3).Value = "": s1.Cells(j, 4).Value = "" 'Удаление записей из текущей строки.
                    j = j - 1 'Откат счетчика.
                End If
                'Конец блока сжатия
                            
                kni = 0: artsamecol = 0: hpr = 0 'восстановление исходного значения для корректной работы перезаписи в будущем
                'Конец блока поиска и записи цвета
                                                    
                j = j + 1 'увеличение счетчика записей в новом файле на единицу
            End If
        Else
            If (Workbooks(hitekfn).Worksheets(1).Cells(i, 2).Value <> "" _
                And Workbooks(hitekfn).Worksheets(1).Cells(i, 2).Font.Bold = False) Then 'проверка наличия в ячейке названия и отсутствия у оного полужирного начертания
                cdc = Workbooks(hitekfn).Worksheets(1).Cells(i, 2).Value 'копирование в переменную названия товара
                cemp = 1 'изменение значения показателя для того, чтобы в следующей/их итерации/ях цикла было известно,
                         'что идет работа с опциями одного товара и его название известно и записано в переменную
            Else
                cemp = 0
            End If
            
        End If
    Next i

    '------- Конец блока обработки Хитэк -------
      
    
    'Блок постобработки столбца опций/цветов/свойств.
    If s4.Cells(10, 2).Value = True Then 'Если в настройках включено заполнение пустых ячеек опций,
        For i = 2 To j - 1 'Перебираются все ячейки столбца с опциями.
            If s1.Cells(i, 4).Value = "" Then 'при найденной пустой ячейке
                If s1.Cells(i, 2).Value > 0 Then ' проверяется кол-во и, в зависимости от значения, ячейка заполняется.
                    s1.Cells(i, 4).Value = s4.Cells(11, 2).Value
                Else
                    s1.Cells(i, 4).Value = s4.Cells(12, 2).Value
                End If
            End If
        Next i
    End If

    'Очистка форматирования just for fun
    For i = 1 To 5
        s1.Columns(i).ClearFormats
    Next i
    
    
    'Удаление символов, которые используются как разделители.
    For i = 1 To j - 1
        For k = 1 To 5
            name = s1.Cells(i, k).Value
            s1.Cells(i, k).Value = Replace(name, s4.Cells(5, 2).Value, "")
            s1.Cells(i, k).Value = Replace(name, s4.Cells(6, 2).Value, "")
            s1.Cells(i, k).Value = Replace(name, s4.Cells(8, 2).Value, "")
            s1.Cells(i, k).Value = Replace(name, s4.Cells(5, 3).Value, "")
        Next k
    Next i
    
    'Сохранение в .csv; в название сохраняемого файла прописывать дату-время
    Dim saveName, outputStr, pathFile As String
        
    'Формируется путь и имя нового файла.
    pathFile = ThisWorkbook.Path & "\out_csv\"
    dt = Format(Now, "DD_MM_YYYY_HH-NN-SS")
    saveName = Left(parsfn, InStr(parsfn, ".") - 1) & "_" & dt & ".csv"
    
    'Определение кол-ва столбцов на выгрузку в зависимости от того, включен ли вывод аналогов у Мэдисон
    If (s4.Cells(4, 2).Value = True Or s4.Cells(4, 2).Value = "") Then
        st = 5
    Else
        st = 4
    End If
    
    'Блок записи всех ячеек в одну строковую переменную.
    For i = 2 To j - 1
        For k = 1 To st
            'Если включены границы столбца, идет запись с символами границ. Если нет - без них.
            If s4.Cells(7, 2) = True Then
                outputStr = outputStr & s4.Cells(8, 2).Value & CStr(s1.Cells(i, k).Value) & s4.Cells(8, 3).Value
            Else
                outputStr = outputStr & CStr(s1.Cells(i, k).Value)
            End If
            'В зависимости от того последний ли столбец(ячейка) в строке, дописывается нужный разделитель.
            If k = st Then
                'Дописывается разделитель строк
                outputStr = outputStr & s4.Cells(6, 2).Value
            Else
                'Дописывается разделитель столбцов
                outputStr = outputStr & s4.Cells(5, 2).Value
            End If
        Next k
        
        'Если в настройках указано, что вывод построчно, добавляется знак переноса строки.
        If (s4.Cells(1, 2).Value = True Or s4.Cells(1, 2).Value = "") Then
            outputStr = outputStr & Chr(10)
        End If
    Next i
    'Конец блока записи в строку.
     
    'Если в настройках активирована кодировка в UTF-8, осуществляется перекодировка.
    If (s4.Cells(3, 2).Value = True Or s4.Cells(3, 2).Value = "") Then
        outputStr = ToUTF8(outputStr) 'смена кодировки
    End If
        
    Open pathFile & saveName For Output As #1
    Print #1, outputStr;
    Close #1
    'Конец блока сохранения
    
    'Закрытие книг
    Workbooks(hitekfn).Close SaveChanges:=False
    'Workbooks(parsfn).Close SaveChanges:=True
    'Конец блока
    
    'Application.Visible = 1
    
End Sub

Public Function ToUTF8(ByVal sText As String) As String
 Dim nRet As Long, strRet As String
 
 strRet = String(Len(sText) * 2, vbNullChar)
 nRet = WideCharToMultiByte(65001, &H0, StrPtr(sText), Len(sText), StrPtr(strRet), Len(sText) * 2, 0&, 0&)
    
 ToUTF8 = Left(StrConv(strRet, vbUnicode), nRet)
End Function

Function findfile(ByVal lffpath As String, ByVal np As Integer)
    Dim myPath As String, myName As String, f As String
    
    myPath = lffpath 'Путь к папке с файлами
    myName = Dir(myPath & "*.xls*")
    f = myName
    'Перебор файлов по маске и дате изменения
    Do While myName <> ""
        'Если выбранный файл создан позже или одновременно с предыдущим...
        If (FileDateTime(myPath & myName) >= FileDateTime(myPath & f)) Then
            Select Case np 'Проверяется, с файлами какого поставщика идет работа
                Case 2 'Хитэк
                    'Проверка на содержание названием файла слова "остатки"
                    If (InStr(1, LCase(myName), "остатки", 0) >= 1) Then
                        f = myName
                    End If
                Case 3 'Мэдисон
                    'Проверка на содержание названием файла слова "остатки" и отсутствия в нем части слова "моск"
                    If (InStr(1, LCase(myName), "остатки", 0) >= 1 _
                        And InStr(1, LCase(myName), "моск", 0) = 0) Then
                            f = myName
                    End If
                Case Else 'Остальные
                    f = myName
            End Select
        End If
        myName = Dir
    Loop
    findfile = f
End Function

Sub ShowOptions() 'Вывод формы с настройками
    Options.Show
End Sub
