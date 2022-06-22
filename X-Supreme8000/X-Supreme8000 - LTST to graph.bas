Attribute VB_Name = "Module1"
Sub Data_to_Chart()
'   Переменные
    Dim file_adres As String, file_name As String, current_file_adres As String
    Dim First_data_row As Integer, Last_data_row As Integer, New_initial_row As Integer, New_initial_column As Integer
    Dim Batches As Integer, Mesurements As Integer, number_of_elements As Integer
    Dim i As Integer, j As Integer, i_new As Integer, j_new As Integer, k As Integer
    Dim Int_average As Single, Offset As Single
    Dim Is_Dot As Integer
'
' при возникновении ошибки перескакиваем на метку "some_error", она в самом конце программы
' если выбираем уже открытый файл, будет сообщение об этом. как его перехватить - на нашел.
' если ответить "да", то уже открытый файл перезапишется по новой.
' если ответить "нет", то вылезет ошибка где-то на уровне 45-й строки кода.
'  чтобы для того, чтобы ее избежать используется On Error GoTo - все тело кода проскакиваем и все дела :)
    On Error GoTo some_error
'
' предупреждение о том, что разделитель десятичной части должен быть точкой, а не запятой.
' 52 = 4 (да/нет) + 48 (тип сообщения - предупреждение)
    Is_Dot = MsgBox("Проверьте, что разделитель десятичной части - это точка.", 52)
    If Is_Dot = 7 Then
        Exit Sub
    End If
'
    Application.ScreenUpdating = False
    current_file_adres = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name  'Определяем текущее имя файла
'
' Выбираем файл
    file_adres = "False"
    file_adres = Application.GetOpenFilename("CSV(*.csv*),*.scv*", 1, , , False)
'
' Проверяем, выбран ли файл для открытия.
    If file_adres = "False" Then
        Exit Sub
    End If
'
'
' почему-то данная проверка не заработала. пока не стал углубляться и закомментил
' вместо этого использовал "On Error GoTo"
' Собственно проверка совпадения имен файла.
'    If current_file_adres = file_adres Then
'        MsgBox "Данный файл уже открыт, выберите другой."
'        Exit Sub
'    End If
'
' Вставляем данные из файла с именем "file_adres" и разбиваем данные по столбцам
    Workbooks.OpenText Filename:= _
        file_adres, Origin:=xlWindows, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
        ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=False, Comma:=True, _
        Space:=True, Other:=True, OtherChar:="""", FieldInfo:=Array(Array(1, 9), _
        Array(2, 1)), DecimalSeparator:=".", TrailingMinusNumbers:=True
'
' Извлекаем из полного адреса файла только его имя. "-4" - 3 символа расширения csv + символ точки.
    file_name = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 4)
'
' Определяем длину файла, если имя слишком длинное, сокращаем до 31 символа
    If Len(file_name) > 31 Then
        file_name = Left(file_name, 31)
    End If
'
' проверяем, тот ли файл открыт
    If Cells(14, 2) <> "PBF" Then
        MsgBox "Это не файл результатов теста на стабильность X-Supreme8000, либо поврежден, либо новая версия и надо править макрос"
        Exit Sub
    End If
'
'
' Проверяем, какой разделитель десятичной части используется. Нам нужна точка. Есил используется запятая, меняем ее на точку.

'
' Задаем начальные константы обработки файла - где хранятся данные по количеству измерений, серий, с какой строки начинаются данные и прочее
    Batches = Cells(5, 2)
    Mesurements = Cells(6, 2) + 1
' вычисляем первую строку с данными - для этого перебираем ячейки столбца A начиная с 11-ой строки до тех пор, пока не доберемся до первого блока с данными.
    i = 11
    Do While Cells(i, 1) = ""
        i = i + 1
    Loop
'
    First_data_row = i + 1
    Last_data_row = First_data_row + Mesurements - 1
    Offset = 8 'в данном случае это количество строк от последней строки данных до первой строки следующего блока серии данных
'               начало следующего блока данных будет определяться как Last_data_row + Offset и так количество раз, равное Batches
'   If Batches = 0 Then
'        MsgBox "Файл не содержит данных, либо поврежден, либо новая версия и надо править макрос"
'        Exit Sub
'    End If
'
     If Batches <= 2 Or Mesurements <= 2 Then
        MsgBox "Не достаточно данных для построения графика."
        Exit Sub
     End If
'
'Теперь необходимо посчитать, сколько всего элементов задействовано в тесте. Данные хранятся в строке 10 начиная со столбца "C" - столбец "3"
    i = 3
    number_of_elements = 0
    Do While Len(Cells(10, i))
        number_of_elements = number_of_elements + 1
        i = i + 1
    Loop
'
'
'       MsgBox " количество элементов/сегментов: " & number_of_elements 'для дэбагинга выводим сколько сегментов насчитали
'
'
' задаем шапку таблицы
    New_initial_column = 3 + number_of_elements + 2 '3 - в этом столбце должны находиться первые данные, +2 - смещаемся на 2 столбца
    Range(Cells(First_data_row - 1, 3), Cells(First_data_row - 1, 3 + number_of_elements - 1)).Select
    Selection.Copy
    Range(Cells(1, New_initial_column), Cells(1, New_initial_column)).Select
    ActiveSheet.Paste
'
' Теперь собираем данные интенсивностей в единую таблицу
    i = 3 'счетчик сегментов/столбцов, первый столбец с данными = 3
    j = First_data_row 'счетчик строк, первая строка с данными
    i_new = New_initial_column 'счетчик новых столбцов
    j_new = 2 'счетчик новых строк. 1-я строка занята шапкой, поэтому данные пойдут со 2-ой строки
    For k = 1 To Batches
        Range(Cells(j, i), Cells(j + Mesurements - 1, i + number_of_elements - 1)).Select
        Selection.Copy
        Range(Cells(j_new, i_new), Cells(j_new, i_new)).Select
        ActiveSheet.Paste
        j = j + Mesurements - 1 + Offset
        j_new = j_new + Mesurements
'        k = k + 1
    Next k
'
' Рисуем графики
    New_initial_row = 2
    k = 0
    For k = 0 To number_of_elements - 1
        i = 0
        ' вычисляем среднюю интенсивность для подстройки масштаба графика
        Int_average = 0
        For i = 1 To Mesurements * Batches
            Int_average = Int_average + Cells(1 + i, New_initial_column + k)
            If Len(Cells(1 + i, New_initial_column + k)) = 0 Then 'условие того, что в ячейке пусто
                Exit For
            End If
        Next i
        Int_average = Int_average / (i - 1)
        '
        Range(Selection, Selection.End(xlDown)).Select
        Range(Cells(New_initial_row, New_initial_column + k), Cells(New_initial_row + Mesurements * Batches - 1, New_initial_column + k)).Select
        ActiveSheet.Shapes.AddChart.Select
        ActiveChart.ChartType = xlXYScatterSmooth
        ActiveChart.SetSourceData Source:=Range(Cells(New_initial_row, New_initial_column + k), Cells(New_initial_row + Mesurements * Batches - 1, New_initial_column + k))
        ActiveChart.Axes(xlValue).MinimumScale = Int_average - Int_average * 0.1
        ActiveChart.Axes(xlValue).MaximumScale = Int_average + Int_average * 0.1
        ' Задаем название диаграмы
        ActiveChart.SetElement (msoElementChartTitleAboveChart)
        ActiveChart.ChartTitle.Text = Sheets(file_name).Cells(1, New_initial_column + k)
        '
        ' переносим диаграму на новый лист
        ActiveChart.Location Where:=xlLocationAsNewSheet
        ' задаем название нового листа
        ActiveSheet.Name = Sheets(file_name).Cells(1, New_initial_column + k)
        Sheets(file_name).Select
        Cells(1.1).Select
    Next k
'
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
'
'
'метка для перехода при возникновении ошибки
some_error:
' эта част может вывести сообщение с ошибкой. пока закомментил, чтобы не мешалось
'    If Err.Description <> "" Then
'        MsgBox "ошибка: " & Err.Description
'    End If

End Sub





