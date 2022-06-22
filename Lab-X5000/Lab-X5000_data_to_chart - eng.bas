Sub Кнопка1_Щелчок()
'Sub Макрос1()
'
' Макрос1 Макрос
'   Задаем переменные
    Dim file_adres As String, file_name As String, current_file_adres As String, version As String
    Dim First_data_row As Integer, Second_data_row As Integer, Last_data_row As Integer
    Dim Offset As Single, Energy_channel As Single
    Dim max_kV As Integer
    Dim i As Integer, j As Integer
'
Application.ScreenUpdating = False
current_file_adres = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name  'узнаем текущее имя файла
'
' выбираем файл из проводника
    file_adres = "False"
    file_adres = Application.GetOpenFilename("JSON(*.json*),*.json*", 1, , , False)
'
' проверяем, был ли выбран адрес. если нет, то тормозим макрос
    If file_adres = "False" Then
        Exit Sub
    End If
'
' проверяем, открыт ли уже выбранный файл
    If current_file_adres = file_adres Then
        MsgBox "Этот файл уже открыт"                     'rus
'        MsgBox "This file is open. Chose another one."   'eng
        Exit Sub
    End If
'
' Открываем файл по адресу file_adres и разбиваем данные по столбцам
    Workbooks.OpenText Filename:= _
        file_adres, Origin:=xlWindows, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
        ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=False, Comma:=True, _
        Space:=True, Other:=True, OtherChar:="""", FieldInfo:=Array(Array(1, 9), _
        Array(2, 1)), DecimalSeparator:=".", TrailingMinusNumbers:=True
'
    file_name = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) ' извлекаем имя файла без расширения. Располагаться код должен именно здесь так как здесь уже активна новая рабочая книга
'
' проверяем, тот ли файл открыт
    If Range("A2") <> "FWHM" Then
        MsgBox "Это не файл спектра Lab-X5000"      'rus
'        MsgBox "It wasn't Lab-X5000 spectrum file"  'eng
        Exit Sub
    End If
'
' выясняем версию отчета (их два) - в зависимости от версии завит расположение данных в строках
    If Range("A4157") = "version" Then
        version = "1.0.0"
    ElseIf Range("A4158") = "version" Then
        version = "2.0.0"
    Else
        MsgBox "Либо файл поврежден, либо не тот, либо новая версия отчета и нужно менять номера ячеек в скрипте"   'rus
'        MsgBox "File corupted or it's new version of Sepctrum file and change should be done in macros"             'eng
        Exit Sub
    End If
'
    max_kV = Range("C24") + 1    ' Запоминаем максимальное значение энергии, будет откладываться по шкале X - берем из ускоряющего напряжения понадобится для подстройки оси X графика
'
' отталкиваясь от версии отчета выбираем ячейки с данными
    If version = "1.0.0" Then
        First_data_row = 40
        Second_data_row = 41
        Last_data_row = 4135
        Offset = Cells(4141, 1)
        Energy_channel = Cells(4145, 1)
    ElseIf version = "2.0.0" Then
        First_data_row = 41
        Second_data_row = 42
        Last_data_row = 4136
        Offset = Cells(4142, 1)
        Energy_channel = Cells(4146, 1)
    End If
'
' добавляем столбец с номерами каналов и рассчитанными энергиями
    j = 0
    Cells(First_data_row, 2).Select
    For i = First_data_row To Last_data_row
        j = j + 1
        Cells(i, 3) = j                                            'столбец номеров каналов
        Cells(i, 2) = Offset + Cells(i, 3) * Energy_channel / 1000 'столбец энергий соответствующих каналам
    Next i
'
' строим график - строим его очень херово, через говнокод
    Range(Selection, Selection.End(xlDown)).Select
    Range(Cells(First_data_row, 1), Cells(Last_data_row, 2)).Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlXYScatterSmoothNoMarkers
    ActiveChart.SetSourceData Source:=Range(Cells(First_data_row, 1), Cells(Last_data_row, 2))
    ActiveChart.SeriesCollection(1).Delete
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).XValues = Sheets(file_name).Range(Sheets(file_name).Cells(First_data_row, 2), Sheets(file_name).Cells(Last_data_row, 2))
    ActiveChart.SeriesCollection(1).Values = Sheets(file_name).Range(Sheets(file_name).Cells(First_data_row, 1), Sheets(file_name).Cells(Last_data_row, 1))
'    ActiveChart.Axes(xlCategory).MinimumScale = 1 / 2
    ActiveChart.Axes(xlCategory).MaximumScale = max_kV
    ActiveChart.Location Where:=xlLocationAsNewSheet
'
Application.ScreenUpdating = True



End Sub





