Attribute VB_Name = "Module1"
Sub Data_to_Chart()
'   ����������
    Dim file_adres As String, file_name As String, current_file_adres As String
    Dim First_data_row As Integer, Last_data_row As Integer, New_initial_row As Integer, New_initial_column As Integer
    Dim Batches As Integer, Mesurements As Integer, number_of_elements As Integer
    Dim i As Integer, j As Integer, i_new As Integer, j_new As Integer, k As Integer
    Dim Int_average As Single, Offset As Single
    Dim Is_Dot As Integer
'
' ��� ������������� ������ ������������� �� ����� "some_error", ��� � ����� ����� ���������
' ���� �������� ��� �������� ����, ����� ��������� �� ����. ��� ��� ����������� - �� �����.
' ���� �������� "��", �� ��� �������� ���� ������������� �� �����.
' ���� �������� "���", �� ������� ������ ���-�� �� ������ 45-� ������ ����.
'  ����� ��� ����, ����� �� �������� ������������ On Error GoTo - ��� ���� ���� ������������ � ��� ���� :)
    On Error GoTo some_error
'
' �������������� � ���, ��� ����������� ���������� ����� ������ ���� ������, � �� �������.
' 52 = 4 (��/���) + 48 (��� ��������� - ��������������)
    Is_Dot = MsgBox("���������, ��� ����������� ���������� ����� - ��� �����.", 52)
    If Is_Dot = 7 Then
        Exit Sub
    End If
'
    Application.ScreenUpdating = False
    current_file_adres = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name  '���������� ������� ��� �����
'
' �������� ����
    file_adres = "False"
    file_adres = Application.GetOpenFilename("CSV(*.csv*),*.scv*", 1, , , False)
'
' ���������, ������ �� ���� ��� ��������.
    If file_adres = "False" Then
        Exit Sub
    End If
'
'
' ������-�� ������ �������� �� ����������. ���� �� ���� ����������� � �����������
' ������ ����� ����������� "On Error GoTo"
' ���������� �������� ���������� ���� �����.
'    If current_file_adres = file_adres Then
'        MsgBox "������ ���� ��� ������, �������� ������."
'        Exit Sub
'    End If
'
' ��������� ������ �� ����� � ������ "file_adres" � ��������� ������ �� ��������
    Workbooks.OpenText Filename:= _
        file_adres, Origin:=xlWindows, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
        ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=False, Comma:=True, _
        Space:=True, Other:=True, OtherChar:="""", FieldInfo:=Array(Array(1, 9), _
        Array(2, 1)), DecimalSeparator:=".", TrailingMinusNumbers:=True
'
' ��������� �� ������� ������ ����� ������ ��� ���. "-4" - 3 ������� ���������� csv + ������ �����.
    file_name = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 4)
'
' ���������� ����� �����, ���� ��� ������� �������, ��������� �� 31 �������
    If Len(file_name) > 31 Then
        file_name = Left(file_name, 31)
    End If
'
' ���������, ��� �� ���� ������
    If Cells(14, 2) <> "PBF" Then
        MsgBox "��� �� ���� ����������� ����� �� ������������ X-Supreme8000, ���� ���������, ���� ����� ������ � ���� ������� ������"
        Exit Sub
    End If
'
'
' ���������, ����� ����������� ���������� ����� ������������. ��� ����� �����. ���� ������������ �������, ������ �� �� �����.

'
' ������ ��������� ��������� ��������� ����� - ��� �������� ������ �� ���������� ���������, �����, � ����� ������ ���������� ������ � ������
    Batches = Cells(5, 2)
    Mesurements = Cells(6, 2) + 1
' ��������� ������ ������ � ������� - ��� ����� ���������� ������ ������� A ������� � 11-�� ������ �� ��� ���, ���� �� ��������� �� ������� ����� � �������.
    i = 11
    Do While Cells(i, 1) = ""
        i = i + 1
    Loop
'
    First_data_row = i + 1
    Last_data_row = First_data_row + Mesurements - 1
    Offset = 8 '� ������ ������ ��� ���������� ����� �� ��������� ������ ������ �� ������ ������ ���������� ����� ����� ������
'               ������ ���������� ����� ������ ����� ������������ ��� Last_data_row + Offset � ��� ���������� ���, ������ Batches
'   If Batches = 0 Then
'        MsgBox "���� �� �������� ������, ���� ���������, ���� ����� ������ � ���� ������� ������"
'        Exit Sub
'    End If
'
     If Batches <= 2 Or Mesurements <= 2 Then
        MsgBox "�� ���������� ������ ��� ���������� �������."
        Exit Sub
     End If
'
'������ ���������� ���������, ������� ����� ��������� ������������� � �����. ������ �������� � ������ 10 ������� �� ������� "C" - ������� "3"
    i = 3
    number_of_elements = 0
    Do While Len(Cells(10, i))
        number_of_elements = number_of_elements + 1
        i = i + 1
    Loop
'
'
'       MsgBox " ���������� ���������/���������: " & number_of_elements '��� ��������� ������� ������� ��������� ���������
'
'
' ������ ����� �������
    New_initial_column = 3 + number_of_elements + 2 '3 - � ���� ������� ������ ���������� ������ ������, +2 - ��������� �� 2 �������
    Range(Cells(First_data_row - 1, 3), Cells(First_data_row - 1, 3 + number_of_elements - 1)).Select
    Selection.Copy
    Range(Cells(1, New_initial_column), Cells(1, New_initial_column)).Select
    ActiveSheet.Paste
'
' ������ �������� ������ �������������� � ������ �������
    i = 3 '������� ���������/��������, ������ ������� � ������� = 3
    j = First_data_row '������� �����, ������ ������ � �������
    i_new = New_initial_column '������� ����� ��������
    j_new = 2 '������� ����� �����. 1-� ������ ������ ������, ������� ������ ������ �� 2-�� ������
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
' ������ �������
    New_initial_row = 2
    k = 0
    For k = 0 To number_of_elements - 1
        i = 0
        ' ��������� ������� ������������� ��� ���������� �������� �������
        Int_average = 0
        For i = 1 To Mesurements * Batches
            Int_average = Int_average + Cells(1 + i, New_initial_column + k)
            If Len(Cells(1 + i, New_initial_column + k)) = 0 Then '������� ����, ��� � ������ �����
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
        ' ������ �������� ��������
        ActiveChart.SetElement (msoElementChartTitleAboveChart)
        ActiveChart.ChartTitle.Text = Sheets(file_name).Cells(1, New_initial_column + k)
        '
        ' ��������� �������� �� ����� ����
        ActiveChart.Location Where:=xlLocationAsNewSheet
        ' ������ �������� ������ �����
        ActiveSheet.Name = Sheets(file_name).Cells(1, New_initial_column + k)
        Sheets(file_name).Select
        Cells(1.1).Select
    Next k
'
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
'
'
'����� ��� �������� ��� ������������� ������
some_error:
' ��� ���� ����� ������� ��������� � �������. ���� �����������, ����� �� ��������
'    If Err.Description <> "" Then
'        MsgBox "������: " & Err.Description
'    End If

End Sub





