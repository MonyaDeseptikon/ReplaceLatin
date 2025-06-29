'Проверка кодировки 2
Attribute VB_Name = "replaceLatin"
Option Explicit
Sub НаЛистеЗаменаЛатинскихБуквНаКириллическиеВнутреняя(ByRef cellCount As Long, ByRef replaceCount As Long, sheetNum As Variant)
'Переменные
Dim sheetSource As Worksheet, cellCheck As Range, rowCount As Long, colFIAS As Integer, FIAS As Object
colFIAS = 0

'Задание и подготовка рабочего листа 2
Set sheetSource = ActiveWorkbook.Worksheets(sheetNum)
If sheetSource.Visible = False Then MsgBox ("Указанный лист скрыт, сделайте лист видимым или введите другой номер листа"): Exit Sub
If sheetSource.FilterMode Then sheetSource.ShowAllData
If ActiveWindow.FreezePanes = True Then ActiveWindow.FreezePanes = False
sheetSource.Columns.EntireColumn.Hidden = False
sheetSource.Rows.EntireRow.Hidden = False
'Счетчики
rowCount = sheetSource.UsedRange.Cells(Rows.Count, 1).End(xlUp).Row
Set FIAS = sheetSource.Range(Cells(1, 1).End(xlToRight), Cells(2, 1)).Find("ФИАС", MatchCase:=True) 'Поиск ФИАС только в первой и второй строках
If Not FIAS Is Nothing Then colFIAS = FIAS.Column
cellCount = cellCount + sheetSource.UsedRange.Count
'Тело
    For Each cellCheck In sheetSource.UsedRange
        If IsError(cellCheck) Or cellCheck.HasFormula Or cellCheck.Column = colFIAS Then
        Else








           If cellCheck Like "*[CcEeTOopPAaHKkXxBMy]*" Then
            replaceCount = replaceCount + 1 'Счетчик ячеек, в которых проведена замена
            cellCheck.Value = Replace(cellCheck, "C", "С", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "c", "с", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "E", "Е", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "e", "е", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "T", "Т", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "O", "О", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "o", "о", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "p", "р", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "P", "Р", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "A", "А", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "a", "а", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "H", "Н", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "K", "К", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "k", "к", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "X", "Х", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "x", "х", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "B", "В", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "M", "М", , , vbBinaryCompare)
            cellCheck.Value = Replace(cellCheck, "y", "у", , , vbBinaryCompare)
            End If
        End If
    Next cellCheck
End Sub

Sub ЗаменаЛатинскихБуквНаКириллические()
'Переменные
Dim startTime As String, cellCount As Long, replaceCount As Long, sheetNum As Variant
replaceCount = 0
cellCount = 0
'Время старта
startTime = "Время старта: " & Now
'Отключение среды Эксель
Application.Calculation = xlManual
Application.ScreenUpdating = False
Application.EnableCancelKey = xlDisabled
Application.EnableEvents = False
    'Вечный модуль, - пока не будет введен корректный путь , либо закрыт макрос
    On Error Resume Next
        Do
            If Err <> 0 Then Err.Clear: MsgBox "Введен не верный номер"
                sheetNum = Application.InputBox(prompt:="Проверка проводится по всем столбцам, кроме ""Уникальный номер в ГАР (ID FIAS)""." & vbCrLf & _
                "Будут удалены все фильтры и закрепы окна" & vbCrLf & _
                "Введите номер обрабатываемого листа (слева направо): ", Title:="Задание рабочего листа", Type:=1)
            If sheetNum = False Then Exit Sub
            If sheetNum <= 0 Or sheetNum > ActiveWorkbook.Worksheets.Count Then Err = 1
        Loop While Err <> 0
    On Error GoTo 0
    Call НаЛистеЗаменаЛатинскихБуквНаКириллическиеВнутреняя(cellCount, replaceCount, sheetNum)
'Включение среды Эксель

Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.EnableCancelKey = xlInterrupt
Application.EnableEvents = False
Application.CutCopyMode = False
'Сообщения, время финиша
MsgBox (startTime & vbCrLf & "Время окончания: " & Now)
MsgBox ("Всего ячеек к рассмотрению: " & cellCount & vbCrLf & _
"Найдены латинские символы в " & replaceCount & " ячейках")
End Sub


