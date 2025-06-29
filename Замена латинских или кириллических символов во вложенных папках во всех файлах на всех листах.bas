'Проверка кодировки
Attribute VB_Name = "FolderFile"
Option Explicit
Sub ВПапкахЗаменаЛатинскихБуквНаКириллическиеВнутреняя() 'опасность в том, что работа происходит с открытой книгой, которая открыта н экране. А может это и замедление работы макроса
'Переменные
Dim path As Variant, folder As Object, fileSys As Object, startTime As String, cellCount As Long, replaceCount As Long, fileCount As Long
fileCount = 0
replaceCount = 0
cellCount = 0
'Время старта
startTime = "Время старта: " & Now
'Отключение среды Эксель
Application.Calculation = xlManual
Application.ScreenUpdating = False
Application.EnableCancelKey = xlDisabled
Application.EnableEvents = False
'Задание пространства и подготовка
Set fileSys = CreateObject("Scripting.FileSystemObject")
'Вечный модуль, - пока не будет введен корректный путь , либо закрыт макрос
    On Error Resume Next
        Do
            If Err <> 0 Then Err.Clear: MsgBox "Введен не верный путь"
            path = Application.InputBox(prompt:="Скопируйте и вставьте путь к папке", Title:="Задание пути к папке", Type:=2)
            If path = False Then Exit Sub
            Set folder = fileSys.GetFolder(path)
        Loop While Err <> 0
    On Error GoTo 0
'Тело
    Call РаботаСКаждойПапкой(fileSys, folder, cellCount, replaceCount, fileCount)
'Включение среды Эксель
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.EnableCancelKey = xlInterrupt
Application.EnableEvents = False
Application.CutCopyMode = False

'Сообщения, время финиша
MsgBox (startTime & vbCrLf & "Время окончания: " & Now)
MsgBox ("Проверка проводилась по всем столбцам, кроме ""Уникальный номер в ГАР (ID FIAS)""." & vbCrLf & _
"Всего файлов рассмотрено: " & fileCount & vbCrLf & _
"Всего ячеек рассмотрено: " & cellCount & vbCrLf & _
"Найдены латинские символы в " & replaceCount & " ячейках")
End Sub

Sub РаботаСКаждымЛистом(ByRef cellCount As Long, ByRef replaceCount As Long)
'Переменные
Dim sheetChek As Worksheet, wbOpen As Workbook
Set wbOpen = ActiveWorkbook
    For Each sheetChek In wbOpen.Sheets
        wbOpen.Worksheets(sheetChek.Index).Activate
        Call НаЛистеЗаменаЛатинскихБуквНаКириллическиеВнутреняя(cellCount, replaceCount, sheetChek.Index)
    Next sheetChek
End Sub

Sub РаботаСКаждымФайлом(fileSys As Object, folder As Object, ByRef cellCount As Long, ByRef replaceCount As Long, ByRef fileCount As Long)
'Переменные
Dim file As Object, ext As String
    For Each file In folder.files
        ext = StrConv(fileSys.GetExtensionName(file.path), vbLowerCase)
        '2 - это обозначение HIDDEN
        If file.Attributes And 2 Or Not (ext = "xlsx" Or ext = "xlsb" Or ext = "xls" Or ext = "ods") Then
            Else
            Workbooks.Open file, UpdateLinks:=0, Notify:=False, ReadOnly:=False, IgnoreReadOnlyRecommended:=True
                  Call РаботаСКаждымЛистом(cellCount, replaceCount)
            Workbooks(fileSys.GetFileName(file)).CheckCompatibility = False 'Отключить проверку совместимости при сохранении
            Workbooks(fileSys.GetFileName(file)).Close SaveChanges:=True
        End If
            fileCount = fileCount + 1
    Next file
End Sub

Sub РаботаСКаждойПапкой(fileSys As Object, folder As Object, ByRef cellCount As Long, ByRef replaceCount As Long, ByRef fileCount As Long)
'Переменные
Dim subFolders As Object, files As Object, fld As Object
Set subFolders = folder.subFolders
Set files = folder.files
    If files.Count <> 0 Then Call РаботаСКаждымФайлом(fileSys, folder, cellCount, replaceCount, fileCount)
    If subFolders.Count <> 0 Then
        For Each fld In subFolders
           Call РаботаСКаждойПапкой(fileSys, fld, cellCount, replaceCount, fileCount)
        Next fld
    End If
End Sub
