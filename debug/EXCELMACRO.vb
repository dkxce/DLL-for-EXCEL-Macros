Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpFunc As Long, ByVal Param1 As Long, ByVal Param2 As Long, ByVal Param3 As Long, ByVal Param4 As Long) As Long
 
Private Const DLL_NAME = "exceltools.dll"
Dim handle As Long, address As Long, Unload As Long

Function GetLibPath() As String 'Полный путь к DLL
    Dim nPath As String
    nPath = ActiveWorkbook.FullName
    nPath = Left(nPath, Len(nPath) - Len(Split(nPath, "\")(UBound(Split(nPath, "\"))))) & DLL_NAME
    GetLibPath = nPath
End Function

 
Function LoadLib() 'ЗАГРУЗКА DLL
    If handle = 0 Then
        handle = LoadLibrary(GetLibPath)
        If handle <> 0 Then
            'MsgBox ("Библиотека успешно загружена.")
        ElseIf handle = 0 Then
            MsgBox ("Ошибка при загрузке DLL!")
        End If
    End If
End Function


Function UnloadLib() 'ВЫГРУЗКА DLL
    If handle <> 0 Then
        Unload = FreeLibrary(handle)
        If Unload <> 0 Then
            'MsgBox ("Библиотека выгружена успешно.")
        ElseIf Unload = 0 Then
            MsgBox ("Ошибка при выгрузке библиотеки из памяти!")
        End If
    End If
End Function
 
Function ShowInfo() 'ИНФОРМАЦИЯ О DLL
    If handle <> 0 Then
        Dim msginfo As String
        Dim libName As String
        Dim libMethods As Integer
        Dim libMNames As String
        Dim libScripts As Integer
        Dim libSNames As String
        Dim i As Integer
                
        libName = GetLibName
        libMethods = GetLibraryMethods
        libScripts = GetLibraryScripts
            
        For i = 0 To libMethods - 1
            libMNames = libMNames & vbNewLine & " - " & GetLibraryMethodName(i) & " - " & CInt(i)
        Next i
        
        For i = 0 To libScripts - 1
            libSNames = libSNames & vbNewLine & " - " & GetLibraryScriptName(i) & " - " & CInt(i)
        Next i
        
        msginfo = "Загружена библиотека: " & vbNewLine & "  " & libName & vbNewLine & vbNewLine & "Методов в библиотеке: " & libMethods & ":" & vbNewLine & libMNames
        msginfo = msginfo & vbNewLine & vbNewLine & "Скриптов в библиотеке: " & libScripts & ":" & vbNewLine & libSNames
        MsgBox msginfo
    End If
End Function

 
Function GetLibName() As String ' Method 0 - Внутреннее имя библиотеки
    GetLibName = ""
    Dim result As String
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetLibraryName") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            CopyMemory result, address, 4
            GetLibName = result
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function

Function GetLibraryMethods() As Integer ' Method 1 - Число поддерживаемых методов библиотеки
    GetLibraryMethods = 0
    Dim result As String
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetLibraryMethods") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            GetLibraryMethods = address
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function

Function GetLibraryMethodName(num As Integer) As String ' Method 2 - Имя метода в библиотеки из списка
    GetLibraryMethodName = ""
    Dim result As String
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetLibraryMethodName") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal num, ByVal 0&, ByVal 0&, ByVal 0&)
            CopyMemory result, address, 4
            GetLibraryMethodName = result
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function


Function PassCell(row As Long, column As Long, value As String, formula As String) As Integer ' Method 3 - ЗАПОЛНЕНИЕ ЯЧЕЕК В БИБЛИОТЕКЕ ИЗ ЛИСТА
    PassCell = -1
    If handle <> 0 Then
        address = GetProcAddress(handle, "PassCell") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal row, ByVal column, StrPtr(StrConv(value, vbFromUnicode)), StrPtr(StrConv(formula, vbFromUnicode)))
            PassCell = address
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function

Function GetCellValue(row As Long, column As Long) As String ' Method 4 - ПОЛУЧЕНИЕ ЗНАЧЕНИЕ ЯЧЕЙКИ ИЗ БИБЛИОТЕКИ
    GetCellValue = ""
    Dim result As String
    
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetCellValue") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal row, ByVal column, ByVal 0&, ByVal 0&)
            CopyMemory result, address, 4
            GetCellValue = result
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function

Function GetCellFormula(row As Long, column As Long) As String ' Method 5 - ПОЛУЧЕНИЕ ЗНАЧЕНИЕ ФОРМУЛЫ ИЗ БИБЛИОТЕКИ
    GetCellFormula = ""
    Dim result As String
    
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetCellFormula") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal row, ByVal column, ByVal 0&, ByVal 0&)
            CopyMemory result, address, 4
            GetCellFormula = result
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
 
Function ClearData() As Integer ' Method 6 - Очистка всех ячеек в библиотеке
    ClearData = -1
    If handle <> 0 Then
        address = GetProcAddress(handle, "ClearData") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            ClearData = address
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
 
Function GetFilledCells() As Integer ' Method 7 - Число заполненных ячеек в библиотеке
    GetFilledCells = -1
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetFilledCells") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            GetFilledCells = address
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
 
Function GetBounds(index As Integer) As String ' Method 8 - Диапазон заполненных ячеек в библиотеке (границы)
' 0 - MinRow, 1 - MaxRow, 2 - MinCol, 3 - MaxCol
    GetBounds = ""
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetBounds") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal index, ByVal 0&, ByVal 0&, ByVal 0&)
            GetBounds = address
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function

Function RunScript(script As String) As Integer ' Method 9 - Запуск скрипта из библиотеки
    RunScript = -1
    If handle <> 0 Then
        address = GetProcAddress(handle, "RunScript") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, StrPtr(StrConv(script, vbFromUnicode)), ByVal 0&, ByVal 0&, ByVal 0&)
            RunScript = address
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
 
Function GetLibraryScripts() As Integer ' Method 10 - Число поддерживаемых скриптов библиотеки
    GetLibraryScripts = 0
    Dim result As String
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetLibraryScripts") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            GetLibraryScripts = address
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function

Function GetLibraryScriptName(num As Integer) As String ' Method 11 - Имя скрипта в библиотеки из списка
    GetLibraryScriptName = ""
    Dim result As String
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetLibraryScriptName") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal num, ByVal 0&, ByVal 0&, ByVal 0&)
            CopyMemory result, address, 4
            GetLibraryScriptName = result
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
  
Sub TestLoadLibrary() ' Проверка загрузки библиотеки
    LoadLib
    ShowInfo
    
    'PassCell 1, 1, Cells(1, 1).value, Cells(1, 1).FormulaR1C1
    'PassCell 2, 1, Cells(2, 1).value, Cells(2, 1).FormulaR1C1
    'PassCell 3, 1, Cells(3, 1).value, Cells(3, 1).FormulaR1C1
    
    'MsgBox "Заполнено ячеек: " & GetFilledCells
    'MsgBox "1:1 = " & GetCellValue(1, 1) & vbNewLine & GetCellFormula(1, 1)
    'MsgBox "2:1 = " & GetCellValue(2, 1) & vbNewLine & GetCellFormula(2, 1)
    'MsgBox "3:1 = " & GetCellValue(3, 1) & vbNewLine & GetCellFormula(3, 1)
    'MsgBox GetBounds(0) & ":" & GetBounds(2) & " - " & GetBounds(1) & ":" & GetBounds(3)
    'MsgBox "Очищено ячеек: " & ClearData
    
    UnloadLib ' Выгрузка библиотеки
End Sub

Sub RunScript_GetLengthBetween2Points()
' Вызов скрипта GetLengthBetween2Points из библиотеки
'   Получение расстояние между двумя точками (ячейками): 55.555555,33.333333 (широта,долгота)
'   В ячейках должны быть географические координаты: 55.555555,33.333333 (широта,долгота)
' Обработка выбранных ячеек
' Может быть выбран диапазон из трех столбцов - в третий столбец будет внесен результат между 1-ым и 2-ым стоблцами
' Либо из двух столбцов, во второй столбец будет внесен результат из двух левых ячеек (текущая - предыдущая)

    Dim dlgRes As Integer
    Dim cel As Range
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    Dim pCo As Integer
    
    If selectedRange.Columns.Count < 2 Or selectedRange.Columns.Count > 3 Then
        MsgBox "Необходимо выбрать диапазон из двух или трех столбцов" & vbNewLine & "В правый столбец будет выведено расстояние в метрах"
        Exit Sub
    End If
    
    If selectedRange.Columns.Count = 2 And selectedRange.Rows.Count < 2 Then
        MsgBox "Дли диапазона из двух столбцов необходимо выбрать как минимум две строки" & vbNewLine & "В правый столбец будет выведено расстояние в метрах"
        Exit Sub
    End If
    
    
    pCo = selectedRange.Rows.Count
    dlgRes = MsgBox("Запустить скрипт для " & pCo & " выбранных точек?", vbYesNo + vbQuestion, "GetLengthBetween2Points")
    If dlgRes = vbNo Then
        Exit Sub
    End If
    
    LoadLib
    ClearData
                
    For Each cel In selectedRange.Cells
        PassCell cel.row, cel.column, Cells(cel.row, cel.column).value, Cells(cel.row, cel.column).FormulaR1C1
    Next cel
    
    Dim res As Integer
    res = RunScript("GetLengthBetween2Points")
    
    If res < 0 Then
        MsgBox "Ошибка обработки данных", vbOKOnly + vbExclamation, "GetLengthBetween2Points"
        Exit Sub
    ElseIf res = 0 Then
        MsgBox "Обработано точек скриптом: " & res & vbNewLine, vbOKOnly + vbInformation, "SearchAddressOSM"
        Exit Sub
    End If
    
    
    dlgRes = MsgBox("Обработано точек скриптом: " & res & vbNewLine & "Результат это расстояние в метрах." & vbNewLine & "Заполнить правый столбец выбоки полученным результатом?", vbYesNo + vbQuestion, "GetLengthBetween2Points")
    If dlgRes = vbYes Then
        For Each cel In selectedRange.Cells
            Cells(cel.row, cel.column).FormulaR1C1 = GetCellFormula(cel.row, cel.column)
        Next cel
    End If
    
    'UnloadLib ' Выгрузка библиотеки
End Sub

Sub RunScript_SearchAddressOSM()
' Выхов скрипта SearchAddressOSM из библиотеки
'   Получение координат по адресу на основе данныз OSM (RU)
' Обработка выбранных ячеек
' Поиск координат по адресу в OSM
' Сохранение результата 55.555555,33.333333 в исходные ячейки
    
    Dim dlgRes As Integer
    Dim cel As Range
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    
    dlgRes = MsgBox("Запустить скрипт для " & selectedRange.Cells.Count & " выбранных ячеек?", vbYesNo + vbQuestion, "SearchAddressOSM")
    If dlgRes = vbNo Then
        Exit Sub
    End If

    LoadLib
    ClearData
                
    For Each cel In selectedRange.Cells
        PassCell cel.row, cel.column, Cells(cel.row, cel.column).value, Cells(cel.row, cel.column).FormulaR1C1
    Next cel
    
    Dim res As Integer
    res = RunScript("SearchAddressInOSM")
    
    If res < 0 Then
        MsgBox "Ошибка обработки данных", vbOKOnly + vbExclamation, "SearchAddressOSM"
        Exit Sub
    ElseIf res = 0 Then
        MsgBox "Обработано ячеек скриптом: " & res & vbNewLine, vbOKOnly + vbInformation, "SearchAddressOSM"
        Exit Sub
    End If
    
    
    dlgRes = MsgBox("Обработано ячеек скриптом: " & res & vbNewLine & "Заполнить выбранные ячейки полученным результатом?", vbYesNo + vbQuestion, "SearchAddressOSM")
    If dlgRes = vbYes Then
        For Each cel In selectedRange.Cells
            Cells(cel.row, cel.column).FormulaR1C1 = GetCellFormula(cel.row, cel.column)
        Next cel
    End If
    
    'UnloadLib ' Выгрузка библиотеки
End Sub












