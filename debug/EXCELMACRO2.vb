'
'
' DLL Calling Macro Script for Excel
' Author: Milok Zbrozek <milokz@gmail.com>
'
' Скрипт обращения к внешней библиотеки и
'  запуска из нее скриптов для Ms Excel
'
'

#If win64 Then
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
    Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
    Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpFunc As LongPtr, ByVal Param1 As LongPtr, ByVal Param2 As LongPtr, ByVal Param3 As LongPtr, ByVal Param4 As LongPtr) As LongPtr

    Private Const x64 = True
    Private Const DLL_NAME = "exceltools_x64.dll"
    Dim handle As LongPtr, address As LongPtr, Unload As LongPtr
#Else
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpFunc As Long, ByVal Param1 As Long, ByVal Param2 As Long, ByVal Param3 As Long, ByVal Param4 As Long) As Long
    
    Private Const x64 = False
    Private Const DLL_NAME = "exceltools.dll"
    Dim handle As Long, address As Long, Unload As Long
#End If

Function GetLibPath() As String
'Полный путь к exceltools.dll
    Dim nPath As String
    nPath = ActiveWorkbook.FullName
    nPath = Left(nPath, Len(nPath) - Len(Split(nPath, "\")(UBound(Split(nPath, "\"))))) & DLL_NAME
    GetLibPath = nPath
End Function

 
Function LoadLib()
'ЗАГРУЗКА DLL
    If handle = 0 Then
        handle = LoadLibrary(GetLibPath)
        If handle <> 0 Then
            'MsgBox ("Библиотека успешно загружена.")
        ElseIf handle = 0 Then
            MsgBox ("Ошибка при загрузке DLL!")
        End If
    End If
End Function


Function UnloadLib()
'ВЫГРУЗКА DLL
    If handle <> 0 Then
        Unload = FreeLibrary(handle)
        If Unload <> 0 Then
            'MsgBox ("Библиотека выгружена успешно.")
        ElseIf Unload = 0 Then
            MsgBox ("Ошибка при выгрузке библиотеки из памяти!")
        End If
    End If
End Function
 
Function ShowInfo()
'ИНФОРМАЦИЯ О DLL
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
            libMNames = libMNames & GetLibraryMethodName(i) & " - " & CInt(i) & "; "
        Next i
        
        For i = 0 To libScripts - 1
            libSNames = libSNames & vbNewLine & " - " & GetLibraryScriptName(i) & " - " & CInt(i)
        Next i
        
        msginfo = "Загружена библиотека `" & DLL_NAME & "`: " & vbNewLine & "  " & libName & vbNewLine & vbNewLine & "Методов в библиотеке (" & libMethods & "): " & vbNewLine & "  " & libMNames
        msginfo = msginfo & vbNewLine & vbNewLine & "Скриптов в библиотеке (" & libScripts & "):" & libSNames
        MsgBox msginfo, vbOKOnly + vbInformation, "Microsoft Excel - " & DLL_NAME
    End If
End Function

Function GetPredefindedLengthString() As String
    GetPredefindedLengthString = ""
    Dim i As Integer
    For i = 1 To 6500
        GetPredefindedLengthString = GetPredefindedLengthString & "0000000000"
    Next i
End Function

 
Function GetLibName() As String
' Method 0 - Внутреннее имя библиотеки
    GetLibName = ""
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetLibraryName") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            Dim result As String
            result = GetPredefindedLengthString
            address = CallWindowProc(address, ByVal StrPtr(result), ByVal 0&, ByVal 0&, ByVal 0&)
            GetLibName = Left(result, CInt(address))
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function

Function GetLibraryMethods() As Integer
' Method 1 - Число поддерживаемых методов библиотеки
    GetLibraryMethods = 0
    Dim result As String
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetLibraryMethods") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            GetLibraryMethods = CInt(address)
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function

Function GetLibraryMethodName(ByVal num As Integer) As String
' Method 2 - Имя метода в библиотеки из списка по индексу (нумерация с нуля, всего: GetLibraryMethods)
'      num - индекс метода
    GetLibraryMethodName = ""
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetLibraryMethodName") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            Dim result As String
            result = GetPredefindedLengthString
            address = CallWindowProc(address, ByVal StrPtr(result), ByVal num, ByVal 0&, ByVal 0&)
            GetLibraryMethodName = Left(result, CInt(address))
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function


Function PassCell(Row As Long, column As Long, value As String, formula As String) As Integer
' Method 3 - ЗАПОЛНЕНИЕ ЯЧЕЙКИ В БИБЛИОТЕКЕ
'      Row - строка
'   column - столбец
'    value - текст
'  formula - формула
    PassCell = -1
    If handle <> 0 Then
        address = GetProcAddress(handle, "PassCell") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal Row, ByVal column, StrPtr(value), StrPtr(formula))
            PassCell = CInt(address)
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function

Function GetCellValue(Row As Long, column As Long) As String
' Method 4 - ПОЛУЧЕНИЕ ЗНАЧЕНИЕ ЯЧЕЙКИ ИЗ БИБЛИОТЕКИ
'      Row - строка
'   column - столбец

    GetCellValue = ""
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetCellValue") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            Dim result As String
            result = GetPredefindedLengthString
            address = CallWindowProc(address, ByVal Row, ByVal column, ByVal StrPtr(result), ByVal 0&)
            GetCellValue = Left(result, CInt(address))
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function

Function GetCellFormula(Row As Long, column As Long) As String
' Method 5 - ПОЛУЧЕНИЕ ЗНАЧЕНИЕ ФОРМУЛЫ ИЗ БИБЛИОТЕКИ
'      Row - строка
'   column - столбец
    GetCellFormula = ""
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetCellFormula") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            Dim result As String
            result = GetPredefindedLengthString
            address = CallWindowProc(address, ByVal Row, ByVal column, ByVal StrPtr(result), ByVal 0&)
            GetCellFormula = Left(result, CInt(address))
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
 
Function ClearData() As Integer
' Method 6 - Очистка всех ячеек в библиотеке
    ClearData = -1
    If handle <> 0 Then
        address = GetProcAddress(handle, "ClearData") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            ClearData = CInt(address)
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
 
Function GetFilledCells() As Integer
' Method 7 - Число заполненных ячеек в библиотеке
    GetFilledCells = -1
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetFilledCells") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            GetFilledCells = CInt(address)
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
 
Function GetBounds(index As Integer) As String
' Method 8 - Диапазон заполненных ячеек в библиотеке (границы)
'    index - значение из списка: 0 - MinRow, 1 - MaxRow, 2 - MinCol, 3 - MaxCol
    GetBounds = ""
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetBounds") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal index, ByVal 0&, ByVal 0&, ByVal 0&)
            GetBounds = CInt(address)
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function

Function RunScript(script As String) As Integer
' Method 9 - Запуск скрипта из библиотеки под именем
'   script - наименование скрипта
    RunScript = -1
    If handle <> 0 Then
        address = GetProcAddress(handle, "RunScript") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, StrPtr(script), ByVal 0&, ByVal 0&, ByVal 0&)
            RunScript = CInt(address)
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
 
Function GetLibraryScripts() As Integer
' Method 10 - Число поддерживаемых скриптов библиотеки
    GetLibraryScripts = 0
    Dim result As String
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetLibraryScripts") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            GetLibraryScripts = CInt(address)
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function

Function GetLibraryScriptName(ByVal num As Integer) As String
' Method 11 - Имя скрипта в библиотеки из списка по индексу (нумерация с нуля, всего: GetLibraryScripts)
'       num - индекс скрипта
    GetLibraryScriptName = ""
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetLibraryScriptName") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            Dim result As String
            result = GetPredefindedLengthString
            address = CallWindowProc(address, ByVal StrPtr(result), ByVal num, ByVal 0&, ByVal 0&)
            GetLibraryScriptName = Left(result, CInt(address))
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
  
Function GetChangedCells() As Integer
' Method 12 - Число измененных ячеек в библиотеке (последних или скриптом)
    GetChangedCells = -1
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetChangedCells") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            GetChangedCells = CInt(address)
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
    
Function GetFilledCellByNum(ByVal index As Integer, Row As Integer, Col As Integer) As String
' Method 13 - Получение формулы и адреса заполненной ячейки по индексу (нумерация с еуля, всего: GetFilledCells)
'     index - индекс заполенной ячейки
'       Row - возращаемый номер строки
'       Col - возращаемый номер столбца
    GetFilledCellByNum = ""
    #If win64 Then
        Dim addr0 As LongPtr
        Dim addr1 As LongPtr
    #Else
        Dim addr0 As Long
        Dim addr1 As Long
    #End If
    Dim rc(0 To 1) As Long
    addr0 = VarPtr(rc(0))
    addr1 = VarPtr(rc(1))
    
    Dim X As Integer
    Dim Y As Integer
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetFilledCellByNum") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            Dim result As String
            result = GetPredefindedLengthString
            address = CallWindowProc(address, ByVal index, ByVal addr1 - addr0, ByVal addr0, ByVal StrPtr(result))
            GetFilledCellByNum = Left(result, CInt(address))
            Row = CInt(rc(0))
            Col = CInt(rc(1))
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
  
Function GetChangedCellByNum(ByVal index As Integer, Row As Integer, Col As Integer) As String
' Method 14 - Получение формулы и адреса измененной ячейки по индексу (нумерация с нуля, всего: GetChangedCells)
'     index - индекс заполенной ячейки
'       Row - возращаемый номер строки
'       Col - возращаемый номер столбца
    GetChangedCellByNum = ""
     #If win64 Then
        Dim addr0 As LongPtr
        Dim addr1 As LongPtr
    #Else
        Dim addr0 As Long
        Dim addr1 As Long
    #End If
    Dim rc(0 To 1) As Long
    addr0 = VarPtr(rc(0))
    addr1 = VarPtr(rc(1))
    
    Dim X As Integer
    Dim Y As Integer
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetChangedCellByNum") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            Dim result As String
            result = GetPredefindedLengthString
            address = CallWindowProc(address, ByVal index, ByVal addr1 - addr0, ByVal addr0, ByVal StrPtr(result))
            GetChangedCellByNum = Left(result, CInt(address))
            Row = CInt(rc(0))
            Col = CInt(rc(1))
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
   
Function SelectAndRunScript() As Integer
' Method 15 - Выбор и запуск скрипта из DLL
    SelectAndRunScript = -1
    If handle <> 0 Then
        address = GetProcAddress(handle, "SelectAndRunScript") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)
            SelectAndRunScript = CInt(address)
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function


Function GetMinMax(ByRef MinRow As Integer, ByRef MaxRow As Integer, ByRef MinCol As Integer, ByRef MaxCol As Integer)
' Method 16 - Получение границ заполненных ячеек в библиотеке
'    MinRow - миниальный номер строки
'    MaxRow - максимальный номер строки
'    MinCol - миниальный номер столбца
'    MaxCol - максимальный номер столбца
     #If win64 Then
        Dim addr0 As LongPtr
        Dim addr1 As LongPtr
    #Else
        Dim addr0 As Long
        Dim addr1 As Long
    #End If
    Dim rc(0 To 4) As Long ' 0 - MinRow, 1 - MaxRow, 2 - MinCol, 3 - MaxCol
    addr0 = VarPtr(rc(0))
    addr1 = VarPtr(rc(1))
    
    Dim X As Integer
    Dim Y As Integer
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetMinMax") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, ByVal addr0, ByVal addr1 - addr0, ByVal 0&, ByVal 0&)
            If address = 0 Then
                MinRow = CInt(rc(0))
                MaxRow = CInt(rc(1))
                MinCol = CInt(rc(2))
                MaxCol = CInt(rc(3))
            End If
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
  
Function GetFilledRange() As String
' Method 17 - Получение диапазона заполненных ячеек (A1:B2)
    GetFilledRange = ""
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetFilledRange") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            Dim result As String
            result = GetPredefindedLengthString
            address = CallWindowProc(address, ByVal StrPtr(result), ByVal 0&, ByVal 0&, ByVal 0&)
            GetFilledRange = Left(result, CInt(address))
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
  
Function GetChangedRange() As String
' Method 18 - Получение диапазона измененных ячеек (A1:B2)
    GetChangedRange = ""
    If handle <> 0 Then
        address = GetProcAddress(handle, "GetChangedRange") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            Dim result As String
            result = GetPredefindedLengthString
            address = CallWindowProc(address, ByVal StrPtr(result), ByVal 0&, ByVal 0&, ByVal 0&)
            GetChangedRange = Left(result, CInt(address))
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
  
Function SetExcelFileName() As Integer
' Method 19 - Установка имени файла документа Excel
    SetExcelFileName = -1
    If handle <> 0 Then
        address = GetProcAddress(handle, "SetExcelFileName") ' получаем адрес функции
        If address <> 0 Then ' успешное получение адреса
            address = CallWindowProc(address, StrPtr(ActiveWorkbook.FullName), ByVal 0&, ByVal 0&, ByVal 0&)
            SetExcelFileName = CInt(address)
        ElseIf address = 0 Then ' ошибка при получении адреса
            Exit Function
        End If
    End If
End Function
  
Sub test__Load__Library()
' Проверка загрузки библиотеки
    LoadLib ' Загрузка библиотеки
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

Sub test__DLL__Methods()
    Dim X As Integer
    Dim Y As Integer
    Dim R As String

    LoadLib
    MsgBox "Method 0: " & GetLibName
    MsgBox "Method 1: " & GetLibraryMethods
    MsgBox "Method 2: " & GetLibraryMethodName(0)
    ClearData
    PassCell 1, 1, "A", "B" ' Method 3
    MsgBox "Method 3,4: " & GetCellValue(1, 1) ' Method 3,4
    MsgBox "Method 3,5: " & GetCellFormula(1, 1) ' Method 3,5
    MsgBox "Method 7: " & GetFilledCells ' Method 7
    MsgBox "Method 6: " & ClearData ' Method 6
    PassCell 1, 2, "ValB1", "ForB1" ' Method 3
    PassCell 2, 3, "ValC2", "ForC2" ' Method 3
    MsgBox "Method 8: " & GetBounds(0) & ":" & GetBounds(2) & " - " & GetBounds(1) & ":" & GetBounds(3) ' Method 8
    MsgBox "Method 9: " & RunScript("SearchAddressInOSM") ' Method 9
    MsgBox "Method 10: " & GetLibraryScripts
    MsgBox "Method 11: " & GetLibraryScriptName(0)
    MsgBox "Method 12: " & GetChangedCells
    R = GetFilledCellByNum(0, X, Y)  ' Method 13
    MsgBox "Method 13: " & X & "-" & Y & ": " & R  ' Method 13
    R = GetChangedCellByNum(1, X, Y)  ' Method 14
    MsgBox "Method 14: " & X & "-" & Y & ": " & R  ' Method 14
    SelectAndRunScript ' Method 15
    Dim Bounds(0 To 3) As Integer ' Method 16
    GetMinMax Bounds(0), Bounds(1), Bounds(2), Bounds(3) ' Method 16
    MsgBox "Method 16: " & Bounds(0) & " - " & Bounds(1) & " - " & Bounds(2) & " - " & Bounds(3) ' Method 16
    MsgBox "Method 17: " & GetFilledRange ' Method 17
    MsgBox "Method 18: " & GetChangedRange ' Method 18
End Sub

Sub Select_And_Run_Script_from_DLL()
' Передача выделенных ячеек в библиотеку, выбор и запуск скрипта из библиотеки

    Dim dlgRes As Integer
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    
    dlgRes = MsgBox("Загрузить в библиотеку " & selectedRange.Cells.Count & " выбранных ячеек и запустить выбор скрипта?", vbYesNo + vbQuestion, "Запуск скрипта из DLL")
    If dlgRes = vbNo Then
        Exit Sub
    End If
    
    LoadLib ' Загрузка библиотеки
    SetExcelFileName '  Установка имени файла документа
    ClearData ' Очистка ячеек в библиотеке
               
    ' Запись значений из выбранных ячеек в библиотеку
    Dim cel As Range
    For Each cel In selectedRange.Cells
        PassCell cel.Row, cel.column, Cells(cel.Row, cel.column).value, Cells(cel.Row, cel.column).FormulaR1C1
    Next cel
    
    ' Диапазон заполненных ячеек в библиотеке
    Dim dRange As String
    dRange = GetFilledRange
        
    ' Выбор и запуск скрипта
    Dim res As Integer
    res = SelectAndRunScript
    
    ' Получение диапазона заполненных ячеек в библиотеке
    Dim Bounds(0 To 3) As Integer
    GetMinMax Bounds(0), Bounds(1), Bounds(2), Bounds(3)
    
    ' Если скрипт не вызван или вернул 0
    If res < 0 Then
        MsgBox "Ошибка обработки данных", vbOKOnly + vbExclamation, "Запуск скрипта из DLL"
        Exit Sub
    ElseIf res = 0 Then
        MsgBox "Обработано ячеек скриптом: " & res & vbNewLine, vbOKOnly + vbInformation, "Запуск скрипта из DLL"
        Exit Sub
    End If
        
    ' Число измененных скриптом ячеек в библиотеке
    Dim changed As Integer
    changed = GetChangedCells
    
    ' Диапазон измененных скриптом ячеек в библиотеке
    Dim pRange As String
    pRange = GetChangedRange
    
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim R As String
    dlgRes = MsgBox("Обработано ячеек скриптом: " & res & " (" & dRange & ")" & vbNewLine & "Изменено ячеек скриптом: " & changed & " (" & pRange & ")" & vbNewLine & vbNewLine & "Заполнить измененные ячейки полученным результатом?", vbYesNo + vbQuestion, "Запуск скрипта из DLL")
    If dlgRes = vbYes Then
        ' Заполнениее только тех ячеек в таблице, которые изменены скриптом
        For i = 0 To changed - 1
           R = GetChangedCellByNum(i, X, Y)
           Cells(X, Y).FormulaR1C1 = R
        Next i
    End If
    
    'UnloadLib ' Выгрузка библиотеки
End Sub


