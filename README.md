# DLL Library for EXCEL Macros

C# DLL for use in Excel Macros 
There is two scripts in DLL and Excel file:
1) Find coordiantes (lat & lon) of the address (cell value is an address) with OSM RU AP
2) Find length in meters between two points (cell value is coordinates) in DLL method

This is a C# library & Excel VB Macro for Calling dll functions. No OLE or ActiveX is used.  
Can be used to create custom functions in external dll with C# and call it from Excel with Macros.   
Functions calls with cdecl using standard WinAPI. So you must export them in your C# code.   
Repository contains full working example. MSVS solutions, Excel file and Macro. 

[DLL & Excel Macro](https://github.com/dkxce/DLL-for-EXCEL-Macros/tree/main/debug)

## Библиотека внешний функций для вызова с помощью макросов в Excel

Данный скрипт (макрос) в файле Excel и библиотека могут:
1) Искать координаты для заданных в ячейках адресов и возвращать эти координаты как содержимое ячеек (по данным OSM RU)
2) Рассчитывать расстояние между двумя точками (точка это координаты в ячейке) по прямой

Примечание:
  - Для поиска координат: результат выозвращается в выделенные ячейки, т.е. адрес замещается координатами
  - Для расстояния: можно выделить диапазон из двух столбцов - в этом случае в правый столбец заносится расстояние между двумя левыми ячейками
  - Для расстояния: можно выделить диапазон из трех столбцов - в этом случае в правый (3-ий) столбец заносится расстояние между 1-ым и 2-ым столбцами

[Библиотека и файл с макросом](https://github.com/dkxce/DLL-for-EXCEL-Macros/tree/main/debug)
