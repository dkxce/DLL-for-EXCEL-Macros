//      DO NOT BUILD `AnyCPU`
// in AnyCPU build DLL will not work!
//   !!! Build only x86 or x64 !!!
//
// Author: Milok Zbrozek <milokz@gmail.com>
//
// Это библиотека C#, которая экспортирует функции для макросов в Excel
//
// VBA & DLL
//   http://basic.ucoz.net/publ/vyzov_funkcij_po_ukazatelju_v_visual_basic_chast_2/1-1-0-33
//

using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using RGiesecke.DllExport;

using System.Windows.Forms;

namespace exceltools
{
    internal static class UnmanagedExports
    {
        private static string LIBNAME = "Excel Tools by Milok Zbrozek <milokz@gmail.com> " + (IntPtr.Size == 4 ? "x86" : "x64") + " v21.12.28.9";
        private static string CALLER = "Unknown";

        #region __cdelc
        // C++    -- typedef void (__cdecl * TestFunc)();
        // python -- _cdll.Test()
        // delphi -- procedure Test(); cdecl; external 'UnmanagedExports.dll';
        [DllExport("Test", CallingConvention = CallingConvention.Cdecl)]
        static void Test()
        {
            
        }

        // C++    -- typedef char* (__cdecl * GetLibNameFunc)();
        // python -- c_char_p(_cdll.GetLibName())
        // delphi -- function GetLibName(): PChar; cdecl; external 'UnmanagedExports.dll';
        [DllExport("GetLibName", CallingConvention = CallingConvention.Cdecl)]
        static string GetLibName()
        {
          return LIBNAME;
        }        

        // C++    -- typedef void* (__cdecl * SetCallerNameFunc)(char*);
        // python -- _cdll.SetCallerName(c_char_p(b"Passed by Python"))
        // delphi -- procedure SetCallerName(str: PChar); cdecl; external 'UnmanagedExports.dll';
        [DllExport("SetCallerName", CallingConvention = CallingConvention.Cdecl)]
        static void SetCallerName(string name)
        {
            CALLER = String.IsNullOrEmpty(name) ? "Unknown" : name;
        }

        // C++    -- typedef char* (__cdecl * GetCallerNameFunc)();
        // python -- c_char_p(_cdll.GetCallerName())
        // delphi -- function GetCallerName(): PChar; cdecl; external 'UnmanagedExports.dll';
        [DllExport("GetCallerName", CallingConvention = CallingConvention.Cdecl)]
        static string GetCallerName()
        {
            return CALLER;
        }
        #endregion

        #region Excel Methods

        /// <summary>
        ///     Exported Method Names For Excel
        /// </summary>
        private static string[] methods = new string[]
            {
                "GetLibraryName", // 0
                "GetLibraryMethods", // 1
                "GetLibraryMethodName", // 2
                "PassCell", // 3
                "GetCellValue", // 4
                "GetCellFormula", //5
                "ClearData", // 6
                "GetFilledCells", // 7
                "GetBounds", // 8
                "RunScript", // 9
                "GetLibraryScripts", // 10
                "GetLibraryScriptName", // 11
                "GetChangedCells", // 12
                "GetFilledCellByNum", // 13
                "GetChangedCellByNum", // 14
                "SelectAndRunScript", // 15
                "GetMinMax", // 16
                "GetFilledRange", // 17
                "GetChangedRange", // 18
                "SetExcelFileName", //19
                "CallTextFunction" // 20
            };

        /// <summary>
        ///     Exported Script Names For Excel
        /// </summary>
        private static string[] scripts = new string[]
            {
                "SearchAddressInOSM", // 0                
                "GetLengthBetween2Points" // 1
            };

        /// <summary>
        ///     Excel Simple Active Sheet Model
        /// </summary>
        private static ExcelData exd = new ExcelData();       

        /// <summary>
        ///     Get Library Name
        /// </summary>
        /// <param name="ptr">Pointer To Unicode String (max len: 65000)</param>
        /// <returns>string length</returns>
        [DllExport("GetLibraryName", CallingConvention = CallingConvention.Cdecl)]
        public static int GetLibraryName(IntPtr ptr) // METHOD 0
        {
            byte[] b = System.Text.Encoding.Unicode.GetBytes(LIBNAME);
            Marshal.Copy(b, 0, ptr, b.Length);
            return LIBNAME.Length;
        }

        /// <summary>
        ///     Get Library Exported Methods Count For Excel
        /// </summary>
        /// <returns>Count</returns>
        [DllExport("GetLibraryMethods", CallingConvention = CallingConvention.Cdecl)]
        public static int GetLibraryMethods() // METHOD 1
        {
            return methods.Length;
        }

        /// <summary>
        ///     Get Exported Library Method Name For Excel By Number (zero-indexed)
        /// </summary>
        /// <param name="ptr">Pointer To Unicode String (max len: 65000)</param>
        /// <param name="numberFromZero">index (zero-indexed)</param>
        /// <returns>string length</returns>
        [DllExport("GetLibraryMethodName", CallingConvention = CallingConvention.Cdecl)]
        public static int GetLibraryMethodName(IntPtr ptr, int numberFromZero) // METHOD 2
        {
            if (numberFromZero < 0) return 0;
            if (numberFromZero >= methods.Length) return 0;
            byte[] b = System.Text.Encoding.Unicode.GetBytes(methods[numberFromZero]);
            Marshal.Copy(b, 0, ptr, b.Length);
            return methods[numberFromZero].Length;
        }

        /// <summary>
        ///     Fill Cell From Excel
        /// </summary>
        /// <param name="row">Row</param>
        /// <param name="col">Column</param>
        /// <param name="value">Text Value of Cell (Pointer To Unicode String (max len: 65000))</param>
        /// <param name="formula">FormulaR1C1 of Cell (Pointer To Unicode String (max len: 65000))</param>
        /// <returns>zero</returns>
        [DllExport("PassCell", CallingConvention = CallingConvention.Cdecl)]
        public static int PassCell(int row, int col, [MarshalAs(UnmanagedType.BStr)] string value, [MarshalAs(UnmanagedType.BStr)] string formula)  // METHOD 3
        {
            if (row < 1) return -1;
            if (col < 1) return -1;
            exd.Set((uint)row, (uint)col, value, formula);
            return 0;
        }

        /// <summary>
        ///     Get Text Value of DLL Sheet Cell
        /// </summary>
        /// <param name="row">Row</param>
        /// <param name="col">Column</param>
        /// <param name="ptr">Pointer To Unicode String (max len: 65000)</param>
        /// <returns>text length</returns>
        [DllExport("GetCellValue", CallingConvention = CallingConvention.Cdecl)]
        public static int GetCellValue(int row, int col, IntPtr ptr)  // METHOD 4
        {
            if (row < 1) return 0;
            if (col < 1) return 0;
            string val = exd.Get((uint)row, (uint)col).value;
            byte[] b = System.Text.Encoding.Unicode.GetBytes(val);
            Marshal.Copy(b, 0, ptr, b.Length);
            return val.Length;
        }

        /// <summary>
        ///     Get FormulaR1C1 of DLL Sheet Cell
        /// </summary>
        /// <param name="row">Row</param>
        /// <param name="col">Column</param>
        /// <param name="ptr">Pointer To Unicode String (max len: 65000)</param>
        /// <returns>text length</returns>
        [DllExport("GetCellFormula", CallingConvention = CallingConvention.Cdecl)]
        public static int GetCellFormula(int row, int col, IntPtr ptr)  // METHOD 5
        {
            if (row < 1) return 0;
            if (col < 1) return 0;
            string val = exd.Get((uint)row, (uint)col).formula;
            byte[] b = System.Text.Encoding.Unicode.GetBytes(val);
            Marshal.Copy(b, 0, ptr, b.Length);
            return val.Length;
        }

        /// <summary>
        ///     Clear Sheet Data in DLL
        /// </summary>
        /// <returns>Count of cleared cells</returns>
        [DllExport("ClearData", CallingConvention = CallingConvention.Cdecl)]
        public static int ClearData()  // METHOD 6
        {
            int res = exd.FilledCount;
            exd.Clear();
            return res;
        }

        /// <summary>
        ///     Get count of filled cells in DLL
        /// </summary>
        /// <returns>count of filled cells</returns>
        [DllExport("GetFilledCells", CallingConvention = CallingConvention.Cdecl)]
        public static int GetFilledCells()  // METHOD 7
        {
            return exd.FilledCount;
        }

        /// <summary>
        ///     Get Bounds of filled cells in DLL (0 - MinRow, 1 - MaxRow, 2 - MinCol, 3 - MaxCol)
        /// </summary>
        /// <param name="index">0 - MinRow, 1 - MaxRow, 2 - MinCol, 3 - MaxCol</param>
        /// <returns>value</returns>
        [DllExport("GetBounds", CallingConvention = CallingConvention.Cdecl)]
        public static int GetBounds(int index)   // METHOD 8
        {
            if (index < 0) return 0;
            if (index >= 4) return 0;
            return exd.GetFilledBounds()[index];
        }

        /// <summary>
        ///     Run Script from DLL with Name
        /// </summary>
        /// <param name="methodName">name of the script</param>
        /// <returns>count of computed cells</returns>
        [DllExport("RunScript", CallingConvention = CallingConvention.Cdecl)]
        public static int RunScript([MarshalAs(UnmanagedType.BStr)] string methodName)   // METHOD 9
        {
            return RunScriptInt(methodName);
        }

        /// <summary>
        ///     Get Library Excel Scripts Count
        /// </summary>
        /// <returns>count of library scripts for Excel</returns>
        [DllExport("GetLibraryScripts", CallingConvention = CallingConvention.Cdecl)]
        public static int GetLibraryScripts()  // METHOD 10
        {
            return scripts.Length;
        }

        /// <summary>
        ///     Get Library Excel Script Name by index (zero-indexed)
        /// </summary>
        /// <param name="ptr">Pointer To Unicode String (max len: 65000)</param>
        /// <param name="numberFromZero">script index (zero-indexed)</param>
        /// <returns>string length</returns>
        [DllExport("GetLibraryScriptName", CallingConvention = CallingConvention.Cdecl)]
        public static int GetLibraryScriptName(IntPtr ptr, int numberFromZero)  // METHOD 11
        {
            if (numberFromZero < 0) return 0;
            if (numberFromZero >= scripts.Length) return 0;
            byte[] b = System.Text.Encoding.Unicode.GetBytes(scripts[numberFromZero]);
            Marshal.Copy(b, 0, ptr, b.Length);
            return scripts[numberFromZero].Length;
        }

        /// <summary>
        ///     Get Last Changed Cells Count by Macro or Script
        /// </summary>
        /// <returns>count of changed cells</returns>
        [DllExport("GetChangedCells", CallingConvention = CallingConvention.Cdecl)]
        public static int GetChangedCells()   // METHOD 12
        {
            return exd.ChangedCount;
        }        

        /// <summary>
        ///     Get Filled Cell Row, Column & FormulaR1C1 by index (zero-indexed)
        /// </summary>
        /// <param name="index">index (zero-indexed)</param>
        /// <param name="rc">RCItem Struct or Array</param>
        /// <param name="elSize">size of RCItem / 2</param>
        /// <param name="ptrs">Pointer To Unicode String (max len: 65000)</param>
        /// <returns>string length</returns>
        [DllExport("GetFilledCellByNum", CallingConvention = CallingConvention.Cdecl)]
        public unsafe static int GetFilledCellByNum(int index, int elSize, IntPtr rc, IntPtr ptrs)   // METHOD 13
        {
            ExcelData.RCItem res = exd.GetFilled(index);
            if (res == null) res = new ExcelData.RCItem(0, 0);

            if (elSize == 4)
            {
                int* ptr = (int*)rc;
                *ptr = res.row;
                ptr++;
                *ptr = res.col;
            }
            else if (elSize == 8)
            {
                long* ptr = (long*)rc;
                *ptr = res.row;
                ptr++;
                *ptr = res.col;
            };
            string val = exd.Get(res.row, res.col).formula;
            byte[] b = System.Text.Encoding.Unicode.GetBytes(val);
            Marshal.Copy(b, 0, ptrs, b.Length);
            return val.Length;
        }

        /// <summary>
        ///     Get Last Changed Cells by Macro or Script: Cell Row, Column & FormulaR1C1 by index (zero-indexed)
        /// </summary>
        /// <param name="index">index (zero-indexed)</param>
        /// <param name="rc">RCItem Struct or Array</param>
        /// <param name="elSize">size of RCItem / 2</param>
        /// <param name="ptrs">Pointer To Unicode String (max len: 65000)</param>
        /// <returns>string length</returns>
        [DllExport("GetChangedCellByNum", CallingConvention = CallingConvention.Cdecl)]
        public unsafe static int GetChangedCellByNum(int index, int elSize, IntPtr rc, IntPtr ptrs)   // METHOD 14
        {
            ExcelData.RCItem res = exd.GetChanged(index);
            if (res == null) res = new ExcelData.RCItem(0, 0);

            if (elSize == 4)
            {
                int* ptr = (int*)rc;
                *ptr = res.row;
                ptr++;
                *ptr = res.col;
            }
            else if (elSize == 8)
            {
                long* ptr = (long*)rc;
                *ptr = res.row;
                ptr++;
                *ptr = res.col;
            };
            string val = exd.Get(res.row, res.col).formula;
            byte[] b = System.Text.Encoding.Unicode.GetBytes(val);
            Marshal.Copy(b, 0, ptrs, b.Length);
            return val.Length;
        }

        /// <summary>
        ///     Select Script And Run 
        /// </summary>
        /// <returns>count of computed cells</returns>
        [DllExport("SelectAndRunScript", CallingConvention = CallingConvention.Cdecl)]
        public static int SelectAndRunScript()   // METHOD 15
        {
            if (exd.FilledCount == 0)
            {
                MessageBox.Show("Необходимо выбрать хотя бы одну ячейку для выбора скрипта!", LIBNAME, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return 0;
            };

            int[] bounds = exd.GetChangedBounds();
            string dRange = String.Format("{0}{1}:{2}{3}", ExcelData.RCItem.ColumnIndex(bounds[2]), bounds[0], ExcelData.RCItem.ColumnIndex(bounds[3]), bounds[1]);
            
            int sel = 0;
            InputBox.defWidth = 700;
            DialogResult dr = InputBox.Show(LIBNAME, "Выберите скрипт для запуска (к обработке " + exd.FilledCount.ToString() + " ячеек в диапазоне " + dRange + ") - всего (" + scripts.Length.ToString() + "):", scripts, ref sel);
            if (dr != DialogResult.OK) return 0;
            return RunScriptInt(scripts[sel]);            
        }

        /// <summary>
        ///     Get Bounds of filled cells (0 - MinRow, 1 - MaxRow, 2 - MinCol, 3 - MaxCol)
        /// </summary>
        /// <param name="rect">0 - MinRow, 1 - MaxRow, 2 - MinCol, 3 - MaxCol</param>
        /// <returns>value</returns>
        [DllExport("GetMinMax", CallingConvention = CallingConvention.Cdecl)]
        public unsafe static int GetMinMax(IntPtr rect, int elSize) // METHOD 16
        {
            int[] bounds = exd.GetFilledBounds();
            if (elSize == 4)
            {
                int* ptr = (int*)rect;
                *ptr = bounds[0]; ptr++;
                *ptr = bounds[1]; ptr++;
                *ptr = bounds[2]; ptr++;
                *ptr = bounds[3];
            }
            else if (elSize == 8)
            {
                long* ptr = (long*)rect;
                *ptr = bounds[0]; ptr++;
                *ptr = bounds[1]; ptr++;
                *ptr = bounds[2]; ptr++;
                *ptr = bounds[3];
            };
            return 0;
        }

        /// <summary>
        ///     Get Filled Range
        /// </summary>
        /// <param name="ptr">Pointer To Unicode String (max len: 65000)</param>
        /// <returns>string length</returns>
        [DllExport("GetFilledRange", CallingConvention = CallingConvention.Cdecl)]
        public static int GetFilledRange(IntPtr ptr) // METHOD 17
        {
            int[] bounds = exd.GetFilledBounds();
            string val = String.Format("{0}{1}:{2}{3}", ExcelData.RCItem.ColumnIndex(bounds[2]), bounds[0], ExcelData.RCItem.ColumnIndex(bounds[3]), bounds[1]);
            byte[] b = System.Text.Encoding.Unicode.GetBytes(val);
            Marshal.Copy(b, 0, ptr, b.Length);
            return val.Length;
        }

        /// <summary>
        ///     Get Changed Range
        /// </summary>
        /// <param name="ptr">Pointer To Unicode String (max len: 65000)</param>
        /// <returns>string length</returns>
        [DllExport("GetChangedRange", CallingConvention = CallingConvention.Cdecl)]
        public static int GetChangedRange(IntPtr ptr) // METHOD 18
        {
            int[] bounds = exd.GetChangedBounds();
            string val = String.Format("{0}{1}:{2}{3}", ExcelData.RCItem.ColumnIndex(bounds[2]), bounds[0], ExcelData.RCItem.ColumnIndex(bounds[3]), bounds[1]);
            byte[] b = System.Text.Encoding.Unicode.GetBytes(val);
            Marshal.Copy(b, 0, ptr, b.Length);
            return val.Length;
        }

        /// <summary>
        ///     Set Excel File Name
        /// </summary>
        /// <param name="fileName">fileName</param>
        /// <returns>zero</returns>
        [DllExport("SetExcelFileName", CallingConvention = CallingConvention.Cdecl)]
        public static int SetExcelFileName([MarshalAs(UnmanagedType.BStr)] string fileName)  // METHOD 19
        {
            exd.FileName = fileName;
            return 0;
        }

        /// <summary>
        ///     Call Text Function From DLL (funcName:funcParam)
        /// </summary>
        /// <param name="QueryAndResult">Pointer To Unicode String (max len: 65000)</param>
        /// <param name="sLength">length of passed string in QueryAndResult</param>
        /// <returns>string length</returns>
        [DllExport("CallTextFunction", CallingConvention = CallingConvention.Cdecl)]
        public static int CallTextFunction(IntPtr QueryAndResult, int sLength) // METHOD 20
        {
            string val = "";
            if (QueryAndResult != IntPtr.Zero)
            {
                try
                {
                    string[] pv = Marshal.PtrToStringBSTR(QueryAndResult).Substring(0, sLength).Split(new char[] { ':' }, 2);
                    val = CallTextFunctionInt(pv[0], pv[1]);
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, LIBNAME, MessageBoxButtons.OK, MessageBoxIcon.Error); };
            };
            byte[] b = System.Text.Encoding.Unicode.GetBytes(val);
            Marshal.Copy(b, 0, QueryAndResult, b.Length);
            return val.Length;
        }


        private static int RunScriptInt(string methodName)
        {
            if (String.IsNullOrEmpty(methodName)) return 0;
            try
            {
                if (methodName == "SearchAddressInOSM") return UTILS_OSM.SearchAddressInOSM(ref exd);
                if (methodName == "GetLengthBetween2Points") return UTILS_DKXCE.GetLengthBetween2Points(ref exd);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), LIBNAME, MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            return 0;
        }

        private static string CallTextFunctionInt(string funcName, string paramValue)
        {
            if (funcName == "GetTime") return DateTime.Now.ToString("HH:mm:ss");
            if (funcName == "GetDate") return DateTime.Now.ToString("dd.MM.yyyy");
            if (funcName == "GetDateTime") return DateTime.Now.ToString("HH:mm:ss dd.MM.yyyy");
            if (funcName == "Random")
            {
                if (String.IsNullOrEmpty(paramValue))
                    return (new Random()).Next().ToString();
                else
                {
                    int max = 100;
                    int.TryParse(paramValue, out max);
                    return (new Random()).Next(max).ToString();
                };
            };
            return "";
        }

        #endregion
    }

    public static class UTILS_OSM
    {
        /// <summary>
        ///     Search Address with OSM
        /// </summary>
        /// <param name="exd">Excel Data</param>
        /// <returns>computed cells</returns>
        public static int SearchAddressInOSM(ref ExcelData exd)
        {
            int toRead = exd.FilledCount;
            if (toRead == 0) return 0;
            exd.ResetChanged();

            KMZRebuilder.WaitingBoxForm wbf = new KMZRebuilder.WaitingBoxForm();
            wbf.Show("Script: SearchAddressInOSM", "Загрузка");
            wbf.Text = String.Format("Обработано ячеек {0}/{1}", 0, exd.FilledCount);

            int processed = 0;
            bool nodialogs = false;            
            for (int i = 0; i < toRead; i++ )
            {
                ExcelData.RCItem item = exd.GetFilled(i);
                if (0 < processed++) System.Threading.Thread.Sleep(300); // avoid server exception
                string val = exd.Get(item.row, item.col).formula;
                if (!String.IsNullOrEmpty(val))
                {
                    try
                    {
                        val = SearchAddressInOSMInt(val, ref nodialogs, ref wbf);
                    }
                    catch (Exception ex)
                    {
                        val = ex.Message;
                    };
                    if (!String.IsNullOrEmpty(val))
                        exd.Set(item.row, item.col, val);
                };
                wbf.Text = String.Format("Обработано ячеек {0}/{1}", processed, exd.FilledCount);
                wbf.Percent = (float)(100.0 * (float)processed / (float)exd.FilledCount);
            };

            wbf.Hide();
            wbf = null;
            return processed;
        }

        private static string SearchAddressInOSMInt(string address, ref bool noDialogs, ref KMZRebuilder.WaitingBoxForm wbf)
        {
            string text = address.Trim();
            if (String.IsNullOrEmpty(text)) return "";
            {
                string result = null;
                try
                {
                    System.Net.HttpWebRequest wq = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(@"http://openstreetmap.ru/api/search?q=" + System.Security.SecurityElement.Escape(text));
                    System.Net.HttpWebResponse wr = (System.Net.HttpWebResponse)wq.GetResponse();
                    StreamReader sr = new StreamReader(wr.GetResponseStream(), System.Text.Encoding.ASCII);
                    string response = sr.ReadToEnd();
                    sr.Close();
                    wr.Close();

                    Regex rN = new Regex("\"display_name\":\\s\"(?<name>[^\"]*)", RegexOptions.IgnoreCase);
                    Regex rY = new Regex("\"lat\":\\s(?<lat>[\\d.]*)", RegexOptions.IgnoreCase);
                    Regex rX = new Regex("\"lon\":\\s(?<lon>[\\d.]*)", RegexOptions.IgnoreCase);

                    MatchCollection mN = rN.Matches(response);
                    MatchCollection mY = rY.Matches(response);
                    MatchCollection mX = rX.Matches(response);

                    int count = Math.Min(mN.Count, Math.Min(mY.Count, mX.Count));                    
                    if(count > 0) 
                    {
                        List<string[]> res = new List<string[]>();
                        for (int i = 0; i < count; i++)
                            res.Add(new string[] { Regex.Unescape(mN[i].Groups["name"].Value), mY[i].Groups["lat"].Value + "," + mX[i].Groups["lon"].Value });
                        int rr = 0;
                        if ((!noDialogs) && (res.Count > 1))
                        {
                            string[] toSel = new string[res.Count];
                            for (int i = 0; i < res.Count; i++) toSel[i] = res[i][0];
                            InputBox.defWidth = 700;
                            if (wbf != null) wbf.Hide();
                            DialogResult dr = InputBox.Show(text, "Выберите одно из значений", toSel, ref rr);
                            if (wbf != null) wbf.Show();
                            if (dr != DialogResult.OK) noDialogs = true;
                        };
                        return res[rr][1];
                    };                    
                }
                catch (Exception ex)
                {
                    return ex.Message;
                };
                if (result == null)
                {
                    return "Ничего не найдено";
                };               
                return result;
            };
        }
    }

    public static class UTILS_DKXCE
    {
        /// <summary>
        ///     Get Length between 2 points lat,lon
        /// </summary>
        /// <param name="exd">Excel Data</param>
        /// <returns>computed cells</returns>
        public static int GetLengthBetween2Points(ref ExcelData exd)
        {
            if ((exd.Width < 2) || (exd.Width > 3))
            {
                MessageBox.Show("Число столбцов может быть 2 или 3!\r\nВ правый столбец будет выведено расстояние в метрах!", "GetLengthBetween2Points");
                return 0;
            };
            exd.ResetChanged();
            if (exd.Width == 2) return GetLengthBetween2Points2(ref exd);
            if (exd.Width == 3) return GetLengthBetween2Points3(ref exd);
            return 0;
        }

        private static int GetLengthBetween2Points2(ref ExcelData exd)
        {
            KMZRebuilder.WaitingBoxForm wbf = new KMZRebuilder.WaitingBoxForm();
            wbf.Show("Script: GetLengthBetween2Points", "Загрузка");
            wbf.Text = String.Format("Обработано точек {0}/{1}", 0, exd.FilledCount);

            Regex ex = new Regex(@"^(?<lat>[+-]?[\d.]*),(?<lon>[+-]?[\d.]*)$");
            int processed = 1;
            exd.Set(exd.MinX, exd.MinY + 1, "0");
            for (int x = exd.MinX + 1; x <= exd.MaxX; x++)
            {
                processed++;
                string valA = exd.Get(x - 1, exd.MinY).formula;
                string valB = exd.Get(x, exd.MinY).formula;

                if ((!String.IsNullOrEmpty(valA)) && (!String.IsNullOrEmpty(valB)))
                {
                    Match mxA = ex.Match(valA);
                    Match mxB = ex.Match(valB);
                    if (mxA.Success && mxB.Success)
                    {
                        uint l = 0;
                        try
                        {
                            double aLat, aLon, bLat, bLon;
                            double.TryParse(mxA.Groups["lat"].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out aLat);
                            double.TryParse(mxA.Groups["lon"].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out aLon);
                            double.TryParse(mxB.Groups["lat"].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out bLat);
                            double.TryParse(mxB.Groups["lon"].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out bLon);
                            l = GetLengthMetersC(aLat, aLon, bLat, bLon, false);
                            exd.Set(x, exd.MinY + 1, l.ToString());
                        }
                        catch (Exception err)
                        {
                            exd.Set(x, exd.MinY + 1, err.Message);
                        };
                    }
                    else
                        exd.Set(x, exd.MinY + 1, "0");
                };

                wbf.Text = String.Format("Обработано точек {0}/{1}", processed, exd.FilledCount);
                wbf.Percent = (float)(100.0 * (float)processed / (float)exd.FilledCount);
            };

            wbf.Hide();
            wbf = null;
            return processed;
        }

        private static int GetLengthBetween2Points3(ref ExcelData exd)
        {
            KMZRebuilder.WaitingBoxForm wbf = new KMZRebuilder.WaitingBoxForm();
            wbf.Show("Script: GetLengthBetween2Points", "Загрузка");
            wbf.Text = String.Format("Обработано точек {0}/{1}", 0, exd.FilledCount);

            Regex ex = new Regex(@"^(?<lat>[+-]?[\d.]*),(?<lon>[+-]?[\d.]*)$");
            int processed = 0;
            for (int x = exd.MinX; x <= exd.MaxX; x++)
            {
                processed++;
                string valA = exd.Get(x, exd.MinY).formula;
                string valB = exd.Get(x, exd.MinY + 1).formula;

                if((!String.IsNullOrEmpty(valA)) && (!String.IsNullOrEmpty(valB)))
                {
                    Match mxA = ex.Match(valA);
                    Match mxB = ex.Match(valB);
                    if(mxA.Success && mxB.Success)
                    {
                        uint l = 0;
                        try
                        {
                            double aLat, aLon, bLat, bLon;
                            double.TryParse(mxA.Groups["lat"].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out aLat);
                            double.TryParse(mxA.Groups["lon"].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out aLon);
                            double.TryParse(mxB.Groups["lat"].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out bLat);
                            double.TryParse(mxB.Groups["lon"].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out bLon);
                            l = GetLengthMetersC(aLat, aLon, bLat, bLon, false);
                            exd.Set(x, exd.MinY + 2, l.ToString());
                        }
                        catch (Exception err)
                        {
                            exd.Set(x, exd.MinY + 2, err.Message);
                        };                        
                    };
                };

                wbf.Text = String.Format("Обработано точек {0}/{1}", processed, exd.FilledCount);
                wbf.Percent = (float)(100.0 * (float)processed / (float)exd.FilledCount);
            };            

            wbf.Hide();
            wbf = null;
            return processed * 2;
        }

        private static uint GetLengthMetersC(double StartLat, double StartLong, double EndLat, double EndLong, bool radians)
        {
            double D2R = Math.PI / 180;
            if (radians) D2R = 1;
            double dDistance = Double.MinValue;
            double dLat1InRad = StartLat * D2R;
            double dLong1InRad = StartLong * D2R;
            double dLat2InRad = EndLat * D2R;
            double dLong2InRad = EndLong * D2R;

            double dLongitude = dLong2InRad - dLong1InRad;
            double dLatitude = dLat2InRad - dLat1InRad;

            // Intermediate result a.
            double a = Math.Pow(Math.Sin(dLatitude / 2.0), 2.0) +
                       Math.Cos(dLat1InRad) * Math.Cos(dLat2InRad) *
                       Math.Pow(Math.Sin(dLongitude / 2.0), 2.0);

            // Intermediate result c (great circle distance in Radians).
            double c = 2.0 * Math.Atan2(Math.Sqrt(a), Math.Sqrt(1.0 - a));

            const double kEarthRadiusKms = 6378137.0000;
            dDistance = kEarthRadiusKms * c;

            return (uint)Math.Round(dDistance);
        }

    }
}