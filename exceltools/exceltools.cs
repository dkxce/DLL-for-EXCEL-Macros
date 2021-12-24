//      DO NOT BUILD `AnyCPU`
// in AnyCPU build DLL will not work!
//   !!! Build only x86 or x64 !!!
//
// Author: Milok Zbrozek <milokz@gmail.com>
//
// Это библиотека C#, которая экспортирует функции для макросов в Excel
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
        private const string LIBNAME = "Excel Tools by Milok Zbrozek <milokz@gmail.com>";
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
                "GetLibraryScriptName" // 11
            };

        private static string[] scripts = new string[]
            {
                "SearchAddressInOSM", // 0                
                "GetLengthBetween2Points" // 1
            };



        private static ExcelData exd = new ExcelData();

        [DllExport("GetLibraryName", CallingConvention = CallingConvention.Cdecl)]
        [return: MarshalAs(UnmanagedType.AnsiBStr)]
        static string GetLibraryName()
        {
            return LIBNAME;
        }

        [DllExport("GetLibraryMethods", CallingConvention = CallingConvention.Cdecl)]
        static int GetLibraryMethods()
        {
            return methods.Length;
        }

        [DllExport("GetLibraryMethodName", CallingConvention = CallingConvention.Cdecl)]
        [return: MarshalAs(UnmanagedType.AnsiBStr)]
        static string GetLibraryMethodName(int numberFromZero)
        {
            if (numberFromZero < 0) return "";
            if (numberFromZero >= methods.Length) return "";
            return methods[numberFromZero];
        }

        [DllExport("GetLibraryScripts", CallingConvention = CallingConvention.Cdecl)]
        static int GetLibraryScripts()
        {
            return scripts.Length;
        }

        [DllExport("GetLibraryScriptName", CallingConvention = CallingConvention.Cdecl)]
        [return: MarshalAs(UnmanagedType.AnsiBStr)]
        static string GetLibraryScriptName(int numberFromZero)
        {
            if (numberFromZero < 0) return "";
            if (numberFromZero >= scripts.Length) return "";
            return scripts[numberFromZero];
        }

        [DllExport("PassCell", CallingConvention = CallingConvention.Cdecl)]
        static int PassCell(int row, int col, [MarshalAs(UnmanagedType.AnsiBStr)] string value, [MarshalAs(UnmanagedType.AnsiBStr)] string formula)
        {
            if (row < 1) return -1;
            if (col < 1) return -1;
            exd.Set((uint)row, (uint)col, value, formula);
            return 0;
        }

        [DllExport("GetCellValue", CallingConvention = CallingConvention.Cdecl)]
        [return: MarshalAs(UnmanagedType.AnsiBStr)]
        static string GetCellValue(int row, int col)
        {
            if (row < 1) return "";
            if (col < 1) return "";
            return exd.Get((uint)row, (uint)col).value;
        }

        [DllExport("GetCellFormula", CallingConvention = CallingConvention.Cdecl)]
        [return: MarshalAs(UnmanagedType.AnsiBStr)]
        static string GetCellFormula(int row, int col)
        {
            if (row < 1) return "";
            if (col < 1) return "";
            return exd.Get((uint)row, (uint)col).formula;
        }

        [DllExport("ClearData", CallingConvention = CallingConvention.Cdecl)]
        static int ClearData()
        {
            int res = exd.Count;
            exd.Clear();
            return res;
        }

        [DllExport("GetFilledCells", CallingConvention = CallingConvention.Cdecl)]
        static int GetFilledCells()
        {
            return exd.Count;
        }

        [DllExport("GetBounds", CallingConvention = CallingConvention.Cdecl)]
        static int GetBounds(int index)
        {
            if (index < 0) return 0;
            if (index >= 4) return 0;
            return exd.GetBounds()[index];
        }

        [DllExport("RunScript", CallingConvention = CallingConvention.Cdecl)]
        static int RunScript([MarshalAs(UnmanagedType.AnsiBStr)] string methodName)
        {
            if (String.IsNullOrEmpty(methodName)) return 0;
            if (methodName == "SearchAddressInOSM") return UTILS_OSM.SearchAddressInOSM(ref exd);
            if (methodName == "GetLengthBetween2Points") return UTILS_DKXCE.GetLengthBetween2Points(ref exd);
            return 0;
        }

        #endregion
    }

    public static class UTILS_OSM
    {
        public static int SearchAddressInOSM(ref ExcelData exd)
        {
            KMZRebuilder.WaitingBoxForm wbf = new KMZRebuilder.WaitingBoxForm();
            wbf.Show("Script: SearchAddressInOSM", "Загрузка");
            wbf.Text = String.Format("Обработано ячеек {0}/{1}", 0, exd.data.Count);

            int processed = 0;
            bool nodialogs = false;
            foreach (KeyValuePair<ulong, ValueFormula> item in exd.data)
            {
                if (0 < processed++) System.Threading.Thread.Sleep(300); // avoid server exception
                string val = item.Value.formula;
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
                        item.Value.formula = val;
                };
                wbf.Text = String.Format("Обработано ячеек {0}/{1}", processed, exd.data.Count);
                wbf.Percent = (float)(100.0 * (float)processed / (float)exd.data.Count);                
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
        public static int GetLengthBetween2Points(ref ExcelData exd)
        {
            if ((exd.Width < 2) || (exd.Width > 3))
            {
                MessageBox.Show("Число столбцов может быть 2 или 3!\r\nВ правый столбец будет выведено расстояние в метрах!");
                return 0;
            };
            if (exd.Width == 2) return GetLengthBetween2Points2(ref exd);
            if (exd.Width == 3) return GetLengthBetween2Points3(ref exd);
            return 0;
        }

        private static int GetLengthBetween2Points2(ref ExcelData exd)
        {
            KMZRebuilder.WaitingBoxForm wbf = new KMZRebuilder.WaitingBoxForm();
            wbf.Show("Script: GetLengthBetween2Points", "Загрузка");
            wbf.Text = String.Format("Обработано точек {0}/{1}", 0, exd.data.Count);

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

                wbf.Text = String.Format("Обработано точек {0}/{1}", processed, exd.data.Count);
                wbf.Percent = (float)(100.0 * (float)processed / (float)exd.data.Count);
            };

            wbf.Hide();
            wbf = null;
            return processed;
        }

        private static int GetLengthBetween2Points3(ref ExcelData exd)
        {
            KMZRebuilder.WaitingBoxForm wbf = new KMZRebuilder.WaitingBoxForm();
            wbf.Show("Script: GetLengthBetween2Points", "Загрузка");
            wbf.Text = String.Format("Обработано точек {0}/{1}", 0, exd.data.Count);

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

                wbf.Text = String.Format("Обработано точек {0}/{1}", processed, exd.data.Count);
                wbf.Percent = (float)(100.0 * (float)processed / (float)exd.data.Count);
            };            

            wbf.Hide();
            wbf = null;
            return processed;
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