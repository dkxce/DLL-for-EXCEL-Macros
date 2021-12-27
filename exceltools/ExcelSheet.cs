using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace exceltools
{    
    /// <summary>
    ///     Excel Simple Active Sheet Model
    /// </summary>
    public class ExcelData
    {
        /// <summary>
        ///     Cell Address RC
        /// </summary>
        public class RCItem
        {
            /// <summary>
            ///     Row
            /// </summary>
            public int row = 0;
            /// <summary>
            ///     Column
            /// </summary>
            public int col = 0;
            /// <summary>
            ///     Address (A1,B2...)
            /// </summary>
            public string Address { get { return String.Format("{0}:{1}", ColumnIndex(col), row); } }

            /// <summary>
            ///     Basic constructor
            /// </summary>
            public RCItem() { }

            /// <summary>
            ///     Constructor
            /// </summary>
            /// <param name="row">Row</param>
            /// <param name="col">Column</param>
            public RCItem(int row, int col) 
            { 
                this.row = row; 
                this.col = col;
            }

            /// <summary>
            ///     Constructor
            /// </summary>
            /// <param name="row">Row</param>
            /// <param name="col">Column (A,B...)</param>
            public RCItem(int row, string col)
            {
                this.row = row;
                this.col = ColumnIndex(col);
            }

            /// <summary>
            ///     Constructor
            /// </summary>
            /// <param name="reference">Address (A1,B2...)</param>
            public RCItem(/* A1, B2 ... */ string reference)
            {
                reference = reference.Trim();
                if (Char.IsDigit(reference[0]))
                {
                    string[] rc = reference.Split(new char[] { ':' });
                    row = int.Parse(rc[0]);
                    col = int.Parse(rc[1]);
                }
                else
                {
                    if (reference.IndexOf(":") > 0)
                    {
                        string[] cr = reference.Split(new char[] { ':' });
                        col = ColumnIndex(cr[0]);
                        row = int.Parse(cr[1]);
                    }
                    else
                    {
                        int sci = 0;
                        while (sci < reference.Length) if (char.IsDigit(reference[sci++])) break;
                        col = ColumnIndex(reference.Substring(0, sci));
                        row = int.Parse(reference.Substring(sci - 1));
                    };
                };
            }

            /// <summary>
            ///     Get Column Text Representation from Index
            /// </summary>
            /// <param name="reference">columnt text address</param>
            /// <returns>column index address</returns>
            public static int ColumnIndex(string reference)
            {
                int ci = 0;
                reference = reference.ToUpper();
                for (int ix = 0; ix < reference.Length && reference[ix] >= 'A'; ix++)
                    ci = (ci * 26) + ((int)reference[ix] - 64);
                return ci;
            }

            /// <summary>
            ///     Get Column Index Representation from Text
            /// </summary>
            /// <param name="reference">column index address</param>
            /// <returns>columnt text address</returns>
            public static string ColumnIndex(int reference)
            {
                string result = String.Empty;
                while (reference > 0)
                {
                    int mod = (reference - 1) % 26;
                    result = Convert.ToChar(65 + mod).ToString() + result;
                    reference = (int)((reference - mod) / 26);
                };
                return result;
            }
        }

        private uint MinRow = 0;
        private uint MaxRow = 0;
        private uint MinCol = 0;
        private uint MaxCol = 0;

        private Dictionary<ulong, ValueFormula> data = new Dictionary<ulong, ValueFormula>();
        private List<ulong> filled = new List<ulong>();
        private List<ulong> changed = new List<ulong>();

        public int MinX { get { return (int)MinRow; } }
        public int MaxX { get { return (int)MaxRow; } }
        public int MinY { get { return (int)MinCol; } }
        public int MaxY { get { return (int)MaxCol; } }
        public int Width { get { return (int)(MaxCol == 0 ? 0 : MaxCol - MinCol + 1); } }
        public int Height { get { return (int)(MaxRow == 0 ? 0 : MaxRow - MinRow + 1); } }

        /// <summary>
        ///     Clear Cells
        /// </summary>
        public void Clear()
        {
            data.Clear();
            filled.Clear();
            changed.Clear();
            MinRow = 0;
            MaxRow = 0;
            MinCol = 0;
            MaxCol = 0;            
        }

        /// <summary>
        ///     Reset Last Changes in Cells
        /// </summary>
        public void ResetChanged()
        {
            changed.Clear();
        }

        private void Change(ulong item)
        {
            if (!filled.Contains(item)) filled.Add(item);
            if (!changed.Contains(item)) changed.Add(item);
        }

        /// <summary>
        ///     Get Filled Cell Address by index (zero-indexed)
        /// </summary>
        /// <param name="index">index (zero-indexed)</param>
        /// <returns>RC value</returns>
        public RCItem GetFilled(int index)
        {
            if (index < 0) return null;
            if (index >= this.filled.Count) return null;
            ulong item = this.filled[index];
            int row = (int)((item >> 32) & 0xFFFFFFFF);
            int col = (int)(item & 0xFFFFFFFF);
            return new RCItem(row, col);
        }

        /// <summary>
        ///     Get Last Changed Cell Address by index (zero-indexed)
        /// </summary>
        /// <param name="index">index (zero-indexed)</param>
        /// <returns>RC value</returns>
        public RCItem GetChanged(int index)
        {
            if (index < 0) return null;
            if (index >= this.changed.Count) return null;
            ulong item = this.changed[index];
            int row = (int)((item >> 32) & 0xFFFFFFFF);
            int col = (int)(item & 0xFFFFFFFF);
            return new RCItem(row, col);
        }

        /// <summary>
        ///     Get Filled Cells Count
        /// </summary>
        public int FilledCount
        {
            get
            {
                return data.Count;
            }
        }

        /// <summary>
        ///     Get Last Changed Cells Count
        /// </summary>
        public int ChangedCount
        {
            get
            {
                return changed.Count;
            }
        }

        /// <summary>
        ///     Get Filled Cells Bounds (MinRow, MaxRow, MinCol, MaxCol)
        /// </summary>
        /// <returns></returns>
        public int[] GetFilledBounds()
        {
            return new int[] { (int)MinRow, (int)MaxRow, (int)MinCol, (int)MaxCol };
        }

        /// <summary>
        ///     Get Changed Cells Bounds (MinRow, MaxRow, MinCol, MaxCol)
        /// </summary>
        /// <returns></returns>
        public int[] GetChangedBounds()
        {
            if (this.changed.Count == 0) return new int[] { 0, 0, 0, 0 };
            int[] res = new int[] { int.MaxValue, int.MinValue, int.MaxValue, int.MinValue };
            foreach (long item in this.changed)
            {
                int row = (int)((item >> 32) & 0xFFFFFFFF);
                int col = (int)(item & 0xFFFFFFFF);
                if (row < res[0]) res[0] = row;
                if (row > res[1]) res[1] = row;
                if (col < res[2]) res[2] = col;
                if (col > res[3]) res[3] = col;
            };
            return res;
        }

        /// <summary>
        ///     Set Cell Value & Formula
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="value">value</param>
        /// <param name="formula">formula</param>
        public void Set(uint row, uint col, string value, string formula)
        {
            if (String.IsNullOrEmpty(value)) value = "";
            if (String.IsNullOrEmpty(formula)) formula = "";
            SetMinMax(row, col);
            ulong index = (((ulong)row) << 32) + (ulong) col;
            Change(index);
            if (data.ContainsKey(index))
                data[index] = new ValueFormula(value, formula);
            else
                data.Add(index, new ValueFormula(value, formula));
        }

        /// <summary>
        ///     Set Cell Formula
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="formula">formula</param>
        public void Set(int row, int col, string formula) { Set((uint)row, (uint)col, formula); }

        /// <summary>
        ///     Set Cell Formula
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="formula">formula</param>
        public void Set(uint row, uint col, string formula)
        {
            if (String.IsNullOrEmpty(formula)) formula = "";
            SetMinMax(row, col);
            ulong index = (((ulong)row) << 32) + (ulong)col;
            Change(index);
            if (data.ContainsKey(index))
                data[index] = new ValueFormula("", formula);
            else
                data.Add(index, new ValueFormula("", formula));
        }

        /// <summary>
        ///     Set Cell Value & Formula
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="ValueFormula">Value & Formula</param>
        public void Set(int row, int col, ValueFormula formula) { Set((uint)row, (uint)col, formula); }

        /// <summary>
        ///     Set Cell Value & Formula
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="ValueFormula">Value & Formula</param>
        public void Set(uint row, uint col, ValueFormula vf)
        {
            if (String.IsNullOrEmpty(vf.value)) vf.value = "";
            if (String.IsNullOrEmpty(vf.formula)) vf.formula = "";
            SetMinMax(row, col);
            ulong index = (((ulong)row) << 32) + (ulong)col;
            Change(index);
            if (data.ContainsKey(index))
                data[index] = vf;
            else
                data.Add(index, vf);
        }

        /// <summary>
        ///     Get Cell Value & Formula
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <returns>Value & Formula</returns>
        public ValueFormula Get(int row, int col) { return Get((uint)row, (uint)col); }

        /// <summary>
        ///     Get Cell Value & Formula
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <returns>Value & Formula</returns>
        public ValueFormula Get(uint row, uint col)
        {
            ulong index = (((ulong)row) << 32) + (ulong)col;
            if (data.ContainsKey(index))
                return data[index];
            else
                return new ValueFormula("", "");
        }

        private void SetMinMax(uint row, uint col)
        {
            if (MinRow == 0) MinRow = row;
            if (MaxRow == 0) MinRow = row;
            if (MinCol == 0) MinCol = col;
            if (MaxCol == 0) MinCol = col;

            if (row < MinRow) MinRow = row;
            if (row > MaxRow) MaxRow = row;
            if (col < MinCol) MinCol = col;
            if (col > MaxCol) MaxCol = col;
        }
    }

    /// <summary>
    ///     Cell & Formula
    /// </summary>
    public class ValueFormula
    {
        public string value;
        public string formula;
        public ValueFormula(string value, string formula) { this.value =value; this.formula = formula;}
    }
}
