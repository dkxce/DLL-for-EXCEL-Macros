using System;
using System.Collections.Generic;
using System.Text;

namespace exceltools
{    
    public class ExcelData
    {
        private uint MinRow = 0;
        private uint MaxRow = 0;
        private uint MinCol = 0;
        private uint MaxCol = 0;

        public Dictionary<ulong, ValueFormula> data = new Dictionary<ulong, ValueFormula>();

        public int MinX { get { return (int)MinRow; } }
        public int MaxX { get { return (int)MaxRow; } }
        public int MinY { get { return (int)MinCol; } }
        public int MaxY { get { return (int)MaxCol; } }
        public int Width { get { return (int)(MaxCol == 0 ? 0 : MaxCol - MinCol + 1); } }
        public int Height { get { return (int)(MaxRow == 0 ? 0 : MaxRow - MinRow + 1); } }

        public void Clear()
        {
            data.Clear();
            MinRow = 0;
            MaxRow = 0;
            MinCol = 0;
            MaxCol = 0;
        }

        public int Count
        {
            get
            {
                return data.Count;
            }
        }

        public int[] GetBounds()
        {
            return new int[] { (int)MinRow, (int)MaxRow, (int)MinCol, (int)MaxCol };
        }

        public void Set(uint row, uint col, string value, string formula)
        {
            if (String.IsNullOrEmpty(value)) value = "";
            if (String.IsNullOrEmpty(formula)) formula = "";
            SetMinMax(row, col);
            ulong index = (((ulong)row) << 32) + (ulong) col;
            if (data.ContainsKey(index))
                data[index] = new ValueFormula(value, formula);
            else
                data.Add(index, new ValueFormula(value, formula));
        }

        public void Set(int row, int col, string formula) { Set((uint)row, (uint)col, formula); }
        public void Set(uint row, uint col, string formula)
        {
            if (String.IsNullOrEmpty(formula)) formula = "";
            SetMinMax(row, col);
            ulong index = (((ulong)row) << 32) + (ulong)col;
            if (data.ContainsKey(index))
                data[index] = new ValueFormula("", formula);
            else
                data.Add(index, new ValueFormula("", formula));
        }

        public void Set(int row, int col, ValueFormula formula) { Set((uint)row, (uint)col, formula); }
        public void Set(uint row, uint col, ValueFormula vf)
        {
            if (String.IsNullOrEmpty(vf.value)) vf.value = "";
            if (String.IsNullOrEmpty(vf.formula)) vf.formula = "";
            SetMinMax(row, col);
            ulong index = (((ulong)row) << 32) + (ulong)col;
            if (data.ContainsKey(index))
                data[index] = vf;
            else
                data.Add(index, vf);
        }

        public ValueFormula Get(int row, int col) { return Get((uint)row, (uint)col); }
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

    public class ValueFormula
    {
        public string value;
        public string formula;
        public ValueFormula(string value, string formula) { this.value =value; this.formula = formula;}
    }
}
