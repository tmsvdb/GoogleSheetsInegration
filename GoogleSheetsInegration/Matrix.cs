using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beardiegames.GoogleSheetsInegration
{
    // types
    public class Cell
    {
        // properties
        public CellID id;
        public string value;

        // public class methodes

        public Cell(CellID id, string value)
        {
            this.id = id;
            this.value = value;
        }

        public bool Compare(int row, int col)
        {
            return (id.colNumber == col && id.rowNumber == row);
        }
        public bool Compare(string idString)
        {
            return (id.ToStringFormat() == idString);
        }
    }

    public class CellID 
    {
        // properties
        int row;
        int col;

        // class information
        public int rowNumber { get { return row; } }
        public int colNumber { get { return col; } }

        // public class methodes

        public CellID (int row, int col)
        {
            this.row = row;
            this.col = col;
        }

        public string ToStringFormat()
        {
            return NumberToColumn(col) + row.ToString();
        }


        // Static Class Features

        public static CellID FromStringFormat(string str)
        {
            string s_row = "";
            string s_col = "";

            for (int i = 0; i < str.Length; i++)
            {
                if (Char.IsNumber(str[i]))
                    s_row += str[i];
                else
                    s_col += str[i];
            }

            int col = ColumnToNumber(s_col);
            int row = int.Parse(s_row);

            return new CellID(row, col);
        }

        public static int ColumnToNumber(string name)
        {
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
        }

        public static string NumberToColumn(int number)
        {
            string output = "";

            while (number != 0)
            {
                int temp = number % 26;

                output += Convert.ToChar(temp + 64);

                if (number > 26)
                    number = number / 26;
                else
                    break;
            }

            return output;
        }
    }

    public class Matrix
    {
        // properties
        List<Cell> cells;
        int width, height;

        // public class info
        public int numberOfCells { get { return cells.Count; } }

        // Public Methodes

        public Matrix (int width, int height)
        {
            cells = new List<Cell>();

            this.width = width;
            this.height = height;

            for (int row = 0; row < height; row++)
                for (int col = 0; col < width; col++)
                    cells.Add(new Cell(new CellID(row, col), ""));
        }

        public string GetValue (int row, int col)
        {
            foreach (Cell cell in cells)
                if (cell.Compare(row, col))
                    return cell.value;

            return null;
        }
        public string GetValue(string idString)
        {
            foreach (Cell cell in cells)
                if (cell.Compare(idString))
                    return cell.value;

            return null;
        }

        public void SetValue(int row, int col, string value)
        {
            foreach (Cell cell in cells)
                if (cell.Compare(row, col))
                    cell.value = value;
        }
        public void SetValue(string idString, string value)
        {
            foreach (Cell cell in cells)
                if (cell.Compare(idString))
                    cell.value = value;
        }

        public void AddRow ()
        {
            height++;
            for (int col = 0; col < width; col++)
                cells.Add(new Cell(new CellID(height-1, col), ""));
        }

        public void AddColumn()
        {
            width++;
            for (int row = 0; row < height; row++)
                cells.Add(new Cell(new CellID(row, width-1), ""));
        }

        public void RemoveRow()
        {
            // TODO: Remove last row
        }
        public void RemoveRow(int row)
        {
            // TODO: Remove row at index
        }

        public void RemoveColumn()
        {
            // TODO: Remove last col
        }
        public void RemoveColumn(int col)
        {
            // TODO: Remove col at index
        }

        // Static Class Features

        public static Matrix FromObjectList()
        {
            return new Matrix(0, 0);
        }

        public static IList<IList<object>> ToObjectList()
        {
            return null;
        }
    }

    
}
