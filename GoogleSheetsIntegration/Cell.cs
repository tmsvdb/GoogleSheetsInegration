using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beardiegames.GoogleSheetsIntegration
{
    public class Cell
    {
        // properties
        CellID id;
        string val;

        // Cell information
        public CellID Id { get { return id; } }
        public string Value { get { return val; } set { val = value; } }

        // public class methodes

        public Cell(CellID id, string value)
        {
            this.id = id;
            this.val = value;
        }

        public bool Compare(int row, int col)
        {
            return (id.colIndex == col && id.rowIndex == row);
        }
        public bool Compare(string idString)
        {
            return (id.ToStringFormat() == idString);
        }
        public bool CompareColumn(int col)
        {
            return id.colIndex == col;
        }
        public bool CompareRow(int row)
        {
            return id.rowIndex == row;
        }

        public void MoveUp()
        {
            //Console.WriteLine("Move cell UP from " + id.ToStringFormat() + " to " + new CellID(id.rowIndex - 1, id.colIndex).ToStringFormat());
            id = new CellID(id.rowIndex - 1, id.colIndex);
        }
        public void MoveDown()
        {
            //Console.WriteLine("Move cell DOWN from " + id.ToStringFormat() + " to " + new CellID(id.rowIndex + 1, id.colIndex).ToStringFormat());
            id = new CellID(id.rowIndex + 1, id.colIndex);
        }
        public void MoveLeft()
        {
            //Console.WriteLine("Move cell LEFT from " + id.ToStringFormat() + " to " + new CellID(id.rowIndex, id.colIndex - 1).ToStringFormat());
            id = new CellID(id.rowIndex, id.colIndex - 1);
        }
        public void MoveRight()
        {
            //Console.WriteLine("Move cell RIGHT from " + id.ToStringFormat() + " to " + new CellID(id.rowIndex, id.colIndex + 1).ToStringFormat());
            id = new CellID(id.rowIndex, id.colIndex + 1);
        }
    }

    public class CellID
    {
        // properties
        int row;
        int col;

        // class information
        public int rowIndex { get { return row; } }
        public int colIndex { get { return col; } }

        // public class methodes

        public CellID(int row, int col)
        {
            this.row = row;
            this.col = col;
        }

        public string ToStringFormat()
        {
            return NumberToColumn(col + 1) + (row + 1).ToString();
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

            while (number > -1)
            {
                int temp = (number - 1) % 26;

                output = Convert.ToChar(temp + 65) + output;

                if (number > 26) number = ((number - 1) / 26);
                else break;
            }

            return output;
        }
    }
}
