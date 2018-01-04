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
            Console.WriteLine("Move cell UP from " + id.ToStringFormat() + " to " + new CellID(id.rowIndex - 1, id.colIndex).ToStringFormat());
            id = new CellID(id.rowIndex - 1, id.colIndex);
        }
        public void MoveDown()
        {
            Console.WriteLine("Move cell DOWN from " + id.ToStringFormat() + " to " + new CellID(id.rowIndex + 1, id.colIndex).ToStringFormat());
            id = new CellID(id.rowIndex + 1, id.colIndex);
        }
        public void MoveLeft()
        {
            Console.WriteLine("Move cell LEFT from " + id.ToStringFormat() + " to " + new CellID(id.rowIndex, id.colIndex - 1).ToStringFormat());
            id = new CellID(id.rowIndex, id.colIndex - 1);
        }
        public void MoveRight()
        {
            Console.WriteLine("Move cell RIGHT from " + id.ToStringFormat() + " to " + new CellID(id.rowIndex, id.colIndex + 1).ToStringFormat());
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

        public CellID (int row, int col)
        {
            this.row = row;
            this.col = col;
        }

        public string ToStringFormat()
        {
            return NumberToColumn(col+1) + (row+1).ToString();
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
                int temp = (number-1) % 26;

                output = Convert.ToChar(temp + 65) + output;

                if (number > 26) number = ((number-1) / 26);
                else break;
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
        public int peekWidth { get { return width; } }
        public int peekHeight { get { return height; } }

        // Public Methodes

        public Matrix (int width, int height)
        {
            cells = new List<Cell>();

            this.width = width;
            this.height = height;

            for (int row = 0; row < height; row++)
                for (int col = 0; col < width; col++)
                    AddCell(row, col, "");
        }

        public string GetValue (int row, int col)
        {
            foreach (Cell cell in cells)
                if (cell.Compare(row, col))
                    return cell.Value;

            return null;
        }
        public string GetValue(string idString)
        {
            foreach (Cell cell in cells)
                if (cell.Compare(idString))
                    return cell.Value;

            return null;
        }

        public void SetValue(int row, int col, string value)
        {
            foreach (Cell cell in cells)
                if (cell.Compare(row, col))
                {
                    cell.Value = value;
                    return;
                }
        }
        public void SetValue(string idString, string value)
        {
            foreach (Cell cell in cells)
                if (cell.Compare(idString))
                {
                    cell.Value = value;
                    return;
                }
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

        public void InsertRowAt(int row)
        {
            MoveRowsDown(row);
        }

        public void InsertColumnAt(int col)
        {
            MoveColsRight(col);
        }

        public void RemoveLastRow()
        {
            RemoveRowAt(height-1);
        }

        public void RemoveRowAt(int row)
        {
            MoveRowsUp(row+1);
        }

        public void RemoveLastColumn()
        {
            RemoveColumnAt(width-1);
        }

        public void RemoveColumnAt(int col)
        {
            MoveColsLeft(col+1);
        }

        public List<Cell> GetRow (int row)
        {
            List<Cell> cellsInRow = new List<Cell>();

            foreach (Cell cell in cells)
                if (cell.CompareRow(row))
                    cellsInRow.Add(cell);

            return cellsInRow;
        }

        public List<Cell> GetColumn (int col)
        {
            List<Cell> cellsInColumn = new List<Cell>();

            foreach (Cell cell in cells)
                if (cell.CompareColumn(col))
                    cellsInColumn.Add(cell);

            return cellsInColumn;
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

        // Private implementation

        private void AddCell (int row, int col, string value)
        {
            cells.Add(new Cell(new CellID(row, col), value));
        }

        private void MoveRowsUp (int fromRow)
        {
            // remove overwritten rows
            if (fromRow > 0)
                foreach (Cell cell in GetRow(fromRow - 1))
                    cells.Remove(cell);

            // or incase of all rows selected, remove the first row
            else
                foreach (Cell cell in GetRow(0))
                    cells.Remove(cell);

            // move rows up
            for (int row = fromRow; row < height; row++)
                foreach (Cell cell in GetRow(row))
                    cell.MoveUp();

            height--;
        }

        private void MoveRowsDown(int fromRow)
        {
            // move rows down
            for (int row = fromRow; row < height; row++)
                foreach (Cell cell in GetRow(row))
                    cell.MoveDown();

            height++;

            // add a new row at the selected start row
            for (int col = 0; col < width; col++)
                AddCell(fromRow, col, "");
        }

        private void MoveColsLeft(int fromCol)
        {
            // remove overwritten cols
            if (fromCol > 0)
                foreach (Cell cell in GetColumn(fromCol - 1))
                    cells.Remove(cell);
            // or incase of all cols selected, remove the first col
            else
                foreach (Cell cell in GetColumn(0))
                    cells.Remove(cell);

            // move cols left
            for (int col = fromCol; col < width; col++)
                foreach (Cell cell in GetColumn(col))
                    cell.MoveLeft();

            width--;
        }

        private void MoveColsRight(int fromCol)
        {
            // move cols right
            for (int col = fromCol; col < width; col++)
                foreach (Cell cell in GetColumn(col))
                    cell.MoveRight();

            width++;

            // add a new col at the selected start col
            for (int row = 0; row < height; row++)
                AddCell(row, fromCol, "");
        }
    }
}
