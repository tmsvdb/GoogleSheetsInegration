using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beardiegames.GoogleSheetsIntegration
{
    public class PageRange
    {
        string title;
        int from_row;
        int from_col;
        int to_row;
        int to_col;

        public string peekTitle { get { return title; } }
        public int peekStartRow { get { return from_row; } }
        public int peekStartColumn { get { return from_col; } }
        public int peekEndRow { get { return to_row; } }
        public int peekEndColumn { get { return to_col; } }

        public PageRange(string title, int from_row, int from_col, int to_row, int to_col)
        {
            this.title = title;
            this.from_row = from_row;
            this.from_col = from_col;
            this.to_row = to_row;
            this.to_col = to_col;
        }

        public string ToRangeString()
        {
            string format = string.Format(
                "{0}!{1}:{2}",
                title,
                new CellID(from_row, from_col).ToStringFormat(),
                new CellID(to_row, to_col).ToStringFormat()
                );

            return format;
        }

        public static PageRange EmptyRange()
        {
            return new PageRange("", 0, 0, 0, 0);
        }
    }

    public class Page
    {
        // properties
        List<Cell> cells;
        int width, height, startRow, startCol;
        string title;

        // public class info
        public string peekTitle { get { return title; } }
        public int numberOfCells { get { return cells.Count; } }
        public int peekWidth { get { return width; } }
        public int peekHeight { get { return height; } }
        public List<Cell> peekCells { get { return cells; } }

        // Public Methodes

        public Page (string title, int width, int height, int startRow = 0, int startCol = 0)
        {
            cells = new List<Cell>();

            this.title = title;
            this.width = width;
            this.height = height;
            this.startRow = startRow;
            this.startCol = startCol;

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

        public PageRange ToPageRange()
        {
            return new PageRange(title, startRow, startCol, startRow + height - 1, startCol + width - 1);
        }

        // Static Class Features

        public static Page FromObjectList(string pageTitle, IList<IList<Object>> objList)
        {
            if (objList == null)
                throw new Exception("param objList = NULL!");

            if (objList.Count == 0 || objList[0].Count == 0)
                return new Page(pageTitle, 0, 0);

            Page mtx = new Page(pageTitle, objList[0].Count, objList.Count);

            for (int row = 0; row < objList.Count; row++)
                for (int col = 0; col < objList[0].Count; col++)
                    mtx.SetValue(row, col, objList[row][col].ToString());

            return mtx;
        }

        public static IList<IList<Object>> ToObjectList(Page mtx)
        {
            IList<IList<Object>> objList = new List<IList<Object>>();
            for (int row = 0; row < mtx.peekHeight; row++)
            {
                IList<Object> colList = new List<Object>();
                for (int col = 0; col < mtx.peekWidth; col++)
                    colList.Add(null);
                objList.Add(colList);
            }

            foreach (Cell cell in mtx.peekCells)
                objList[cell.Id.rowIndex][cell.Id.colIndex] = cell.Value;

            return objList;
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
