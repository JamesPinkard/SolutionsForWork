using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSharp
{
    class CellGetter
    {
        Excel._Worksheet worksheet;
        public CellGetter(Excel._Worksheet worksheet)
        {
            this.worksheet = worksheet;
        }
        public string[,] GetCells(int rows, int columns)
        {
            worksheet.Activate();
            Excel.Range activeRange = worksheet.get_Range("A1");
            activeRange.Activate();
            string[,] cellArray = new string[rows, columns];
            assignCellArray(rows, columns, cellArray);
            return cellArray;
        }
        public string[,] GetRange(string startCell, string endCell)
        {
            worksheet.Activate();
            Excel.Range activeRange = worksheet.get_Range(startCell);
            activeRange.Activate();
            string[,] cellArray = setCellArray(startCell, endCell);
            assignCellBlock(cellArray);
            return cellArray;
        }
        public string GetCell(int row, int column)
        {
            object cellValue = (object)this.worksheet.Cells[row + 1, column + 1].Value;
            return cellValue.ToString();
        }
        public List<string> GetColumnRange(string column, int startCell, int endCell)
        {
            List<string> columnValues = new List<string>();

            for (int i = startCell; i <= endCell; i++)
            {
                string cellAddress = String.Format("{0}{1}", column, i);
                object cellValue = (object)worksheet.get_Range(cellAddress).Value;
                columnValues.Add(cellValue.ToString());
            }

            return columnValues;
        }


        private Excel.Range startRange;
        private Excel.Range endRange;
        private void assignCellArray(int rows, int columns, string[,] cellArray)
        {
            for (int r = 0; r < rows; r++)
            {
                assignRows(columns, cellArray, r);
            }
        }
        private void assignRows(int columns, string[,] cellArray, int r)
        {
            for (int c = 0; c < columns; c++)
            {
                assignCell(cellArray, r, c);
            }
        }
        private void assignCellBlock(string[,] cellArray)
        {
            for (int row = startRange.Row - 1; row < endRange.Row; row++)
            {
                assignRowBlocks(cellArray, row);
            }
        }
        private void assignRowBlocks(string[,] cellArray, int row)
        {
            for (int col = startRange.Column - 1; col < endRange.Column; col++)
            {
                assignCell(cellArray, row, col);
            }
        }
        private void assignCell(string[,] cellArray, int r, int c)
        {
            object cellValue = (object)this.worksheet.Cells[r + 1, c + 1].Value;
            cellArray[r, c] = (string)cellValue.ToString();
        }

        private string[,] setCellArray(string startCell, string endCell)
        {
            this.startRange = worksheet.get_Range(startCell);
            this.endRange = worksheet.get_Range(endCell);
            string[,] cellRange = new string[1 + endRange.Row - startRange.Row, 1 + endRange.Column - startRange.Column];
            return cellRange;
        }    
    }
}
