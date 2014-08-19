using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSharp
{
    abstract public class Sheet
    {
        public int Index { get { return worksheet.Index; } }        
        protected SheetWriter sheetWriter;       
        private Excel._Worksheet worksheet;
        private Excel.Range startRange;
        private Excel.Range endRange;
        private Excel._Workbook workbook;
        private SheetWriter nullWriter;
        
        public Sheet(Excel._Worksheet worksheet )
        {
            this.worksheet = worksheet;
            InitializeWriter();
        }
        

        /// <summary>
            /// Returns the block of cell values
            /// in an excel spreadsheet. The range starts at the top left corner
            /// and ends at the cell specefied by the entered
            /// row and column parameters.
            /// </summary>
            /// <param name="rows"></param>
            /// <param name="columns"></param>
            /// <returns> </returns>
        public string[,] GetCells(int rows, int columns)
        {
            string[,] cellArray = new string[rows, columns];
            assignCellArray(rows, columns, cellArray);
            return cellArray;
        }
        #region GetCells Helper Methods
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

        private void assignCell(string[,] cellArray, int r, int c)
        {
            cellArray[r, c] = (string)worksheet.Cells[r + 1, c + 1].Value;
        } 
        #endregion

        public string[,] GetRange(string startCell, string endCell)
        {
            string[,] cellArray = setCellArray(startCell, endCell);
            assignCellBlock(cellArray);
            return cellArray;
        }
        #region GetRange Helper Methods
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

        private string[,] setCellArray(string startCell, string endCell)
        {
            this.startRange = worksheet.get_Range(startCell);
            this.endRange = worksheet.get_Range(endCell);
            string[,] cellRange = new string[1 + endRange.Row - startRange.Row, 1 + endRange.Column - startRange.Column];
            return cellRange;
        } 
        #endregion

        abstract protected void InitializeWriter();
    }
}
