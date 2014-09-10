using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSharp
{
    public class Sheet : IOfficeCommandable
    {
        public int Index { get; set; }        
        public string Name { get { return worksheet.Name; } }
        public IEmbedder EmbedTool { get; set; }
        public IRemover RemoveTool { get; set; }
        public IOfficeWriter Writer{ get; set; }
        public IFormatter Formatter { get; set; }
        public string wbPath { get; private set; }
        
        private Excel._Worksheet worksheet;
        private Excel.Range startRange;
        private Excel.Range endRange;
        protected AbstractTableWriter sheetWriter;

        public bool isValid()
        {
            if (worksheet == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        public Sheet(Excel._Worksheet worksheet )
        {
            this.worksheet = worksheet;
        }
        public void DeleteSheet()
        {
            this.Index = 0;
            this.Writer = null;
            this.EmbedTool = null;
            this.worksheet.Delete();
            this.worksheet = null;
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
            cellArray[r, c] = (string)this.worksheet.Cells[r + 1, c + 1].Value;
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
        private string[,] setCellArray(string startCell, string endCell)
        {
            this.startRange = worksheet.get_Range(startCell);
            this.endRange = worksheet.get_Range(endCell);
            string[,] cellRange = new string[1 + endRange.Row - startRange.Row, 1 + endRange.Column - startRange.Column];
            return cellRange;
        } 






  


        public void ClearCells()
        {
            worksheet.Cells.Clear();
        }

    }
}
