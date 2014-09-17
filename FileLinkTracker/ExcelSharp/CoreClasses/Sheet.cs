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
        public string wbPath { get; private set; }
        public bool Exists
        {
            get
            {
                if (worksheet == null) { return false; }
                else { return true; }
            }            
        }
        public IRemover RemoveTool
        {
            get { return remover; }
            set
            {
                remover = value;
                remover.worksheet = this.worksheet;            
            }
        }
        public IOfficeWriter Writer
        {
            get { return writer; }
            set 
            {
                writer = value;
                writer.worksheet = this.worksheet;
            }
        }
        public IFormatter Formatter
        {
            get { return formatter; }
            set 
            {
                formatter = value;
                formatter.worksheet = this.worksheet;
            }
        }       
        public IEmbedder EmbedTool
        {
            get { return embedder; }
            set 
            {
                embedder = value;
                embedder.worksheet = this.worksheet;
            }
        }
        protected NullableTableWriter sheetWriter;

        public Sheet(Excel._Worksheet worksheet )
        {
            this.worksheet = worksheet;
            cellGetter = new CellGetter(worksheet);
        }
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
        public void DeleteSheet()
        {
            this.Index = 0;
            this.writer = null;
            this.embedder = null;
            this.worksheet.Delete();
            this.worksheet = null;
        }
        public void ClearCells()
        {
            worksheet.Cells.Clear();
        }
        public string[,] GetCells(int rows, int columns)
        {
            string [,] cells = cellGetter.GetCells(rows, columns);
            return cells;
        }
        public string[,] GetRange(string startCell, string endCell)
        {
            string[,] range = cellGetter.GetRange(startCell, endCell);
            return range;
        }
        public List<string> GetColumnRange(string column, int startCell, int endCell)
        {
            List<string> range = cellGetter.GetColumnRange(column, startCell, endCell);
            return range;
        }
        public string GetCell(int row, int column)
        {
            string value = cellGetter.GetCell(row, column);
            return value;
        }

        
        private CellGetter cellGetter;
                
        private Excel._Worksheet worksheet;
        private IEmbedder embedder;
        private IRemover remover;
        private IOfficeWriter writer;
        private IFormatter formatter;        
    }
}
