using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSolution.Code
{
    public interface IExcelApp
    {
        bool IsOpen { get; }

        void InitializeExcel();
        void CloseExcel();
    }
}
