using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharpTests
{
    public interface IFileSystem
    {
        bool FileExists(string fileName);
        DateTime GetCreationDate(string fileName);
    }
}
