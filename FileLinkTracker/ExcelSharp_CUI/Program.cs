using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExcelSharp_CUI
{
    class Program
    {
        static void Main(string[] args)
        {
            DirectoryInfo solutionDirectory = new DirectoryInfo("../../../.");
            Console.WriteLine(solutionDirectory.FullName);
            DateTime dt = new DateTime(2014, 9, 15);

            FileInfo[] solutionFiles = solutionDirectory.GetFiles();
            int len = solutionFiles.GetLength(0);

            Console.WriteLine(len);

            foreach (var file in solutionFiles)
            {
                Console.WriteLine(file.Name);
            }
            Console.ReadKey();
        }
    }
}
