using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ExcelSharp;

namespace ExcelSharpTests
{
    [Category("WriterTests")]
    [TestFixture]
    public class WriterTests : BaseTest
    {
        protected LinkWriter testWriter;
        
        [Test]
        public void LinkWriter_NonRecursiveWrite_SheetHasDirectoryFileHyperlinks()
        {

        }
    }
}
