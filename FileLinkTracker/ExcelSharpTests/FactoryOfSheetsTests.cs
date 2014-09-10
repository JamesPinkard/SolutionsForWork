using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ExcelSharp;
using System.IO;

namespace ExcelSharpTests
{   
    [TestFixture]
    public abstract class BaseSheetFactoryTests
    {
        protected ExcelOperator testOperator { get; set; }
        protected Workbook testWorkbook { get; set; }
        protected abstract SheetFactory testFactory { get; set; }
        protected abstract Type writerType { get; }
        protected abstract Type toolType { get; }

        [TestFixtureSetUp]
        public void InitializeWorksheet()
        {
            testOperator = new ExcelOperator();
            testOperator.InitializeExcel();
            testWorkbook = initTestWorkbook();
        }
        public Workbook initTestWorkbook()
        {
            string path = Environment.CurrentDirectory;
            string wbPath = @"{0}\TestingFiles\TestingWorkbook.xlsx";
            Workbook testWorkbook = testOperator.OpenWorkbook(string.Format(wbPath, path));
            return testWorkbook;
        }

        [TestFixtureTearDown]
        public void CloseExcel()
        {
            testOperator.CloseWorkbook();
            testOperator.CloseExcel();
        }

        protected void setUpTableSheetFactory()
        {
            testFactory = getTestFactory();            
        }
        protected abstract SheetFactory getTestFactory();  
        
        [Test]
        public void SheetFactory_ExecuteWithWorkbook_MakesNewSheet()
        {
            testFactory = getTestFactory();
            int initialSheetCount = testWorkbook.sheetCount;

            testFactory.ExecuteMake();
            
            Assert.That(testWorkbook.sheetCount,Is.EqualTo(initialSheetCount + 1));
        }
        
        [Test]
        public void SheetFactory_ExecuteWithWorkbook_NewSheetHasTableWriter()
        {            
            testFactory = getTestFactory();

            testFactory.ExecuteMake();

            Assert.That(testWorkbook.RecentlyAddedSheet.Writer,Is.InstanceOf(writerType));
        }

        [Test]
        public void SheetFactory_ExecuteWithWorkbook_NewSheetHasTableTools()
        {
            testFactory = getTestFactory();
            
            testFactory.ExecuteMake();

            Assert.That(testWorkbook.RecentlyAddedSheet.EmbedTool, Is.InstanceOf(toolType));
        }

        [Test]
        public void SheetFactory_ExecuteWithSheet_MakesNewSheet()
        {
            setUpTableSheetFactory();
            int initialSheetCount = testWorkbook.sheetCount;

            testFactory.ExecuteCopy();

            Assert.That(testWorkbook.sheetCount, Is.EqualTo(initialSheetCount + 1));
        }

        [Test]
        public void SheetFactory_ExecuteWithSheet_NewSheetHasTableWriter()
        {
            setUpTableSheetFactory();            

            testFactory.ExecuteCopy();

            Assert.That(testWorkbook.RecentlyAddedSheet.Writer, Is.InstanceOf(writerType));
        }

        [Test]
        public void SheetFactory_ExecuteWithSheet_NewSheetHasTableTools()
        {
            setUpTableSheetFactory();

            testFactory.ExecuteCopy();

            Assert.That(testWorkbook.RecentlyAddedSheet.EmbedTool, Is.InstanceOf(toolType));
        }

        // TODO Need to implement SheetHandler Class
        //      Needs to overwrite sheets, save, and open
    }

    [Category("Sheet Factory")]
    [TestFixture]
    public class TableSheetFactoryTests : BaseSheetFactoryTests
    {
        private TableSheetFactory testTableFactory;
        protected override SheetFactory testFactory
        {
            get { return testTableFactory; }
            set { testTableFactory = (TableSheetFactory) value ; }
        }
        protected override Type writerType
        {
            get { return typeof(TableSheetWriter); }
        }
        protected override Type toolType
        {
            get { return typeof(TableSheetTool); }
        }
       
        protected override SheetFactory getTestFactory()
        {
            return new TableSheetFactory(testWorkbook);
        }  
    }

    [Category("Sheet Factory")]
    [TestFixture]
    public class LinkSheetFactoryTests : BaseSheetFactoryTests
    {
        private LinkSheetFactory testLinkFactory;


        protected override SheetFactory testFactory
        {
            get
            {
                return testLinkFactory;
            }
            set
            {
                testLinkFactory = (LinkSheetFactory) value;
            }
        }

        protected override Type writerType
        {
            get { return typeof(LinkSheetWriter); }
        }

        protected override Type toolType
        {
            get { return typeof(LinkSheetTool); }
        }

        protected override SheetFactory getTestFactory()
        {
            return new LinkSheetFactory(testWorkbook);
        }
    }
}
