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
    public abstract class BaseSheetFactoryTests : BaseTest
    {
        protected abstract SheetFactory testFactory { get; set; }
        protected void setUpTableSheetFactory()
        {
            testFactory = getTestFactory();            
        }
        protected abstract SheetFactory getTestFactory();
        protected abstract Type writerType { get; }
        protected abstract Type toolType { get; }
        
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
            get { return typeof(TableWriter); }
        }
        protected override Type toolType
        {
            get { return typeof(TableEmbedder); }
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
            get { return typeof(LinkWriter); }
        }

        protected override Type toolType
        {
            get { return typeof(LinkEmbedder); }
        }

        protected override SheetFactory getTestFactory()
        {
            return new LinkSheetFactory(testWorkbook);
        }
    }
}
