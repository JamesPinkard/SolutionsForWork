using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ExcelSharp;

namespace ExcelSharpTests
{
    [Category("Change Command")]
    [TestFixture]
    public class ChangeCommandTests : BaseTest
    {
        protected OfficeCommand testCommand;

        [Test]
        public void ChangeToolCommand_ChangeToLinkTool_ReceiverHasLinkTool()
        {
            testCommand = new ChangeToolCommand(testSheet, new LinkSheetTool());
            
            testCommand.Execute();

            Assert.That(testSheet.EmbedTool, Is.InstanceOf<LinkSheetTool>());
        }

        [Test]
        public void ChangeWriterCommand_ChangeToLinkSheetWriter_ReceiverHasLinkSheetWriter()
        {
            testCommand = new ChangeWriterCommand(testSheet, new LinkSheetWriter());

            testCommand.Execute();

            Assert.That(testSheet.Writer, Is.InstanceOf<LinkSheetWriter>());
        }          

        [Test]
        public void ChangeToDirectoryFormatter_Execute_ReceiverHasDirectoryForm()
        {
            testCommand = new ChangeToDirectoryFormatter(testSheet);

            testCommand.Execute();

            Assert.That(testSheet.Formatter, Is.InstanceOf<DirectoryFormatter>());
        }

        [Test]
        public void ChangeToLinkRemover_Execute_RecieverHasLinkRemover()
        {
            testCommand = new ChangeToLinkRemover(testSheet);

            testCommand.Execute();

            Assert.That(testSheet.RemoveTool, Is.InstanceOf<LinkRemover>());
        }
    }
}
