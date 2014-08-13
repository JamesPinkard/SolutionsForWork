using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Web;
using System.Diagnostics;
using System.IO;
using System.Windows.Markup;

namespace FileLinkTracker
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            PopulateDocument();

            // Rig up some Click handlers for the save/load of the flow doc.
            btnSaveDoc.Click += (o, s) =>
            {
                using (FileStream fStream = File.Open("documentData.xaml", FileMode.Create))
                {
                    XamlWriter.Save(this.myDocumentReader.Document, fStream);                    
                }
            };

            btnLoadDoc.Click += (o, s) =>
            {
                using (FileStream fStream = File.Open("documentData.xaml", FileMode.Open))
                {
                    try
                    {
                        FlowDocument doc = XamlReader.Load(fStream) as FlowDocument;
                        this.myDocumentReader.Document = doc;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error Loading Doc!");
                    }
                }
            };
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private void copyLink(object sender, RoutedEventArgs e)
        {
            Run testLink = new Run("Test Hyperlink");
            Hyperlink myLink = new Hyperlink(testLink);
            myLink.NavigateUri = new Uri("http://search.msn.com");
            Clipboard.SetDataObject(myLink);
        }

        private void PopulateDocument()
        {
            Paragraph hlinkParagraph = new Paragraph();
            
            Run runLink = new Run("place holder");
            Hyperlink myLink = new Hyperlink(runLink);
            myLink.NavigateUri = new Uri("http://search.msn.com");

            hlinkParagraph.Inlines.Add(myLink);

            this.listOfHyperlinks.ListItems.Add(
            new ListItem(hlinkParagraph));
        }
    }
}
