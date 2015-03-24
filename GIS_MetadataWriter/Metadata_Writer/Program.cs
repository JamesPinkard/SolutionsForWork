using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Metadata_Writer
{
    class Program
    {
        static void Main(string[] args)
        {
            Dictionary<string,string> fieldDict = new Dictionary<string,string>()
            {
                {"TE", "The top elevation, at the centroid of the cell, for {0}. (feet above mean sea level)"},
                {"BE", "The bottom elevation, at the centroid of the cell, for {0}. (feet above mean sea level)"},
                {"TH", "The layer thickness, at the centroid of the cell, for {0}. (feet)"},
                {"KH", "The hydraulic conductivity, at the centroid of the cell, for {0}. Calculated by dividing transmissivity by layer thickness. (feet per day)"},
                {"TR", "The transmissivity, at the centroid of the cell, for {0}. (square feet per day)"},
                {"IB", "The code number to designate whether the {0} is active or inactive.[Inactive, 0; Active, 1]"}
            };

            Dictionary<string,string> aqDict = new Dictionary<string,string>()
            {                
                {"LYR02", "Woodbine Aquifer"},
                {"LYR03", "Fredericksburg/Washita Group"},
                {"LYR04", "Paluxy Aquifer"},
                {"LYR05", "Glen Rose Aquifer"},
                {"LYR06", "Hensell Aquifer"},
                {"LYR07", "Pearsall Aquifer"},
                {"LYR08", "Lower Wilcox Aquifer"}
            };

            string defSource = "Developed by INTERA Inc. in 2015 based on Northern Trinity GAM; Kelley, V. A., Deeds, N. E., Fryar, D. G., and Nicot, J-P, 2004.  Groundwater Availability Models for the Queen City and Sparta Aquifers,  prepared for the Texas Water Development Board,  prepared by INTERA, Austin, Texas";

            XElement xmlFile = XElement.Load(@"C:\OneDrive - Intera Inc-\Briefcase\Tarrant-Wichita\XML_Examples\NTGM_metadata.xml");

            XElement xmlFields = xmlFile.Element("eainfo");
            XElement xmlDetails = xmlFields.Element("detailed");

            var query = from e in xmlDetails.Elements()
                        where e.Name == "attr"
                        select e;

            foreach (XElement field in query)
            {
                XElement labelElement = field.Element("attrlabl");
                string labelValue = labelElement.Value;
                if (labelValue.Contains("_LYR"))
                {
                    string theme = labelValue.Substring(0, 2);
                    string aq = labelValue.Substring(3, labelValue.Length - 3);
                    Console.WriteLine("Label:{0} Theme: {1} Aq: {2}", labelValue, theme, aq);
                    
                    string templateString = fieldDict[theme];
                    field.SetElementValue("attrdef", string.Format(templateString, aqDict[aq]));
                    field.SetElementValue("attrdefs", defSource);
                }
            }

            xmlFile.Save(@"C:\OneDrive - Intera Inc-\Briefcase\Tarrant-Wichita\XML_Examples\NTGM_metadata_updated.xml");
            Console.ReadLine();
        }
    }
}
