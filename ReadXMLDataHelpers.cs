using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace ExcelToXML
{
    internal class ReadXMLDataHelpers
    {
        public IEnumerable<string> ReadXMLData()
        {
            // Path to your XML file
           string filePath = Path.Combine(Environment.CurrentDirectory, @"TestData\", "xmluserinfo.xml");

            // Check if the file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine("XML File not found.");
                return null;
            }

            
            XElement xelm = XElement.Load(filePath);
            IEnumerable<string> xmlnames = xelm.Elements("user")
                                                 .Select(user => user.Element("name").Value)
                                                 .Where(value => value != null)
                                                 .ToList();
            return xmlnames;

           /*
            return Enumerable.Range(1, rowCount)
                            .Select(row => worksheet.Cells[$"{columnLetter}{row}"].Text)
                            .Where(value => !string.IsNullOrEmpty(value))
                            .ToList(); // Convert to list or return as IEnumerable

            */
            /*
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(filePath);
            XmlNodeList nodes = xdoc.SelectNodes("//userinfo/user");            
            foreach (XmlNode node in nodes)
            {
                XmlNode name = node.SelectSingleNode("name");
                if (name != null)
                {
                    Console.WriteLine($"{name.InnerText}");

                }


            }
            */

        }
    }
}
