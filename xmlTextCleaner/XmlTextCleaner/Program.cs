using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace XmlTextCleaner
{
    class Program
    {
        public static class ResxReader
        {
            public static XDocument RemoveCertainNodes(string resxFile, string searchLine)
            {   var xdoc = XDocument.Load(resxFile);
                xdoc.XPathSelectElements($"root/data[contains(@name, '{searchLine}')]")?
                    .Remove(); 

                return xdoc;
            }
        }
        static void Main(string[] args)
        {
            if (args.Length == 2)
            {
                var result = ResxReader.RemoveCertainNodes(args[0], args[1]);
                result.Save("result.resx");
            }   
        }
    }
}
