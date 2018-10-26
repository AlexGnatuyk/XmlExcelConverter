using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Spire.Xls;

namespace ExcelToXlfConverter
{
    class Program
    {
        public static List<TranslateData> ReadExcel(string fileName)
        {
            Workbook wb = new Workbook();
            wb.LoadFromFile(fileName);
            var sh = wb.Worksheets[0];
            var excelData = new List<TranslateData>();
            int i = 0;
            foreach (var row in sh.Range.Rows)
            {
                ++i;
                if (i == 1)
                    continue;

                excelData.Add(new TranslateData(row.Cells[0].Value, row.Cells[1].Value, row.Cells[2].Value));
            }
            return excelData;
        }

        public static XDocument CreateXmlDocument(List<TranslateData> data)
        {
            XNamespace ns = "urn:oasis:names:tc:xliff:document:2.0";
            var doc = new XDocument(new XElement(ns + "xliff",
                new XAttribute("version", "2.0"),
                new XElement("file",
                    new XAttribute("original", "ng.template"),
                    new XAttribute("id", "ngi18n"),
                    from row in data
                    select new XElement("unit",
                        new XAttribute("id", row.Id),
                        new XElement("notes",
                            new XElement("note",
                                new XAttribute("category", "description"), row.Description
                            )),
                        new XElement("segment",
                            new XElement("source", row.Source)))
                )
            ));
            return doc;
        }

        static void Main(string[] args)
        {
            if (args.Length == 2)
            {
                var excel = ReadExcel(args[0]);
                var result = CreateXmlDocument(excel);
                result.Save(args[1]);
            }
        }
    }

    [Serializable]
    public class TranslateData
    {
        public string Id;
        public string Source;
        public string Description;

        public TranslateData(string id, string source, string description)
        {
            Id = id;
            Source = source;
            Description = description;
        }
    }
}
