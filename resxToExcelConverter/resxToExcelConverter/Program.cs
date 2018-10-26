using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using OfficeOpenXml;

namespace resxToExcelConverter
{
    class Program
    {
        public static List<I18NElement> ReadResx(string resxFile)
        {
            var xml = new XmlDocument();
            xml.Load(resxFile);
            var dict = new Dictionary<string, string>();
            var listElements = new List<I18NElement>();
            foreach (XmlElement dataElement in xml.GetElementsByTagName("unit"))
            {
                var idValue = dataElement.GetAttribute("id");

                if (string.IsNullOrEmpty(idValue))
                {
                    Console.Error.WriteLine("Found <data> with empty name");
                }

                var segmentElement = dataElement.ChildNodes.OfType<XmlElement>()
                    .SingleOrDefault(e => e.Name == "segment");

                if (segmentElement != null)
                {
                    var sourceElement = segmentElement.ChildNodes.OfType<XmlElement>()
                        .SingleOrDefault(e => e.Name == "source");
                    if (sourceElement == null)
                    {
                        Console.Error.WriteLine("Found <data name='{0}'> without <source>", idValue);
                        continue;
                    }

                    var innerText = sourceElement.InnerXml;
                    var notesElement = dataElement.ChildNodes.OfType<XmlElement>()
                        .SingleOrDefault(e => e.Name == "notes");
                    var noteElement = notesElement?.ChildNodes.OfType<XmlElement>()
                        .Where(e => e.GetAttribute("category") == "description" ||
                                    e.GetAttribute("category") == "meaning")
                        .ToList();

                    string descriptionText = "";

                    if (noteElement != null && noteElement.Any())
                    {
                        var test = noteElement.SingleOrDefault(e => e.GetAttribute("category") == "meaning");
                        if (test != null) descriptionText = test.InnerText;

                        if (descriptionText == "")
                            descriptionText = noteElement.SingleOrDefault(e => e.GetAttribute("category") == "description")
                                ?.InnerText;
                    }

                    listElements.Add(new I18NElement(idValue, innerText, descriptionText));
                    dict.Add(idValue, innerText);
                }
            }
            return listElements;
        }

        static void WriteToSpreadsheet(string filename, List<I18NElement> records)
        {
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }
            var outFile = new FileInfo(filename);

            var pkg = new ExcelPackage(outFile);
            var ws = pkg.Workbook.Worksheets.Add("Texts");

            ws.Cells[1, 1].Value = "id";
            ws.Cells[1, 2].Value = "source";
            ws.Cells[1, 3].Value = "description";
            ws.Column(1).Width = 40;
            ws.Column(2).Width = 100;
            ws.Column(3).Width = 100;

            var i = 2;
            foreach (var record in records)
            {
                ws.Cells[i, 1].Value = record.Id;
                ws.Cells[i, 2].Value = record.Source;
                ws.Cells[i, 3].Value = record.Description;
                i++;
            }
            pkg.Save();
        }

        static void Main(string[] args)
        {
            if (args.Length == 2)
            {
                var result = ReadResx(args[0]);
                WriteToSpreadsheet(args[1], result);
            }
        }
    }

    class I18NElement
    {
        public string Id;
        public string Source;
        public string Description;

        public I18NElement(string id, string source, string description)
        {
            Id = id;
            Source = source;
            Description = description;
        }
    }
}
