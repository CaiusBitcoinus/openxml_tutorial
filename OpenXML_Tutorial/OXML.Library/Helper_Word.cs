using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OXML.Library
{
    public class Helper_Word
    {

        public static DateTime? WDGetCreationDate(string fileName)
        {
            DateTime? creationDate = null;

            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(fileName, false))
            {
                PackageProperties properties = wordprocessingDocument.PackageProperties;
                creationDate = properties.Created;
            }
            return creationDate;
        }

        public static void WDAddTable(string fileName, string[,] data)
        {
            //Take the data from a 2-dimensional array and build a table at the end of the supplied document.
            using (var document = WordprocessingDocument.Open(fileName, true))
            {
                var docPart = document.MainDocumentPart;
                var doc = docPart.Document;

                var table = new Table();

                //Can create object explicitly, or as shown later, implicitly.
                var tb = new TopBorder();
                tb.Val = BorderValues.DashDotStroked;
                tb.Size = 12;

                var borders = new TableBorders();
                borders.TopBorder = tb;
                borders.LeftBorder = new LeftBorder() { Val = BorderValues.Single, Size = 12 };
                borders.BottomBorder = new BottomBorder() { Val = BorderValues.Single, Size = 12 };
                borders.RightBorder = new RightBorder() { Val = BorderValues.Single, Size = 12 };
                borders.InsideHorizontalBorder = new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 8 };
                borders.InsideVerticalBorder = new InsideVerticalBorder() { Val = BorderValues.Single, Size = 8 };

                var props = new TableProperties();
                props.Append(borders);

                table.Append(props);

                for (var i = 0; i <= data.GetUpperBound(0); i++)
                {
                    var tr = new TableRow();

                    for (var j = 0; j <= data.GetUpperBound(1); j++)
                    {
                        var tc = new TableCell();

                        var runProp = new RunProperties();
                        runProp.Append(new Bold());
                        runProp.Append(new Color() { Val = "FF0000" });

                        var run = new Run();
                        run.Append(runProp);

                        var t = new Text(data[i, j]);
                        run.Append(t);

                        var justification = new Justification();
                        justification.Val = JustificationValues.Center;
                        var paraProps = new ParagraphProperties(justification);

                        var p = new Paragraph();
                        p.Append(paraProps);
                        p.Append(run);
                        tc.Append(p);

                        //If you want the column to be the right width for your data, you
                        //examine all the rows in this column, calculate the width of the
                        //for each, find the maximum, and convert to 1/20 points. This
                        //example assumes you want columns that are 2000/20 points wide.

                        var tcp = new TableCellProperties();
                        var tcw = new TableCellWidth();
                        tcw.Type = TableWidthUnitValues.Dxa;
                        tcw.Width = "2000";
                        tcp.Append(tcw);
                        tc.Append(tcp);
                        tr.Append(tc);

                    }
                    table.Append(tr);
                }

                doc.Body.Append(table);
                doc.Save();
            }
        }

        public static string WDRetrieveText(string fileName)
        {
            string results = null;

            using (var document = WordprocessingDocument.Open(fileName, false))
            {
                //Retrieve the document part:
                var docPart = document.MainDocumentPart;
                if (docPart != null && docPart.Document != null)
                {
                    results = GetPlainText(docPart.Document.Body);
                }
            }
            return results;
        }
        public static string GetPlainText(OpenXmlElement element)
        {
            StringBuilder sb = new StringBuilder();

            foreach (OpenXmlElement section in element.Elements())
            {
                switch (section.LocalName)
                {
                    case "cr":              //Carriage return
                    case "br":              //Page break
                        sb.Append(Environment.NewLine);
                        break;
                    //Tab
                    case "tab":
                        sb.Append("\t");
                        break;
                    //Text
                    case "t":
                        sb.Append(section.InnerText);
                        break;
                    //Paragraph
                    case "p":
                        sb.Append(GetPlainText(section));
                        sb.AppendLine(Environment.NewLine);
                        break;

                    default:
                        sb.Append(GetPlainText(section));
                        break;
                }
            }
            return sb.ToString();
        }
    }
}
