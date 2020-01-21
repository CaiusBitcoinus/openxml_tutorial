using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using OXML.DL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace OXML.Library
{
    public class Helper_Excel
    {
        #region Some Excel functions
        public static string XLCreateEmptyFile(string fileName, string Location)
        {
            string result = null;
            string fullPath = Location + fileName;

            using (var document = SpreadsheetDocument.Create(fullPath, SpreadsheetDocumentType.Workbook))
            {
                // Creates 4 things: WorkBookPart, WorkSheetPart, WorkSheet, SheetData
                document.AddWorkbookPart().AddNewPart<WorksheetPart>().Worksheet = new Worksheet(new SheetData());

                document.WorkbookPart.Workbook =
                    new Workbook(
                        new Sheets(
                            new Sheet
                            {
                                // Id is used to create Sheet to WorksheetPart relationship
                                Id = document.WorkbookPart.GetIdOfPart(document.WorkbookPart.WorksheetParts.First()),
                                // SheetId and Name are both required
                                SheetId = 1,
                                Name = "Sheet1"
                            }));
            }
            return result;
        }
        public static string XLCreateEmptyFile_version2(string fileName, string Location)
        {
            string result = null;
            string fullPath = Location + fileName;

            using (var document = SpreadsheetDocument.Create(fullPath, SpreadsheetDocumentType.Workbook))
            {
                // Creates 4 things: WorkBookPart, WorkSheetPart, WorkSheet, SheetData
                document.AddWorkbookPart().AddNewPart<WorksheetPart>().Worksheet = new Worksheet(new SheetData());

                var workbook = new Workbook();
                var sheets = new Sheets();
                var sheet = new Sheet();
                // Id is used to create Sheet to WorksheetPart relationship
                sheet.Id = document.WorkbookPart.GetIdOfPart(document.WorkbookPart.WorksheetParts.First());
                // SheetId and Name are both required
                sheet.SheetId = 1;
                sheet.Name = "Empty Sheet";

                sheets.Append(sheet);
                workbook.Append(sheets);
                document.WorkbookPart.Workbook = workbook;
            }
            return result;
        }

        public static string GetAllExcelSheetsName(string fileName)
        {
            string result = null;

            Sheets sheets = XLGetAllSheets(fileName);
            foreach (Sheet sheet in sheets)
            {
                result += sheet.Name + Environment.NewLine;
            }

            return result;
        }
        private static Sheets XLGetAllSheets(string fileName)
        {
            Sheets theSheets = null;

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart wbPart = spreadsheetDocument.WorkbookPart;
                theSheets = wbPart.Workbook.Sheets;
            }

            return theSheets;
        }

        public static string XLGetCellValue(string fileName, string sheetName, string addressName)
        {
            string value = null;
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheet theSheet = workbookPart.Workbook.Descendants<Sheet>().
                    FirstOrDefault(s => s.Name == sheetName);

                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName Not Found!");
                }
                WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(theSheet.Id));
                Cell theCell = worksheetPart.Worksheet.Descendants<Cell>().
                    FirstOrDefault(c => c.CellReference == addressName);

                if (theCell != null)
                {
                    value = theCell.CellValue.InnerText;
                }
                if (theCell.DataType != null)
                {
                    if (theCell.DataType.Value == CellValues.SharedString)
                    {
                        var stringTable = workbookPart.SharedStringTablePart;
                        if (stringTable != null)
                        {
                            var textItem = stringTable.SharedStringTable.
                                ElementAtOrDefault(int.Parse(value));
                            if (textItem != null)
                            {
                                value = textItem.InnerText;
                            }
                        }
                    }
                    else if (theCell.DataType.Value == CellValues.Number)
                    {
                        value = theCell.CellValue.InnerText;
                    }
                    else if (theCell.DataType.Value == CellValues.Boolean)
                    {
                        switch (value)
                        {
                            case "0":
                                value = "FALSE";
                                break;
                            case "1":
                                value = "TRUE";
                                break;
                        }
                    }
                }
            }
            return value;
        }

        private static Cell RetrieveCellReference(Worksheet worksheet, string addressName)
        {
            //Use regular expressions to get the row number and column name.
            //If the paramenter wasn't well formed this code will fail
            Regex regex = new Regex("^(?<col>\\D+)(?<row>\\d+)");
            Match match = regex.Match(addressName);
            uint rowNumber = uint.Parse(match.Result("${row}"));
            string colName = match.Result("${col");

            //Retrieve reference to the sheet's data.
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            //If the worksheet does not contain a row
            //with the specified row index, insert one.
            var rows = sheetData.Elements<Row>();
            var theRow = rows.FirstOrDefault(r => r.RowIndex.Value == rowNumber);
            if (theRow == null)
            {
                //Rows dosen't exist, so we create a new one.
                theRow = new Row();
                theRow.RowIndex = rowNumber;
                //Must insert row in the appropiate location.
                Row refRow = null;
                foreach (Row row in rows)
                {
                    if (row.RowIndex > rowNumber)
                    {
                        refRow = row;
                        break;
                    }
                }
                //If refRow is null, InsertBefore appends.
                sheetData.InsertBefore(theRow, refRow);
            }
            //At this point, theRow refers to the row to contain the cell value.
            //The cell may or may not exist.

            //If the cell you need already exists, return it.
            //If there is not a cell with the specified address name, insert one.
            var cells = theRow.Elements<Cell>();
            Cell theCell = cells.FirstOrDefault(c => c.CellReference.Value == addressName);
            if (theCell == null)
            {
                //Cell dosen't exist, so create one.
                theCell = new Cell();
                theCell.CellReference = addressName;
                //Must insert cell in the appropiate location.
                Cell refCell = null;
                foreach (Cell cell in cells)
                {
                    if (string.Compare(cell.CellReference.Value, addressName, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }
                //If refCell is null, InsertBefore appends.
                theRow.InsertBefore(theCell, refCell);
            }
            return theCell;
        }
        public static bool XLInsertNumberIntoCell(string fileName, string sheetName, string addressName, int value)
        {
            //Given a file, a sheet, and a cell, insert a specified value.
            //For exemple: InsertNumberIntoCell("E:\003_P\openxml_tutorial\Test_Excel.xlsx", "Sheet1", "C3", 14)

            //Assume failure.
            bool returnValue = false;

            //Open the document for editing.
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                Sheet theSheet = workbookPart.Workbook.Descendants<Sheet>().
                    FirstOrDefault(s => s.Name == sheetName);
                if (theSheet != null)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(theSheet.Id));
                    Worksheet worksheet = worksheetPart.Worksheet;
                    Cell theCell = RetrieveCellReference(worksheet, addressName);

                    theCell.CellValue = new CellValue(value.ToString());
                    theCell.DataType = CellValues.Number;

                    //Save the worksheet
                    worksheet.Save();
                    returnValue = true;
                }

            }
            return returnValue;
        }
        public static bool XLInsertTextIntoCell(string fileName, string sheetName, string addressName, string value)
        {
            //Given a file, a sheet, and a cell, insert a specified value.
            //For exemple: InsertNumberIntoCell("E:\003_P\openxml_tutorial\Test_Excel.xlsx", "Sheet1", "C3", 14)

            //Assume failure.
            bool returnValue = false;

            //Open the document for editing.
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                Sheet theSheet = workbookPart.Workbook.Descendants<Sheet>().
                    FirstOrDefault(s => s.Name == sheetName);
                if (theSheet != null)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(theSheet.Id));
                    Worksheet worksheet = worksheetPart.Worksheet;
                    Cell theCell = RetrieveCellReference(worksheet, addressName);

                    theCell.CellValue = new CellValue(value.ToString());
                    theCell.DataType = CellValues.String;

                    //Save the worksheet
                    worksheet.Save();
                    returnValue = true;
                }

            }
            return returnValue;
        }

        public static void InsertWorksheet(string fullPath)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(fullPath, true))
            {
                //Add a blank WorksheetPart.
                WorksheetPart newWorksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                string relationshipId = spreadsheet.WorkbookPart.GetIdOfPart(newWorksheetPart);

                //Get a unique ID for the new worksheet
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                //Give the new worksheet a name.
                string sheetName = "Sheet" + sheetId;

                //Append the new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = relationshipId,
                    SheetId = sheetId,
                    Name = sheetName
                };
                sheets.Append(sheet);
            }
        }
        #endregion

        /*
         * Here you find a complete, yet more complex solution for:
         *         1_ generating an Excel File
         *         2_ generating a worksheet for him
         *         3_ inserting data
         *         4_ generating a table
         *         __to be continued
         */
        #region I have grouped this functions togheter because they are all interconnected
        /*
         * OpenXml packages used here:
            using DocumentFormat.OpenXml;
            using DocumentFormat.OpenXml.Packaging;
            using DocumentFormat.OpenXml.Spreadsheet;
            using X14 = DocumentFormat.OpenXml.Office2010.Excel;
            using X15 = DocumentFormat.OpenXml.Office2013.Excel;
         */
        /*
         * Code for creating an Excel file into given path.
         */
        public void CreateExcelFile(List<Employee> employees_List, string OutPutFileDirectory)
        {
            var datetime = DateTime.Now.ToString().Replace("/", "_").Replace(":", "_");
            string fileFullname = Path.Combine(OutPutFileDirectory, "Output.xlsx");

            if (File.Exists(fileFullname))
            {
                fileFullname = Path.Combine(OutPutFileDirectory, "Output_" + datetime + ".xlsx");
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Create(fileFullname, SpreadsheetDocumentType.Workbook))
            {
                CreatePartsForExcel(package, employees_List);
            }
        }
        /*
         * Write functions for creating workbook and worksheet into Excel.
         */
        private void CreatePartsForExcel(SpreadsheetDocument document, List<Employee> employees_List)
        {
            SheetData partSheetData = GenerateSheetdataForDetails(employees_List);
            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPartContent(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPartContent(workbookStylesPart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPartContent(worksheetPart1, partSheetData);
        }
        /*
         * Write functions for creating workbook and work sheet content in Excel
         */
        private void GenerateWorkbookPartContent(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook();
            Sheets sheets1 = new Sheets();

            Sheet sheet1 = new Sheet()
            {
                Name = "Sheet1",
                SheetId = (UInt32Value)1U,
                Id = "rId1"
            };

            sheets1.Append(sheet1);
            workbook1.Append(sheets1);
            workbookPart1.Workbook = workbook1;
        }
        private void GenerateWorksheetPartContent(WorksheetPart worksheetPart1, SheetData sheetData1)
        {
            Worksheet worksheet1 = new Worksheet()
            {
                MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" }
            };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            SheetDimension sheetDimension1 = new SheetDimension()
            {
                Reference = "A1"
            };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView()
            {
                TabSelected = true,
                WorkbookViewId = (UInt32Value)0U
            };

            Selection selection1 = new Selection()
            {
                ActiveCell = "A1",
                SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" }
            };

            sheetView1.Append(selection1);
            sheetViews1.Append(sheetView1);

            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties()
            {
                DefaultRowHeight = 15D,
                DyDescent = 0.25D
            };

            PageMargins pageMargins1 = new PageMargins()
            {
                Left = 0.7D,
                Right = 0.7D,
                Top = 0.75D,
                Bottom = 0.75D,
                Header = 0.3D,
                Footer = 0.3D
            };

            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(pageMargins1);

            worksheetPart1.Worksheet = worksheet1;
        }
        /*
         * Write code for workbook styles by giving your own
         * font size, color, font name, border properties, cell style formats etc.
         */
        private void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet()
            {
                MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" }
            };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new Fonts()
            {
                Count = (UInt32Value)2U,
                KnownFonts = true
            };

            #region Fonts
            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);

            fonts1.Append(font1);
            fonts1.Append(font2);
            #endregion
            #region Fills
            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };
            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };
            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);
            #endregion
            #region Borders
            Borders borders1 = new Borders() { Count = (UInt32Value)2U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };
            leftBorder2.Append(color3);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };
            rightBorder2.Append(color4);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color5 = new Color() { Indexed = (UInt32Value)64U };
            topBorder2.Append(color5);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color6 = new Color() { Indexed = (UInt32Value)64U };
            bottomBorder2.Append(color6);

            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);
            #endregion
            #region CellStyleFormats
            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat()
            {
                NumberFormatId = (UInt32Value)0U,
                FontId = (UInt32Value)0U,
                FillId = (UInt32Value)0U,
                BorderId = (UInt32Value)0U
            };
            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32)3U };

            CellFormat cellFormat2 = new CellFormat()
            {
                NumberFormatId = (UInt32Value)0U,
                FontId = (UInt32Value)0U,
                FillId = (UInt32Value)0U,
                BorderId = (UInt32Value)0U,
                FormatId = (UInt32Value)0U
            };
            CellFormat cellFormat3 = new CellFormat()
            {
                NumberFormatId = (UInt32Value)0U,
                FontId = (UInt32Value)0U,
                FillId = (UInt32Value)0U,
                BorderId = (UInt32Value)0U,
                FormatId = (UInt32Value)0U,
                ApplyBorder = true
            };
            CellFormat cellFormat4 = new CellFormat()
            {
                NumberFormatId = (UInt32Value)0U,
                FontId = (UInt32Value)0U,
                FillId = (UInt32Value)0U,
                BorderId = (UInt32Value)0U,
                FormatId = (UInt32Value)0U,
                ApplyFont = true,
                ApplyBorder = true
            };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle()
            {
                Name = "Normal",
                FormatId = (UInt32Value)0U,
                BuiltinId = (UInt32Value)0U
            };

            cellStyle1.Append(cellStyle1);

            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles()
            {
                Count = (UInt32Value)0U,
                DefaultTableStyle = "TableStyleMedium2",
                DefaultPivotStyle = "PivotStyleLight16"
            };
            #endregion
            #region StylesheetExtensionLists
            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles1 = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);
            #endregion

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }
        /*
         * Write the below functions to add data into Excel
         */
        private SheetData GenerateSheetdataForDetails(List<Employee> employees_List)
        {
            SheetData sheetData1 = new SheetData();
            sheetData1.Append(CreateHeaderRowForExcel());

            foreach (Employee employee in employees_List)
            {
                Row partsRows = GenerateRowForChildPartDetail(employee);
                sheetData1.Append(partsRows);
            }
            return sheetData1;
        }
        /*
         * Below function is created for creating Header rows in Excel.
         */
        private Row CreateHeaderRowForExcel()
        {
            Row workRow = new Row();
            workRow.Append(CreateCell("Test Id", 2U));
            workRow.Append(CreateCell("Test Name", 2U));
            workRow.Append(CreateCell("Test Description", 2U));
            workRow.Append(CreateCell("Test Date", 2U));
            return workRow;
        }
        /*
         * Below function is used for generating child rows.
         */
        private Row GenerateRowForChildPartDetail(Employee employee)
        {
            Row tableRow = new Row();
            tableRow.Append(CreateCell(employee.FirstName));
            tableRow.Append(CreateCell(employee.LastName));
            tableRow.Append(CreateCell(employee.Age.ToString()));
            tableRow.Append(CreateCell(employee.HireDate.ToString()));

            return tableRow;
        }
        /*
         * Below function is used for creating cell by passing only cell data and it adds default style.
         */
        private Cell CreateCell(string text)
        {
            Cell cell = new Cell();
            cell.StyleIndex = 1U;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);

            return cell;
        }
        /*
         * Below function is used for creating a cell by passing cell data and cell style.
         */
        private Cell CreateCell(string text, uint styleIndex)
        {
            Cell cell = new Cell();
            cell.StyleIndex = styleIndex;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);

            return cell;
        }
        /*
         * Below function is created for resolving the data type of numeric value in a cell.
         */
        private EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            int intVal;
            double doubleVal;
            if (int.TryParse(text, out intVal) || double.TryParse(text, out doubleVal))
            {
                return CellValues.Number;
            }
            else
            {
                return CellValues.String;
            }
        }
        /*
         * Below function is created for resolving the data type of numeric value in a cell.
         */






        #endregion

    }
}
