using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OXML.DL;
using OXML.Library;

namespace OXML.UI
{
    public partial class MainForm : Form
    {
        private readonly string PPT_Location = @"E:\003_P\openxml_tutorial\Test_PowerPoinT.pptx";
        //private readonly string Excel_Location = @"E:\003_P\openxml_tutorial\Test_Excel.xlsx";
        private readonly string Word_Location = @"E:\003_P\openxml_tutorial\Test_Word.docx";

        public MainForm()
        {
            InitializeComponent();
        }

        private void Btn_GetSlidesNumber_Click(object sender, EventArgs e)
        {
            Helper_PPT helper_PPT = new Helper_PPT();
            AppendTo_ExcelLog(helper_PPT.PPTGetSlideCount(PPT_Location).ToString());
            AppendTo_ExcelLog(helper_PPT.PPTGetSlideCount(PPT_Location, false).ToString());
        }

        private void Btn_GetExcelSheetsNames_Click(object sender, EventArgs e)
        {
            string excelLocation = tbo_CurrentExcelFile.Text;
            AppendTo_ExcelLog(Helper_Excel.GetAllExcelSheetsName(excelLocation));
        }

        private void Btn_GetWordCreationDateTime_Click(object sender, EventArgs e)
        {
            DateTime? created = Helper_Word.WDGetCreationDate(Word_Location);
            AppendTo_ExcelLog(created.ToString());
        }

        private void Btn_GetExcelCellValue_Click(object sender, EventArgs e)
        {
            string excelLocation = tbo_CurrentExcelFile.Text;
            AppendTo_ExcelLog(Helper_Excel.XLGetCellValue(excelLocation, "Sheet1", "A1"));
            AppendTo_ExcelLog(Helper_Excel.XLGetCellValue(excelLocation, "Sheet1", "C4"));
        }

        private void Btn_WriteInExcelCell_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrWhiteSpace(tbo_CurrentExcelFile.Text) == false)
            {
                string excelLocation = tbo_CurrentExcelFile.Text;
                AppendTo_ExcelLog(Helper_Excel.XLInsertTextIntoCell(excelLocation, "Sheet1", "A1", "Hello There"));
                AppendTo_ExcelLog(Helper_Excel.XLInsertNumberIntoCell(excelLocation, "Sheet1", "C4", 420));
            }
            else
            {
                Tbo_ExcelLog.Text += "Pick an Excel File!";
            }
        }

        private void Btn_InsertTableInWorldDocument_Click(object sender, EventArgs e)
        {
            string[,] data =
            {
                { "Row 1, Col 1", "Row 1, Col 2" },
                { "Row 2, Col 1", "Row 2, Col 2" },
                { "Row 3, Col 1", "Row 3, Col 2" },
            };

            Helper_Word.WDAddTable(Word_Location, data);
        }

        private void Btn_RetrievePlainText_Click(object sender, EventArgs e)
        {
            Tbo_ExcelLog.Text += Helper_Word.WDRetrieveText(Word_Location) + Environment.NewLine;
        }

        private void Btn_CreateEmptyExcelFile_Click(object sender, EventArgs e)
        {
            using(FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                //navigate to a folder. If the user input is "OK" ..Continue operation
                if(folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    string ExcelName_01 = "ExcelExemple_01.xlsx";
                    string ExcelName_02 = "ExcelExemple_02.xlsx";
                    string folderPath = folderBrowserDialog.SelectedPath;
                    Helper_Excel.XLCreateEmptyFile(ExcelName_01, folderPath);
                    Helper_Excel.XLCreateEmptyFile_version2(ExcelName_02, folderPath);

                    //keep the location of one of the Excel files for future testing
                    tbo_CurrentExcelFile.Text = folderPath + ExcelName_01;
                }
            }
        }

        private void Btn_AddExcelSheets_Click(object sender, EventArgs e)
        {
            string excelLocation = tbo_CurrentExcelFile.Text;
            for (int i = 0; i < 3; i++)
            {
                Helper_Excel.InsertWorksheet(excelLocation);
            }
        }

        private void AppendTo_ExcelLog(object message)
        {
            string final = $"{Environment.NewLine} {message}";

            Tbo_ExcelLog.AppendText(final);
        }

        private void Btn_CreateExcelFileTableFormattedCells_Click(object sender, EventArgs e)
        {
            #region Creating some objects to write in Excel

            List<Employee> employees_List = new List<Employee>();
            Employee employee1 = new Employee()
            {
                FirstName = "Arakida",
                LastName = "Moritake",
                Age = 30,
                HireDate = DateTime.Now
            };
            Employee employee2 = new Employee()
            {
                FirstName = "Petrescu",
                LastName = "Camil",
                Age = 30,
                HireDate = DateTime.Now.AddDays(-100D)
            };

            employees_List.Add(employee1);
            employees_List.Add(employee2);

            #endregion
            #region Create path of in User\MyDocuments\OXMLTesting

            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string foldername = "OXMLTesting";
            string fullPath = System.IO.Path.Combine(path, foldername);

            #endregion
            #region Create the Excel File with the objects

            Helper_Excel helper_Excel = new Helper_Excel();
            helper_Excel.CreateExcelFile(employees_List, fullPath);
            AppendTo_ExcelLog($"Excel Created in {fullPath}");

            #endregion



        }
    }
}
/*Testing Repository Namespace transfer. --==Ignore this==--

speaking of migration..

    here is a penguin family

          (o_
(o_  (o_  //\
(/)_ (/)_ V_/_

 */

