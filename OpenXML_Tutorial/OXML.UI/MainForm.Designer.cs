namespace OXML.UI
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Btn_GetSlidesNumber = new System.Windows.Forms.Button();
            this.Tbo_ExcelLog = new System.Windows.Forms.TextBox();
            this.Btn_GetExcelSheetsNames = new System.Windows.Forms.Button();
            this.Btn_GetWordCreationDateTime = new System.Windows.Forms.Button();
            this.Btn_GetExcelCellValue = new System.Windows.Forms.Button();
            this.Btn_WriteInExcelCell = new System.Windows.Forms.Button();
            this.Btn_InsertTableInWorldDocument = new System.Windows.Forms.Button();
            this.Btn_RetrievePlainText = new System.Windows.Forms.Button();
            this.tabControl_mainTabControl = new System.Windows.Forms.TabControl();
            this.tPage_Exel = new System.Windows.Forms.TabPage();
            this.Btn_CreateExcelFileTableFormattedCells = new System.Windows.Forms.Button();
            this.btn_AddExcelSheets = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.tbo_CurrentExcelFile = new System.Windows.Forms.TextBox();
            this.btn_CreateEmptyExcelFile = new System.Windows.Forms.Button();
            this.tPage_Word = new System.Windows.Forms.TabPage();
            this.tPage_PPT = new System.Windows.Forms.TabPage();
            this.tPage_disclaimer = new System.Windows.Forms.TabPage();
            this.button1 = new System.Windows.Forms.Button();
            this.tabControl_mainTabControl.SuspendLayout();
            this.tPage_Exel.SuspendLayout();
            this.tPage_Word.SuspendLayout();
            this.tPage_PPT.SuspendLayout();
            this.SuspendLayout();
            // 
            // Btn_GetSlidesNumber
            // 
            this.Btn_GetSlidesNumber.Location = new System.Drawing.Point(823, 26);
            this.Btn_GetSlidesNumber.Margin = new System.Windows.Forms.Padding(4);
            this.Btn_GetSlidesNumber.Name = "Btn_GetSlidesNumber";
            this.Btn_GetSlidesNumber.Size = new System.Drawing.Size(451, 65);
            this.Btn_GetSlidesNumber.TabIndex = 0;
            this.Btn_GetSlidesNumber.Text = "Get number of PPT Slides";
            this.Btn_GetSlidesNumber.UseVisualStyleBackColor = true;
            this.Btn_GetSlidesNumber.Click += new System.EventHandler(this.Btn_GetSlidesNumber_Click);
            // 
            // Tbo_ExcelLog
            // 
            this.Tbo_ExcelLog.Location = new System.Drawing.Point(836, 42);
            this.Tbo_ExcelLog.Multiline = true;
            this.Tbo_ExcelLog.Name = "Tbo_ExcelLog";
            this.Tbo_ExcelLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.Tbo_ExcelLog.Size = new System.Drawing.Size(590, 602);
            this.Tbo_ExcelLog.TabIndex = 1;
            // 
            // Btn_GetExcelSheetsNames
            // 
            this.Btn_GetExcelSheetsNames.Location = new System.Drawing.Point(6, 134);
            this.Btn_GetExcelSheetsNames.Margin = new System.Windows.Forms.Padding(4);
            this.Btn_GetExcelSheetsNames.Name = "Btn_GetExcelSheetsNames";
            this.Btn_GetExcelSheetsNames.Size = new System.Drawing.Size(450, 35);
            this.Btn_GetExcelSheetsNames.TabIndex = 2;
            this.Btn_GetExcelSheetsNames.Text = "4) Get names of Excel Sheets";
            this.Btn_GetExcelSheetsNames.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Btn_GetExcelSheetsNames.UseVisualStyleBackColor = true;
            this.Btn_GetExcelSheetsNames.Click += new System.EventHandler(this.Btn_GetExcelSheetsNames_Click);
            // 
            // Btn_GetWordCreationDateTime
            // 
            this.Btn_GetWordCreationDateTime.Location = new System.Drawing.Point(23, 47);
            this.Btn_GetWordCreationDateTime.Margin = new System.Windows.Forms.Padding(4);
            this.Btn_GetWordCreationDateTime.Name = "Btn_GetWordCreationDateTime";
            this.Btn_GetWordCreationDateTime.Size = new System.Drawing.Size(451, 65);
            this.Btn_GetWordCreationDateTime.TabIndex = 3;
            this.Btn_GetWordCreationDateTime.Text = "Get Creation DateTime of Word";
            this.Btn_GetWordCreationDateTime.UseVisualStyleBackColor = true;
            this.Btn_GetWordCreationDateTime.Click += new System.EventHandler(this.Btn_GetWordCreationDateTime_Click);
            // 
            // Btn_GetExcelCellValue
            // 
            this.Btn_GetExcelCellValue.Location = new System.Drawing.Point(6, 91);
            this.Btn_GetExcelCellValue.Margin = new System.Windows.Forms.Padding(4);
            this.Btn_GetExcelCellValue.Name = "Btn_GetExcelCellValue";
            this.Btn_GetExcelCellValue.Size = new System.Drawing.Size(450, 35);
            this.Btn_GetExcelCellValue.TabIndex = 4;
            this.Btn_GetExcelCellValue.Text = "3) Get Excel Cell Values from A1 and C4";
            this.Btn_GetExcelCellValue.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Btn_GetExcelCellValue.UseVisualStyleBackColor = true;
            this.Btn_GetExcelCellValue.Click += new System.EventHandler(this.Btn_GetExcelCellValue_Click);
            // 
            // Btn_WriteInExcelCell
            // 
            this.Btn_WriteInExcelCell.Location = new System.Drawing.Point(6, 48);
            this.Btn_WriteInExcelCell.Margin = new System.Windows.Forms.Padding(4);
            this.Btn_WriteInExcelCell.Name = "Btn_WriteInExcelCell";
            this.Btn_WriteInExcelCell.Size = new System.Drawing.Size(450, 35);
            this.Btn_WriteInExcelCell.TabIndex = 5;
            this.Btn_WriteInExcelCell.Text = "2) Write \"Hello There\" in A1 and 420 in C4 Excel Cells";
            this.Btn_WriteInExcelCell.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Btn_WriteInExcelCell.UseVisualStyleBackColor = true;
            this.Btn_WriteInExcelCell.Click += new System.EventHandler(this.Btn_WriteInExcelCell_Click);
            // 
            // Btn_InsertTableInWorldDocument
            // 
            this.Btn_InsertTableInWorldDocument.Location = new System.Drawing.Point(23, 120);
            this.Btn_InsertTableInWorldDocument.Margin = new System.Windows.Forms.Padding(4);
            this.Btn_InsertTableInWorldDocument.Name = "Btn_InsertTableInWorldDocument";
            this.Btn_InsertTableInWorldDocument.Size = new System.Drawing.Size(451, 65);
            this.Btn_InsertTableInWorldDocument.TabIndex = 7;
            this.Btn_InsertTableInWorldDocument.Text = "Insert Table In Word Document";
            this.Btn_InsertTableInWorldDocument.UseVisualStyleBackColor = true;
            this.Btn_InsertTableInWorldDocument.Click += new System.EventHandler(this.Btn_InsertTableInWorldDocument_Click);
            // 
            // Btn_RetrievePlainText
            // 
            this.Btn_RetrievePlainText.Location = new System.Drawing.Point(23, 193);
            this.Btn_RetrievePlainText.Margin = new System.Windows.Forms.Padding(4);
            this.Btn_RetrievePlainText.Name = "Btn_RetrievePlainText";
            this.Btn_RetrievePlainText.Size = new System.Drawing.Size(451, 65);
            this.Btn_RetrievePlainText.TabIndex = 8;
            this.Btn_RetrievePlainText.Text = "Retrieve Plain Text from Word";
            this.Btn_RetrievePlainText.UseVisualStyleBackColor = true;
            this.Btn_RetrievePlainText.Click += new System.EventHandler(this.Btn_RetrievePlainText_Click);
            // 
            // tabControl_mainTabControl
            // 
            this.tabControl_mainTabControl.Controls.Add(this.tPage_Exel);
            this.tabControl_mainTabControl.Controls.Add(this.tPage_Word);
            this.tabControl_mainTabControl.Controls.Add(this.tPage_PPT);
            this.tabControl_mainTabControl.Controls.Add(this.tPage_disclaimer);
            this.tabControl_mainTabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl_mainTabControl.Location = new System.Drawing.Point(0, 0);
            this.tabControl_mainTabControl.Name = "tabControl_mainTabControl";
            this.tabControl_mainTabControl.Padding = new System.Drawing.Point(3, 3);
            this.tabControl_mainTabControl.SelectedIndex = 0;
            this.tabControl_mainTabControl.Size = new System.Drawing.Size(1442, 688);
            this.tabControl_mainTabControl.TabIndex = 9;
            // 
            // tPage_Exel
            // 
            this.tPage_Exel.Controls.Add(this.button1);
            this.tPage_Exel.Controls.Add(this.Btn_CreateExcelFileTableFormattedCells);
            this.tPage_Exel.Controls.Add(this.btn_AddExcelSheets);
            this.tPage_Exel.Controls.Add(this.label1);
            this.tPage_Exel.Controls.Add(this.tbo_CurrentExcelFile);
            this.tPage_Exel.Controls.Add(this.btn_CreateEmptyExcelFile);
            this.tPage_Exel.Controls.Add(this.Tbo_ExcelLog);
            this.tPage_Exel.Controls.Add(this.Btn_GetExcelSheetsNames);
            this.tPage_Exel.Controls.Add(this.Btn_WriteInExcelCell);
            this.tPage_Exel.Controls.Add(this.Btn_GetExcelCellValue);
            this.tPage_Exel.Location = new System.Drawing.Point(4, 28);
            this.tPage_Exel.Name = "tPage_Exel";
            this.tPage_Exel.Padding = new System.Windows.Forms.Padding(3);
            this.tPage_Exel.Size = new System.Drawing.Size(1434, 656);
            this.tPage_Exel.TabIndex = 0;
            this.tPage_Exel.Text = "Excel Tutorials";
            this.tPage_Exel.UseVisualStyleBackColor = true;
            // 
            // Btn_CreateExcelFileTableFormattedCells
            // 
            this.Btn_CreateExcelFileTableFormattedCells.Enabled = false;
            this.Btn_CreateExcelFileTableFormattedCells.Location = new System.Drawing.Point(6, 220);
            this.Btn_CreateExcelFileTableFormattedCells.Margin = new System.Windows.Forms.Padding(4);
            this.Btn_CreateExcelFileTableFormattedCells.Name = "Btn_CreateExcelFileTableFormattedCells";
            this.Btn_CreateExcelFileTableFormattedCells.Size = new System.Drawing.Size(450, 35);
            this.Btn_CreateExcelFileTableFormattedCells.TabIndex = 11;
            this.Btn_CreateExcelFileTableFormattedCells.Text = "6) Create Excel File with table and formatted cells";
            this.Btn_CreateExcelFileTableFormattedCells.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Btn_CreateExcelFileTableFormattedCells.UseVisualStyleBackColor = true;
            this.Btn_CreateExcelFileTableFormattedCells.Click += new System.EventHandler(this.Btn_CreateExcelFileTableFormattedCells_Click);
            // 
            // btn_AddExcelSheets
            // 
            this.btn_AddExcelSheets.Location = new System.Drawing.Point(6, 177);
            this.btn_AddExcelSheets.Margin = new System.Windows.Forms.Padding(4);
            this.btn_AddExcelSheets.Name = "btn_AddExcelSheets";
            this.btn_AddExcelSheets.Size = new System.Drawing.Size(450, 35);
            this.btn_AddExcelSheets.TabIndex = 10;
            this.btn_AddExcelSheets.Text = "5) Add 3 more Sheets to Excel";
            this.btn_AddExcelSheets.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btn_AddExcelSheets.UseVisualStyleBackColor = true;
            this.btn_AddExcelSheets.Click += new System.EventHandler(this.Btn_AddExcelSheets_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(832, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(119, 19);
            this.label1.TabIndex = 9;
            this.label1.Text = "Current Excel file:";
            // 
            // tbo_CurrentExcelFile
            // 
            this.tbo_CurrentExcelFile.Location = new System.Drawing.Point(982, 6);
            this.tbo_CurrentExcelFile.Name = "tbo_CurrentExcelFile";
            this.tbo_CurrentExcelFile.Size = new System.Drawing.Size(449, 26);
            this.tbo_CurrentExcelFile.TabIndex = 8;
            this.tbo_CurrentExcelFile.Text = "E:\\ExcelExemple_01.xlsx";
            // 
            // btn_CreateEmptyExcelFile
            // 
            this.btn_CreateEmptyExcelFile.Location = new System.Drawing.Point(6, 6);
            this.btn_CreateEmptyExcelFile.Name = "btn_CreateEmptyExcelFile";
            this.btn_CreateEmptyExcelFile.Size = new System.Drawing.Size(450, 35);
            this.btn_CreateEmptyExcelFile.TabIndex = 7;
            this.btn_CreateEmptyExcelFile.Text = "1) Create empty Excel file in location..";
            this.btn_CreateEmptyExcelFile.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btn_CreateEmptyExcelFile.UseVisualStyleBackColor = true;
            this.btn_CreateEmptyExcelFile.Click += new System.EventHandler(this.Btn_CreateEmptyExcelFile_Click);
            // 
            // tPage_Word
            // 
            this.tPage_Word.Controls.Add(this.Btn_GetWordCreationDateTime);
            this.tPage_Word.Controls.Add(this.Btn_RetrievePlainText);
            this.tPage_Word.Controls.Add(this.Btn_InsertTableInWorldDocument);
            this.tPage_Word.Location = new System.Drawing.Point(4, 28);
            this.tPage_Word.Name = "tPage_Word";
            this.tPage_Word.Padding = new System.Windows.Forms.Padding(3);
            this.tPage_Word.Size = new System.Drawing.Size(1434, 656);
            this.tPage_Word.TabIndex = 1;
            this.tPage_Word.Text = "Word Tutorials";
            this.tPage_Word.UseVisualStyleBackColor = true;
            // 
            // tPage_PPT
            // 
            this.tPage_PPT.Controls.Add(this.Btn_GetSlidesNumber);
            this.tPage_PPT.Location = new System.Drawing.Point(4, 28);
            this.tPage_PPT.Name = "tPage_PPT";
            this.tPage_PPT.Padding = new System.Windows.Forms.Padding(3);
            this.tPage_PPT.Size = new System.Drawing.Size(1434, 656);
            this.tPage_PPT.TabIndex = 2;
            this.tPage_PPT.Text = "PPT Tutorials";
            this.tPage_PPT.UseVisualStyleBackColor = true;
            // 
            // tPage_disclaimer
            // 
            this.tPage_disclaimer.Location = new System.Drawing.Point(4, 28);
            this.tPage_disclaimer.Name = "tPage_disclaimer";
            this.tPage_disclaimer.Padding = new System.Windows.Forms.Padding(3);
            this.tPage_disclaimer.Size = new System.Drawing.Size(1434, 656);
            this.tPage_disclaimer.TabIndex = 3;
            this.tPage_disclaimer.Text = "[Advertising space reserverd by CocaCola]";
            this.tPage_disclaimer.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Enabled = false;
            this.button1.Location = new System.Drawing.Point(6, 263);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(450, 35);
            this.button1.TabIndex = 12;
            this.button1.Text = "7) Test GitHub";
            this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1442, 688);
            this.Controls.Add(this.tabControl_mainTabControl);
            this.Font = new System.Drawing.Font("Segoe UI Semibold", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "MainForm";
            this.Text = "Open XML Tutorial";
            this.tabControl_mainTabControl.ResumeLayout(false);
            this.tPage_Exel.ResumeLayout(false);
            this.tPage_Exel.PerformLayout();
            this.tPage_Word.ResumeLayout(false);
            this.tPage_PPT.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button Btn_GetSlidesNumber;
        private System.Windows.Forms.TextBox Tbo_ExcelLog;
        private System.Windows.Forms.Button Btn_GetExcelSheetsNames;
        private System.Windows.Forms.Button Btn_GetWordCreationDateTime;
        private System.Windows.Forms.Button Btn_GetExcelCellValue;
        private System.Windows.Forms.Button Btn_WriteInExcelCell;
        private System.Windows.Forms.Button Btn_InsertTableInWorldDocument;
        private System.Windows.Forms.Button Btn_RetrievePlainText;
        private System.Windows.Forms.TabControl tabControl_mainTabControl;
        private System.Windows.Forms.TabPage tPage_Exel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbo_CurrentExcelFile;
        private System.Windows.Forms.Button btn_CreateEmptyExcelFile;
        private System.Windows.Forms.TabPage tPage_Word;
        private System.Windows.Forms.TabPage tPage_PPT;
        private System.Windows.Forms.TabPage tPage_disclaimer;
        private System.Windows.Forms.Button btn_AddExcelSheets;
        private System.Windows.Forms.Button Btn_CreateExcelFileTableFormattedCells;
        private System.Windows.Forms.Button button1;
    }
}

