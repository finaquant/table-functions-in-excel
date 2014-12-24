// Finaquant Analytics - http://finaquant.com/
// Copyright © Finaquant Analytics GmbH
// Email: support@finaquant.com

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
// using System.IO; // required for PATH
using FinaquantCalcs;

// NetOffice: Microsoft Office integration without version limitations
// http://netoffice.codeplex.com/
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.GlobalHelperModules;
using NetOffice.ExcelApi.Enums;
using XTable = NetOffice.ExcelApi.ListObject;

// Excel DNA: Integrate .NET into Excel
// http://exceldna.codeplex.com/
using ExcelDna.Integration;
using ExcelDna.ComInterop;
using CustomUI = ExcelDna.Integration.CustomUI;

// Developer Notes
// How to programmatically add controls to Windows forms at run time by using Visual C#
// http://support2.microsoft.com/kb/319266
// Beginning C# - Chapter 13: Using Windows Form Controls
// http://www.codeproject.com/Articles/1465/Beginning-C-Chapter-Using-Windows-Form-Controls

namespace FinaquantInExcel
{
    public partial class ParameterForm1 : Form
    {
        #region "class variables"

        // form size
        private int ImageWidth = 100;    // width of tiger image (+ some margin) in form
        private int FormWidth = 450;
        private int Col1Start = 10;
        private int Col2Start = 220;
        private int LableLen = 200;
        private int LineWidth = 30;
        private int BottomMargin = 80; 

        // controls
        private ComboBox xTable1 = new ComboBox();
        private ComboBox xTable2 = new ComboBox();

        private Label xTable1Lable = new Label();
        private Label xTable2Lable = new Label();

        private ComboBox InKeyFig1 = new ComboBox();
        private ComboBox InKeyFig2 = new ComboBox();
        private ComboBox OutKeyFig1 = new ComboBox();

        private Label InKeyFig1Lable = new Label();
        private Label InKeyFig2Lable = new Label();
        private Label OutKeyFig1Lable = new Label();

        private ComboBox SheetName = new ComboBox();
        private Label SheetNameLable = new Label();

        private ComboBox comboBox1 = new ComboBox();
        private Label comboBox1Lable = new Label();

        private ComboBox comboBox2 = new ComboBox();
        private Label comboBox2Lable = new Label();

        private ComboBox comboBox3 = new ComboBox();
        private Label comboBox3Lable = new Label();   

        private Label RedWarning = new Label();

        private TextBox txtBox1 = new TextBox();
        private Label txtBox1Lable = new Label();

        private TextBox txtBox2 = new TextBox();
        private Label txtBox2Lable = new Label();

        private TextBox txtBox3 = new TextBox();
        private Label txtBox3Lable = new Label();

        private TextBox txtBox4 = new TextBox();
        private Label txtBox4Lable = new Label();

        private TextBox txtBox5 = new TextBox();
        private Label txtBox5Lable = new Label();  

        private TextBox TopLeftCell = new TextBox();
        private Label TopLeftCellLable = new Label();

        private TextBox xOutTable1 = new TextBox();
        private Label xOutTable1Lable = new Label();

        private CheckBox checkBox1 = new CheckBox();
        private Label checkBox1Lable = new Label();

        private ListBox listBox1 = new ListBox();
        private Label listBox1Lable = new Label();

        private ListBox listBox2 = new ListBox();
        private Label listBox2Lable = new Label(); 
 
        private Button button_OK = new Button();
        private Button button_Cancel = new Button();

        private RichTextBox RtxtBox1 = new RichTextBox();
        private Label RtxtBox1Lable = new Label();

        private RichTextBox RtxtBox2 = new RichTextBox();
        private Label RtxtBox2Lable = new Label();  

        // matrix
        private ComboBox Matrix1 = new ComboBox();      // range name
        private ComboBox Matrix2 = new ComboBox();

        private Label Matrix1Lable = new Label();
        private Label Matrix2Lable = new Label();

        private TextBox OutMatrix1 = new TextBox();
        private Label OutMatrix1Lable = new Label();

        private static string formbrand = " - finaquant.com";

        internal static string[] TableArithmetics_FuncTitles = new string[] {
            "Add Tables",
            "Multiply Tables",
            "Subtract Table",
            "Divide Table" 
        };

        internal static string[] TableArithmetics_FuncDescriptions = new string[] {
            "Table Addition: Add two input tables with selected key figures.\n\n"
                + "All the attributes of 2. input table (numeric & text) must be contained in 1. input table.",
            "Table Multiplication: Multiply two input tables with selected key figures.\n\n"
                + "All the attributes of 2. input table (numeric & text) must be contained in 1. input table.",
            "Table Subtraction: Subtract 2. input table from 1. input table, with selected key figures.\n\n"
             + "All the attributes of 2. input table (numeric & text) must be contained in 1. input table.",
            "Table Division: Divide 1. input table by 2. input table, with selected key figures.\n\n"
             + "All the attributes of 2. input table (numeric & text) must be contained in 1. input table."
        };

        // static fields (like global variables)
        internal static bool IfCancel;

        internal static string xTable1_st;
        internal static string xTable2_st;
        internal static string txtBox1_st;
        internal static string txtBox2_st;
        internal static string txtBox3_st;
        internal static string txtBox4_st;
        internal static string txtBox5_st;   

        internal static string RtxtBox1_st;
        internal static string RtxtBox2_st;   

        internal static string InKeyFig1_st;
        internal static string InKeyFig2_st;
        internal static string xOutTable1_st; 
        internal static string OutKeyFig_st;
        internal static string SheetName_st;
        internal static string TopLeftCell_st;
        internal static string comboBox1_st;
        internal static string comboBox2_st;
        internal static string comboBox3_st;  

        internal static bool checkBox1_st;

        internal static Excel.Workbook wbook_st;
        internal static MetaData md_st;

        internal static string Matrix1_st;
        internal static string Matrix2_st;
        internal static string OutMatrix1_st;
        internal static string[] listBox1_st;
        internal static string[] listBox2_st;

        #endregion "class variables"

        internal ParameterForm1()
        {
            InitializeComponent();

            // form
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;
            this.ForeColor = Color.Black;
            // this.Size = new System.Drawing.Size(400, 500);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;

            // RedWarning
            this.RedWarning.Location = new System.Drawing.Point(10, ImageWidth);
            this.RedWarning.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RedWarning.ForeColor = System.Drawing.Color.Red;
            this.RedWarning.Width = 480;
            this.Controls.Add(RedWarning);

            ParameterForm1.IfCancel = true;     // valid status if form is closed
        }

        #region "Parameter forms for table functions"

        // Prepare parameter form for arithmetical operations like table addition, multiplication etc.
        // 1) Name of 1. input table
        // 2) Name of 2. input table
        // 3) Selected key figure from 1. input table
        // 4) Selected key figure from 2. input table
        // 5) Resultant (output) key figure
        internal void PrepForm_TableArithmetics(string FormTitle, string FuncTitle, string FuncDescr, 
            string xOutTable, string SheetName, string TopLeftCell)
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc 

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                this.xTable1.SelectedValueChanged += new System.EventHandler(this.RefreshInKeyFigList1);
                this.xTable1.SelectedValueChanged += new System.EventHandler(this.RefreshOutKeyFigList1);

                AddLabel(this.xTable1Lable, "Select 1. Input Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

                // xTable2
                AddComboBox(this.xTable2, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                this.xTable2.SelectedValueChanged += new System.EventHandler(this.RefreshInKeyFigList2);
                this.xTable2.SelectedValueChanged += new System.EventHandler(this.RefreshOutKeyFigList1);

                AddLabel(this.xTable2Lable, "Select 2. Input Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // InKeyFig1
            AddComboBox(this.InKeyFig1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList);
            AddLabel(this.InKeyFig1Lable, "Select key figure of 1. input table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // InKeyFig2
            AddComboBox(this.InKeyFig2, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList);
            AddLabel(this.InKeyFig2Lable, "Select key figure of 2. input table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OutKeyFig
            AddComboBox(this.OutKeyFig1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown);
            AddLabel(this.OutKeyFig1Lable, "Select/Enter key figure of output table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output excel table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.TableArithmetics_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        // Prepare parameter for all arithmetical operations: Addition, Subtraction, Multiplication, Division
        // Operation type is selected within the form
        internal void PrepForm_TableArithmeticsAll(string FormTitle, string xOutTable, string SheetName, string TopLeftCell)
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = TableArithmetics_FuncTitles[0];
            this.FuncDescription.Text = TableArithmetics_FuncDescriptions[0];
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc 

            // select box with operation types
            AddComboBox(this.comboBox1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, false, 
                TableArithmetics_FuncTitles, null, 0);
            this.comboBox1.SelectedValueChanged += new System.EventHandler(this.RefreshArithmeticOpFunc);

            AddLabel(this.comboBox1Lable, "Select arithmetical operation", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                this.xTable1.SelectedValueChanged += new System.EventHandler(this.RefreshInKeyFigList1);
                this.xTable1.SelectedValueChanged += new System.EventHandler(this.RefreshOutKeyFigList1);

                AddLabel(this.xTable1Lable, "Select 1. Input Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

                // xTable2
                AddComboBox(this.xTable2, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                this.xTable2.SelectedValueChanged += new System.EventHandler(this.RefreshInKeyFigList2);
                this.xTable2.SelectedValueChanged += new System.EventHandler(this.RefreshOutKeyFigList1);

                AddLabel(this.xTable2Lable, "Select 2. Input Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // InKeyFig1
            AddComboBox(this.InKeyFig1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList);
            AddLabel(this.InKeyFig1Lable, "Select key figure of 1. input table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // InKeyFig2
            AddComboBox(this.InKeyFig2, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList);
            AddLabel(this.InKeyFig2Lable, "Select key figure of 2. input table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OutKeyFig
            AddComboBox(this.OutKeyFig1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown);
            AddLabel(this.OutKeyFig1Lable, "Select/Enter key figure of output table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output excel table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.TableArithmeticsALL_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);
            
            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        // Prepare parameter form for Combine Tables
        // 1) Name of 1. input table
        // 2) Name of 2. input table
        internal void PrepForm_CombineTables(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell)
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable1Lable, "Select 1. Input Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

                // xTable2
                AddComboBox(this.xTable2, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable2Lable, "Select 2. Input Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.CombineTables_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        // Prepare parameter form for Filter Table
        // 1) Name of 1. input table (Base)
        // 2) Name of 2. input table (Condition)
        internal void PrepForm_FilterTable(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell) 
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable1Lable, "Select 1. Input Table (Base Table)", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

                // xTable2
                AddComboBox(this.xTable2, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable2Lable, "Select 2. Input Table (Condition Table)", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // ExcludeMatchedRows
            AddCheckBox(this.checkBox1, Col2Start, ++CTR * LineWidth + ImageWidth, 200, 17,
                "ExcludeMatchedRows", "Exclude matched rows");
            AddLabel(this.checkBox1Lable, "If checked, exclude matched rows", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.FilterTable_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        // Prepare parameter form for Distribute Table
        internal void PrepForm_DistributeTable(string FormTitle, string FuncTitle, string FuncDescr,
            string ResultKeyFig, string KeySumsKeyFig, string xOutTable, string SheetName, string TopLeftCell)
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable1Lable, "Select Source Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

                // xTable2
                AddComboBox(this.xTable2, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable2Lable, "Select Key Table with Distribution Ratios", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // resultant key figure
            AddTextBox(this.txtBox1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, ResultKeyFig);
            AddLabel(this.txtBox1Lable, "Enter name of resultant key figure", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // key sums
            AddTextBox(this.txtBox2, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, KeySumsKeyFig);
            AddLabel(this.txtBox2Lable, "Enter name of key figure for key sums", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.DistributeTable_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        // Prepare parameter form for Combinate Field Values
        internal void PrepForm_CombinateFields(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell)
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable1Lable, "Select input table with field values", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.CombinateFields_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        // Prepare parameter form for Combinate Table Rows
        internal void PrepForm_CombinateRows(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell) 
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            int ListBoxWidth = 90;      // vertical with of list box

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // listBox1
                AddListBox(this.listBox1, TableList, Col2Start, ++CTR * LineWidth + ImageWidth, 150, ListBoxWidth, "TableList");
                AddLabel(this.listBox1Lable, "Select at least 2 tables", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.CombinateRows_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin + ListBoxWidth);
        }

        // Prepare parameter form for Aggregate Table
        internal void PrepForm_AggregateTable(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell) 
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                this.xTable1.SelectedValueChanged += new System.EventHandler(this.RefreshAttributeAndKeyfigLists);

                AddLabel(this.xTable1Lable, "Select input table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // add combo box for selecting aggregation function
            string[] AggrOpts = new string[] { "sum", "avg", "min", "max" };

            AddComboBox(this.comboBox1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, false, AggrOpts);
            AddLabel(this.comboBox1Lable, "Select aggregation function", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            this.comboBox1.SelectedIndex = 0;   // sum as default aggregation func

            // add list boxes for selecting attributes and key figures
            int ListBoxWidth = 50;      // vertical with of list box

            // listBox1
            AddListBox(this.listBox1, null, Col2Start, ++CTR * LineWidth + ImageWidth, 150, ListBoxWidth, "AttributeList");
            AddLabel(this.listBox1Lable, "Select at least 1 attribute", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // listBox2
            AddListBox(this.listBox2, null, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 150, ListBoxWidth, "KeyfigList");
            AddLabel(this.listBox2Lable, "Select at least 1 key figure", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth + 2 * ListBoxWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth + 2 * ListBoxWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth + 2 * ListBoxWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth + 2 * ListBoxWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth + 2 * ListBoxWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth + 2 * ListBoxWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth + 2 * ListBoxWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.AggregateTable_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth + 2 * ListBoxWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin + 2 * ListBoxWidth);
        }

        // Prepare parameter form for Applying User-Defined Transformation Function on Table Rows
        internal void PrepForm_UserDefinedTransformFuncOnRows(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell) 
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // additional form with for this particular parameter form
            int AddToFormWith = 150;  
            this.RedWarning.Width = this.RedWarning.Width + AddToFormWith;
            this.FuncDescription.Width = this.FuncDescription.Width + AddToFormWith;

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                this.xTable1.SelectedValueChanged += new System.EventHandler(this.RefreshUserDefFunctionVariables);

                AddLabel(this.xTable1Lable, "Select input table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // add text box for explanation
            int textBoxSizeY = 80;
            int textBoxSizeX = FormWidth + AddToFormWith - Col2Start - 10;
            string Expln = "You can use following variables in your function (valid C# code)\n\n";
            AddRichTextBox(this.RtxtBox1, Col2Start, ++CTR * LineWidth + ImageWidth, textBoxSizeX, textBoxSizeY, Expln);
            AddLabel(this.RtxtBox1Lable, "Explanations for user-defined function", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // add text box for user code
            AddRichTextBox(this.RtxtBox2, Col2Start, ++CTR * LineWidth + ImageWidth + textBoxSizeY, textBoxSizeX, textBoxSizeY, null, false);
            AddLabel(this.RtxtBox2Lable, "Enter your function (valid C# code)", Col1Start, CTR * LineWidth + ImageWidth + textBoxSizeY, LableLen);

            // xOutTable1 
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth + 2*textBoxSizeY, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.UserDefinedTransformFuncOnRows_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth + AddToFormWith, ImageWidth + CTR * LineWidth + BottomMargin + 2 * textBoxSizeY);
        }

        // Prepare parameter form for Selecting Columns
        internal void PrepForm_SelectColumns(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell) 
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                this.xTable1.SelectedValueChanged += new System.EventHandler(this.RefreshFieldList);

                AddLabel(this.xTable1Lable, "Select input table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // add list boxes for selecting attributes and key figures
            int ListBoxWidth = 80;      // vertical with of list box

            // listBox1
            AddListBox(this.listBox1, null, Col2Start, ++CTR * LineWidth + ImageWidth, 150, ListBoxWidth, "FieldList");
            AddLabel(this.listBox1Lable, "Select fields", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // Exclude selected fields
            AddCheckBox(this.checkBox1, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 200, 17,
                "ExcludeSelectedFields", "Exclude selected fields");
            AddLabel(this.checkBox1Lable, "If checked, exclude selected fields", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.SelectColumns_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin + ListBoxWidth);
        }

        // Prepare parameter form for Scalar Operation, like add scalar value to selected key figures
        internal void PrepForm_ScalarOperation(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell)
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                this.xTable1.SelectedValueChanged += new System.EventHandler(this.RefreshKeyfigList);

                AddLabel(this.xTable1Lable, "Select input table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // add combo box for selecting aggregation function
            string[] ScalarOp = new string[] { "Addition", "Multiplication", "Subtraction", "Division" };

            AddComboBox(this.comboBox1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, false, ScalarOp);
            AddLabel(this.comboBox1Lable, "Select arithmetic operation", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            this.comboBox1.SelectedIndex = 0;   // sum as default aggregation func

            // add list boxes for selecting attributes and key figures
            int ListBoxWidth = 50;      // vertical with of list box

            // listBox1
            AddListBox(this.listBox1, null, Col2Start, ++CTR * LineWidth + ImageWidth, 150, ListBoxWidth, "KeyfigList");
            AddLabel(this.listBox1Lable, "Select at least 1 key figure", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // add textbox for scalar value
            AddTextBox(this.txtBox1, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20);
            AddLabel(this.txtBox1Lable, "Enter a scalar number like 2.5", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.ScalarOperation_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin + ListBoxWidth);
        }

        // Prepare parameter form for rounding selected key figures
        internal void PrepForm_RoundNumbers(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell) 
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                this.xTable1.SelectedValueChanged += new System.EventHandler(this.RefreshKeyfigList);

                AddLabel(this.xTable1Lable, "Select input table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // add combo box for selecting #digits after decimal point
            string[] RoundDigits = new string[] { "0", "1", "2", "3", "4", "5", "6" };

            AddComboBox(this.comboBox1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, false, RoundDigits);
            AddLabel(this.comboBox1Lable, "Select number of digits after decimal point", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            this.comboBox1.SelectedIndex = 2;   // sum as default aggregation func

            // add list boxes for selecting attributes and key figures
            int ListBoxWidth = 50;      // vertical with of list box

            // listBox1
            AddListBox(this.listBox1, null, Col2Start, ++CTR * LineWidth + ImageWidth, 150, ListBoxWidth, "KeyfigList");
            AddLabel(this.listBox1Lable, "Select at least 1 key figure", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.RoundNumbers_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin + ListBoxWidth);
        }

        // Prepare parameter form for inserting a new field into input table
        internal void PrepForm_InsertNewField(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell) 
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable1Lable, "Select input table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // field name
            AddTextBox(this.txtBox1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20);
            AddLabel(this.txtBox1Lable, "Enter field's name", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // add combo box for field type
            string[] FieldType = new string[] { "Text Attribute", "Date Attribute", "Integer Attribute", "Key Figure" };

            AddComboBox(this.comboBox1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, false, FieldType);
            this.comboBox1.SelectedValueChanged += new System.EventHandler(this.RefreshDefaultFieldValue);
            
            AddLabel(this.comboBox1Lable, "Select field type", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // field name
            AddTextBox(this.txtBox2, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, xOutTable);
            AddLabel(this.txtBox2Lable, "Enter a valid field value", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.InsertNewField_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        // Prepare parameter form for assigning random values to key figures and numeric attributes
        // with lower and upper limits
        internal void PrepForm_AssignRandomNumbers(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell)
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                this.xTable1.SelectedValueChanged += new System.EventHandler(this.RefreshNumberFieldList);

                AddLabel(this.xTable1Lable, "Select input table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // add list boxes for selecting attributes and key figures
            int ListBoxWidth = 90;      // vertical with of list box

            // listBox1
            AddListBox(this.listBox1, null, Col2Start, ++CTR * LineWidth + ImageWidth, 150, ListBoxWidth, "KeyfigList");
            AddLabel(this.listBox1Lable, "Select at least 1 field", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // lower limit
            AddTextBox(this.txtBox1, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, "0.0");
            AddLabel(this.txtBox1Lable, "Enter a number for lower limit", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // upper limit
            AddTextBox(this.txtBox2, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, "100.0");
            AddLabel(this.txtBox2Lable, "Enter a number for higher limit", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.AssignRandomNumbers_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin + ListBoxWidth);
        }

        // Prepare parameter form for sorting rows of input table
        internal void PrepForm_SortRows(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell) 
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable1Lable, "Select input table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // add text box for user code
            int textBoxSizeY = 50;
            int textBoxSizeX = FormWidth - Col2Start - 20;
            AddRichTextBox(this.RtxtBox1, Col2Start, ++CTR * LineWidth + ImageWidth, textBoxSizeX, textBoxSizeY, null, false);
            AddLabel(this.RtxtBox1Lable, "Enter field names and sort options", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth + textBoxSizeY, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth + textBoxSizeY, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth + textBoxSizeY, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth + textBoxSizeY, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth + textBoxSizeY, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth + textBoxSizeY, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth + textBoxSizeY, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.SortRows_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth + textBoxSizeY, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin + textBoxSizeY);
        }

        // prepare parameter form for aggregate selected key figure
        internal void PrepForm_InsertAggregateKeyFigure(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell) 
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                this.xTable1.SelectedValueChanged += new System.EventHandler(this.RefreshAttributeListAndKeyfigComboBox);

                AddLabel(this.xTable1Lable, "Select input table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // add combo box for selecting aggregation function
            string[] AggrOpts = new string[] { "sum", "avg", "min", "max" };

            AddComboBox(this.comboBox1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, false, AggrOpts);
            AddLabel(this.comboBox1Lable, "Select aggregation function", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            this.comboBox1.SelectedIndex = 0;   // sum as default aggregation func

            AddComboBox(this.InKeyFig1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, false);
            AddLabel(this.InKeyFig1Lable, "Select name of input key figure", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            this.comboBox1.SelectedIndex = 0;   // sum as default aggregation func

            // add list boxes for selecting attributes and key figures
            int ListBoxWidth = 50;      // vertical with of list box

            // listBox1
            AddListBox(this.listBox1, null, Col2Start, ++CTR * LineWidth + ImageWidth, 150, ListBoxWidth, "AttributeList");
            AddLabel(this.listBox1Lable, "Select at least 1 reference attribute", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // inserted new aggregate key figure
            AddTextBox(this.txtBox1, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20);
            AddLabel(this.txtBox1Lable, "Enter name of aggregate key figure", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.InsertAggregateKeyFigure_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin + ListBoxWidth);
        }

        // Prepare parameter form for Applying User-Defined Filter Function on Table Rows
        internal void PrepForm_UserDefinedFilterFuncOnRows(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell) 
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // additional form with for this particular parameter form
            int AddToFormWith = 150;
            this.RedWarning.Width = this.RedWarning.Width + AddToFormWith;
            this.FuncDescription.Width = this.FuncDescription.Width + AddToFormWith;

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                this.xTable1.SelectedValueChanged += new System.EventHandler(this.RefreshUserDefFunctionVariables);

                AddLabel(this.xTable1Lable, "Select input table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // add text box for explanation
            int textBoxSizeY = 80;
            int textBoxSizeX = FormWidth + AddToFormWith - Col2Start - 10;
            string Expln = "You can use following variables in your function (valid C# code)\n\n";
            AddRichTextBox(this.RtxtBox1, Col2Start, ++CTR * LineWidth + ImageWidth, textBoxSizeX, textBoxSizeY, Expln);
            AddLabel(this.RtxtBox1Lable, "Explanations for user-defined function", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // add text box for user code
            AddRichTextBox(this.RtxtBox2, Col2Start, ++CTR * LineWidth + ImageWidth + textBoxSizeY, textBoxSizeX, textBoxSizeY, null, false);
            AddLabel(this.RtxtBox2Lable, "Enter your function (valid C# code)", Col1Start, CTR * LineWidth + ImageWidth + textBoxSizeY, LableLen);

            // Exclude
            AddCheckBox(this.checkBox1, Col2Start, ++CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, 200, 17,
                "ExcludeRows", "Exclude rows");
            AddLabel(this.checkBox1Lable, "If checked, exclude rows", Col1Start, CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, LableLen);

            // xOutTable1 
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.UserDefinedFilterFuncOnRows_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth + 2 * textBoxSizeY, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth + AddToFormWith, ImageWidth + CTR * LineWidth + BottomMargin + 2 * textBoxSizeY);
        }

        // Prepare parameter form for appending two tables vertically
        internal void PrepForm_AppendRows(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell) 
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable1Lable, "Select 1. Input Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

                // xTable2
                AddComboBox(this.xTable2, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable2Lable, "Select 2. Input Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.AppendRows_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        // Prepare parameter form for appending two tables horizontally
        internal void PrepForm_AppendColumns(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell) 
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable1Lable, "Select 1. Input Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

                // xTable2
                AddComboBox(this.xTable2, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable2Lable, "Select 2. Input Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.AppendColumns_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        // Prepare parameter form for simple date range filter
        internal void PrepForm_SimpleDateRangeFilter(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell) 
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                this.xTable1.SelectedValueChanged += new System.EventHandler(this.RefreshDateFieldList);
                
                AddLabel(this.xTable1Lable, "Select input Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // date field of table
            AddComboBox(this.comboBox1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true);
            AddLabel(this.comboBox1Lable, "Select date field as basis for filtering", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // FirstDayOfRange
            AddTextBox(this.txtBox1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20);
            AddLabel(this.txtBox1Lable, "First day of range", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // LastDayOfRange
            AddTextBox(this.txtBox2, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20);
            AddLabel(this.txtBox2Lable, "Last day of range", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // Exclude
            AddCheckBox(this.checkBox1, Col2Start, ++CTR * LineWidth + ImageWidth, 200, 17,
                "ExcludeRange", "Exclude date range");
            AddLabel(this.checkBox1Lable, "If checked, exclude date range", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.SimpleDateRangeFilter_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        // Prepare parameter form for date filter
        internal void PrepForm_DateFilter(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell) 
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                this.xTable1.SelectedValueChanged += new System.EventHandler(this.RefreshDateFieldList);

                AddLabel(this.xTable1Lable, "Select input Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // date field of table
            AddComboBox(this.comboBox1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true);
            AddLabel(this.comboBox1Lable, "Select date field as basis for filtering", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // period
            string[] Periods = new string[] { "Month", "Quarter", "Year" };
            AddComboBox(this.comboBox2, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, false, Periods);
            AddLabel(this.comboBox2Lable, "Select period", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            this.comboBox2.SelectedIndex = 0;

            // FirstDayOfRange
            AddTextBox(this.txtBox1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20);
            AddLabel(this.txtBox1Lable, "Enter first day of range", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // LastDayOfRange
            AddTextBox(this.txtBox2, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20);
            AddLabel(this.txtBox2Lable, "Enter last day of range", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // allowed week days
            int ListBoxWidth = 80;
            string[] WeekDays = new string[] { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday" };
            AddListBox(this.listBox1, WeekDays, Col2Start, ++CTR * LineWidth + ImageWidth, 150, ListBoxWidth, "WeekDayList", false);
            AddLabel(this.listBox1Lable, "Select allowed week days", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // allowed period days
            AddTextBox(this.txtBox3, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 200, 20);
            AddLabel(this.txtBox3Lable, "Enter allowed period days", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.DateFilter_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin + ListBoxWidth);
        }

        // Prepare parameter form for date sampler
        internal void PrepForm_DateSampler(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell) 
        {
            // init form
            int ExtraColWidth = 60;
            this.Col2Start += ExtraColWidth;
            this.FormWidth += ExtraColWidth;
            this.FuncDescription.Width += ExtraColWidth;

            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                this.xTable1.SelectedValueChanged += new System.EventHandler(this.RefreshDateFieldList);

                AddLabel(this.xTable1Lable, "Select input Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // source date
            AddComboBox(this.comboBox1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true);
            AddLabel(this.comboBox1Lable, "Select source date's field name", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // target date
            AddTextBox(this.txtBox1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20);
            AddLabel(this.txtBox1Lable, "Enter target date's field name", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // period
            string[] Periods = new string[] { "Month", "Quarter", "Year" };
            AddComboBox(this.comboBox2, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, false, Periods);
            AddLabel(this.comboBox2Lable, "Select period", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            this.comboBox2.SelectedIndex = 0;

            // FirstDayOfRange
            AddTextBox(this.txtBox2, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20);
            AddLabel(this.txtBox2Lable, "Enter first day of range", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // LastDayOfRange
            AddTextBox(this.txtBox3, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20);
            AddLabel(this.txtBox3Lable, "Enter last day of range", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // search logic
            string[] SearchOpts= new string[] { "Previous date", "Next date", "Nearest date", "Exact date" };
            AddComboBox(this.comboBox3, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, false, SearchOpts);
            AddLabel(this.comboBox3Lable, "Select search logic", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            this.comboBox3.SelectedIndex = 0;

            // Max day-distance
            AddTextBox(this.txtBox4, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, "90");
            AddLabel(this.txtBox4Lable, @"Max days betw. source and target dates", Col1Start, CTR * LineWidth + ImageWidth, LableLen + ExtraColWidth);

            // allowed week days
            int ListBoxWidth = 80;
            string[] WeekDays = new string[] { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday" };
            AddListBox(this.listBox1, WeekDays, Col2Start, ++CTR * LineWidth + ImageWidth, 150, ListBoxWidth, "WeekDayList", false);
            AddLabel(this.listBox1Lable, "Select allowed week days", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // allowed period days
            AddTextBox(this.txtBox5, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 200, 20);
            AddLabel(this.txtBox5Lable, "Enter allowed period days (like 1,15,-1 ..)", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen + ExtraColWidth);

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth + ListBoxWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.DateSampler_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth + ListBoxWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin + ListBoxWidth);
        }

        // Prepare parameter form for get price table (example for user-defined table function)
        internal void PrepForm_GetPriceTable2(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable, string SheetName, string TopLeftCell)
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // xTable1
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable1Lable, "Select Cost Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

                // xTable2
                AddComboBox(this.xTable2, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable2Lable, "Select Margin Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, xOutTable);
            AddLabel(this.xOutTable1Lable, "Enter name of output table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.GetPriceTable2_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        // Prepare parameter form for calculating sales commissions
        internal void PrepForm_CalculateSalesCom(string FormTitle, string FuncTitle, string FuncDescr,
            string xOutTable1, string xOutTable2, string SheetName, string TopLeftCell)
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all excel tables
            string[] TableList;
            ExcelFunc_NO.GetAllExcelTables(wbook_st, out TableList);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (TableList != null && TableList.Count() > 0)
            {
                // sales table
                AddComboBox(this.xTable1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable1Lable, "Select Sales Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

                // sales commission table
                AddComboBox(this.xTable2, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.xTable2Lable, "Select Commission-Scale Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

                // product pool table
                AddComboBox(this.comboBox1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.comboBox1Lable, "Select Product-Pool Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

                // scale-to-pool table
                AddComboBox(this.comboBox2, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, TableList);
                AddLabel(this.comboBox2Lable, "Select Scale-To-Pool Table", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No table was found in excel workbook!";
                return;
            }

            // select calculation & payment period
            string[] Periods = new string[] { "month", "quarter" };
            AddComboBox(this.comboBox3, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, Periods, DefaultIndex: 1);
            AddLabel(this.comboBox3Lable, "Select payment period", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
 
            // xOutTable1
            AddTextBox(this.xOutTable1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, xOutTable1);
            AddLabel(this.xOutTable1Lable, "Enter name of 1. output table (ComPerPool)", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // txtBox2
            AddTextBox(this.txtBox1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, xOutTable2);
            AddLabel(this.txtBox1Lable, "Enter name of 2. output table (ComPerDealer)", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.CalculateSalesCom_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        #endregion "Parameter forms for table functions"

        #region "Parameter forms for matrix functions"

        // prepare parameter form for Create Random Matrix
        internal void PrepForm_CreateRandomMatrix(string FormTitle, string FuncTitle, string FuncDescr,
            string OutMatrix, string SheetName, string TopLeftCell)
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all range names
            string[] RangeNames = ExcelFunc_NO.GetAllRangeNames(wbook_st);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // txtBox1
            AddTextBox(this.txtBox1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20);
            AddLabel(this.txtBox1Lable, "Enter row count", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // txtBox2
            AddTextBox(this.txtBox2, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20);
            AddLabel(this.txtBox2Lable, "Enter column count", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OutMatrix1
            AddTextBox(this.OutMatrix1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, "Matrix1");
            AddLabel(this.OutMatrix1Lable, "Enter range name for output matrix", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.CreateRandomMatrix_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        // Prepare parameter form for binary matrix-matrix operations like Matrix Addition, Matrix Multiplication
        internal void PrepForm_BinaryMatrixOp(string FormTitle, string FuncTitle, string FuncDescr,
            string OutMatrix, string SheetName, string TopLeftCell)
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all range names
            string[] RangeNames = ExcelFunc_NO.GetAllRangeNames(wbook_st);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (RangeNames != null && RangeNames.Count() > 0)
            {
                // xTable1
                AddComboBox(this.Matrix1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, RangeNames);
                AddLabel(this.Matrix1Lable, "Select range name for 1. input matrix", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

                // xTable2
                AddComboBox(this.Matrix2, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, RangeNames);
                AddLabel(this.Matrix2Lable, "Select range name for 2. input matrix", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No range name for a matrix was found in excel workbook!";
                return;
            }

            // OutMatrix1
            AddTextBox(this.OutMatrix1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, OutMatrix);
            AddLabel(this.OutMatrix1Lable, "Enter range name for output matrix", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.BinaryMatrixOp_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        // Prepare parameter form for single matrix operations like Transpose, Inverse..
        internal void PrepForm_SingleMatrixOp(string FormTitle, string FuncTitle, string FuncDescr,
            string OutMatrix, string SheetName, string TopLeftCell) 
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all range names
            string[] RangeNames = ExcelFunc_NO.GetAllRangeNames(wbook_st);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (RangeNames != null && RangeNames.Count() > 0)
            {
                // xTable1
                AddComboBox(this.Matrix1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, RangeNames);
                AddLabel(this.Matrix1Lable, "Select range name for 1. input matrix", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No range name for a matrix was found in excel workbook!";
                return;
            }

            // OutMatrix1
            AddTextBox(this.OutMatrix1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, OutMatrix);
            AddLabel(this.OutMatrix1Lable, "Enter range name for output matrix", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.SingleMatrixOp_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        // Prepare parameter form for matrix-scalar operations like
        // add scalar to matrix, multiply matrix with scalar
        internal void PrepForm_MatrixScalarOp(string FormTitle, string FuncTitle, string FuncDescr,
            string OutMatrix, string SheetName, string TopLeftCell)
        {
            // init form
            this.Text = FormTitle + formbrand;
            this.FuncTitle.Text = FuncTitle;
            this.FuncDescription.Text = FuncDescr;
            int CTR = 0;  // number of form elements like combo boxes, text boxes etc

            // get list of all range names
            string[] RangeNames = ExcelFunc_NO.GetAllRangeNames(wbook_st);

            // get list of all worksheets
            string[] sheets = ExcelFunc_NO.GetAllSheetNames(wbook_st);

            // initiate combo boxes
            if (RangeNames != null && RangeNames.Count() > 0)
            {
                // xTable1
                AddComboBox(this.Matrix1, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDownList, true, RangeNames);
                AddLabel(this.Matrix1Lable, "Select range name for input matrix", Col1Start, CTR * LineWidth + ImageWidth, LableLen);
            }
            else
            {
                this.RedWarning.Text = "No range name for a matrix was found in excel workbook!";
                return;
            }

            // txtBox1 for scalar
            AddTextBox(this.txtBox1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20);
            AddLabel(this.txtBox1Lable, "Enter scalar value", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OutMatrix1
            AddTextBox(this.OutMatrix1, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, OutMatrix);
            AddLabel(this.OutMatrix1Lable, "Enter range name for output matrix", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // SheetName
            AddComboBox(this.SheetName, Col2Start, ++CTR * LineWidth + ImageWidth, ComboBoxStyle.DropDown, true, sheets, SheetName);
            AddLabel(this.SheetNameLable, "Select/Enter sheet name for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // text box (UpperLeftCell)
            AddTextBox(this.TopLeftCell, Col2Start, ++CTR * LineWidth + ImageWidth, 100, 20, TopLeftCell);
            AddLabel(this.TopLeftCellLable, "Address of upper-left cell for output", Col1Start, CTR * LineWidth + ImageWidth, LableLen);

            // OK button
            AddButton(this.button_OK, Col2Start, ++CTR * LineWidth + ImageWidth, 75, 23, "OK", "OK");
            this.button_OK.Click += new System.EventHandler(this.MatrixScalarOp_OKbutton);

            // CANCEL button
            AddButton(this.button_Cancel, Col2Start + 90, CTR * LineWidth + ImageWidth, 75, 23, "CANCEL", "CANCEL");
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);

            // form size
            this.Size = new System.Drawing.Size(FormWidth, ImageWidth + CTR * LineWidth + BottomMargin);
        }

        #endregion "Parameter forms for matrix functions"

        #region "Event handlers"

        private void TableArithmetics_OKbutton(object sender, EventArgs e)
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "1. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xTable2.SelectedItem == null)
            {
                this.RedWarning.Text = "2. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.InKeyFig1.SelectedItem == null)
            {
                this.RedWarning.Text = "Key figure of 1. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.InKeyFig2.SelectedItem == null)
            {
                this.RedWarning.Text = "Key figure of 2. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.OutKeyFig1.SelectedItem == null && (this.OutKeyFig1.Text == null || this.OutKeyFig1.Text.Trim() == ""))
            {
                this.RedWarning.Text = "Key figure of output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            xTable2_st = this.xTable2.SelectedItem.ToString();
            InKeyFig1_st = this.InKeyFig1.SelectedItem.ToString();
            InKeyFig2_st = this.InKeyFig2.SelectedItem.ToString();
            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.OutKeyFig1.SelectedItem == null)
                OutKeyFig_st = this.OutKeyFig1.Text.Trim();
            else
                OutKeyFig_st = this.OutKeyFig1.SelectedItem.ToString();

            if (! MetaData.CheckIfProperFieldName(OutKeyFig_st))
            {
                this.RedWarning.Text = "Improper key figure name '" + OutKeyFig_st + "'!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void TableArithmeticsALL_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "1. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xTable2.SelectedItem == null)
            {
                this.RedWarning.Text = "2. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.InKeyFig1.SelectedItem == null)
            {
                this.RedWarning.Text = "Key figure of 1. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.InKeyFig2.SelectedItem == null)
            {
                this.RedWarning.Text = "Key figure of 2. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.OutKeyFig1.SelectedItem == null && (this.OutKeyFig1.Text == null || this.OutKeyFig1.Text.Trim() == ""))
            {
                this.RedWarning.Text = "Key figure of output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            comboBox1_st = this.comboBox1.SelectedItem.ToString();
            xTable1_st = this.xTable1.SelectedItem.ToString();
            xTable2_st = this.xTable2.SelectedItem.ToString();
            InKeyFig1_st = this.InKeyFig1.SelectedItem.ToString();
            InKeyFig2_st = this.InKeyFig2.SelectedItem.ToString();
            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.OutKeyFig1.SelectedItem == null)
                OutKeyFig_st = this.OutKeyFig1.Text.Trim();
            else
                OutKeyFig_st = this.OutKeyFig1.SelectedItem.ToString();

            if (!MetaData.CheckIfProperFieldName(OutKeyFig_st))
            {
                this.RedWarning.Text = "Improper key figure name '" + OutKeyFig_st + "'!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void CombineTables_OKbutton(object sender, EventArgs e)
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "1. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xTable2.SelectedItem == null)
            {
                this.RedWarning.Text = "2. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xTable2.SelectedItem.ToString() == this.xTable1.SelectedItem.ToString())
            {
                this.RedWarning.Text = "Two different tables with one or more common attributes must be selected.";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            xTable2_st = this.xTable2.SelectedItem.ToString();
            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void FilterTable_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "1. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xTable2.SelectedItem == null)
            {
                this.RedWarning.Text = "2. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            xTable2_st = this.xTable2.SelectedItem.ToString();
            xOutTable1_st = this.xOutTable1.Text.Trim();
            checkBox1_st = this.checkBox1.Checked;

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void DistributeTable_OKbutton(object sender, EventArgs e)
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "1. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xTable2.SelectedItem == null)
            {
                this.RedWarning.Text = "2. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xTable1.SelectedItem == this.xTable2.SelectedItem)
            {
                this.RedWarning.Text = "Source and Key tables cannot be the same!";
                this.Refresh();
                return;
            }

            if (this.txtBox1.Text == null || this.txtBox1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of resultant key figure must be entered!";
                this.Refresh();
                return;
            }

            if (this.txtBox2.Text == null || this.txtBox2.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of key figure for key sums must be entered!";
                this.Refresh();
                return;
            }

            if (this.txtBox1.Text.Trim() == this.txtBox2.Text.Trim())
            {
                this.RedWarning.Text = "Resultant and KeySum key figures cannot be the same!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // Set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            xTable2_st = this.xTable2.SelectedItem.ToString();
            txtBox1_st = this.txtBox1.Text.Trim();
            txtBox2_st = this.txtBox2.Text.Trim();

            if (!MetaData.CheckIfProperFieldName(txtBox1_st))
            {
                this.RedWarning.Text = "Improper key figure name '" + txtBox1_st + "'!";
                this.Refresh();
                return;
            }

            if (!MetaData.CheckIfProperFieldName(txtBox2_st))
            {
                this.RedWarning.Text = "Improper key figure name '" + txtBox2_st + "'!";
                this.Refresh();
                return;
            }

            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void CombinateFields_OKbutton(object sender, EventArgs e)
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "1. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void CombinateRows_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.listBox1.SelectedItems == null || this.listBox1.SelectedItems.Count < 2)
            {
                this.RedWarning.Text = "At least 2 tables must be selected for row combination!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            listBox1_st = new string[this.listBox1.SelectedItems.Count];

            int i = 0;
            foreach (var str in this.listBox1.SelectedItems)
            {
                listBox1_st[i++] = str.ToString().Trim();
            }
            
            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void AggregateTable_OKbutton(object sender, EventArgs e)
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "1. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.listBox1.SelectedItems == null || this.listBox1.SelectedItems.Count < 1)
            {
                this.RedWarning.Text = "At least 1 attribute of input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.listBox2.SelectedItems == null || this.listBox2.SelectedItems.Count < 1)
            {
                this.RedWarning.Text = "At least 1 key figure of input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            comboBox1_st = this.comboBox1.SelectedItem.ToString();

            listBox1_st = new string[this.listBox1.SelectedItems.Count];

            int i = 0;
            foreach (var str in this.listBox1.SelectedItems)
            {
                listBox1_st[i++] = str.ToString().Trim();
            }

            listBox2_st = new string[this.listBox2.SelectedItems.Count];

            i = 0;
            foreach (var str in this.listBox2.SelectedItems)
            {
                listBox2_st[i++] = str.ToString().Trim();
            }

            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void UserDefinedTransformFuncOnRows_OKbutton(object sender, EventArgs e)
        {
            // validate inputs
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.RtxtBox2.Text == null || this.RtxtBox2.Text.Trim() == "")
            {
                this.RedWarning.Text = "A user-defined function (valid C# code) must be entered!";
                this.Refresh();
                return;
            }

            // check validity of user code
            string UserCode = this.RtxtBox2.Text;
            string ErrorStr;

            if (!HelperFunc.CheckUserDefinedTransformFunction(UserCode, out ErrorStr))
            {
                this.RedWarning.Text = ErrorStr;
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            RtxtBox2_st = this.RtxtBox2.Text;

            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void SelectColumns_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            // exclude fields if checked
            bool Ifchecked = this.checkBox1.Checked;    

            if (! Ifchecked && (this.listBox1.SelectedItems == null || this.listBox1.SelectedItems.Count < 1))
            {
                this.RedWarning.Text = "At least 1 field of input table must be selected!";
                this.Refresh();
                return;
            }

            if (Ifchecked && (this.listBox1.SelectedItems != null && this.listBox1.SelectedItems.Count == this.listBox1.Items.Count))
            {
                this.RedWarning.Text = "All fields cannot be selected with exclude option!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            checkBox1_st = this.checkBox1.Checked;

            listBox1_st = new string[this.listBox1.SelectedItems.Count];

            int i = 0;
            foreach (var str in this.listBox1.SelectedItems)
            {
                listBox1_st[i++] = str.ToString().Trim();
            }

            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void ScalarOperation_OKbutton(object sender, EventArgs e)
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.listBox1.SelectedItems == null || this.listBox1.SelectedItems.Count < 1)
            {
                this.RedWarning.Text = "At least 1 key figure must be selected!";
                this.Refresh();
                return;
            }

            if (this.txtBox1.Text == null || this.txtBox1.Text.Trim() == "")
            {
                this.RedWarning.Text = "A scalar value must be entered!";
                this.Refresh();
                return;
            }

            double N;

            // check if double
            if (!double.TryParse(this.txtBox1.Text.Trim(), out N))
            {
                this.RedWarning.Text = "A scalar number like 5.25 must be entered!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            txtBox1_st = this.txtBox1.Text.Trim();
            comboBox1_st = this.comboBox1.SelectedItem.ToString();

            listBox1_st = new string[this.listBox1.SelectedItems.Count];

            int i = 0;
            foreach (var str in this.listBox1.SelectedItems)
            {
                listBox1_st[i++] = str.ToString().Trim();
            }

            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void RoundNumbers_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.listBox1.SelectedItems == null || this.listBox1.SelectedItems.Count < 1)
            {
                this.RedWarning.Text = "At least 1 key figure must be selected!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            comboBox1_st = this.comboBox1.SelectedItem.ToString();

            listBox1_st = new string[this.listBox1.SelectedItems.Count];

            int i = 0;
            foreach (var str in this.listBox1.SelectedItems)
            {
                listBox1_st[i++] = str.ToString().Trim();
            }

            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void InsertNewField_OKbutton(object sender, EventArgs e)
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            // check field name
            if (this.txtBox1.Text == null || this.txtBox1.Text.Trim() == "")
            {
                this.RedWarning.Text = "A field name must be entered!";
                this.Refresh();
                return;
            }

            string fieldname = this.txtBox1.Text.Trim().ToLower();

            if (!MetaData.CheckIfProperFieldName(fieldname))
            {

                this.RedWarning.Text = "Improper field name '" + fieldname + "'!";
                this.Refresh();
                return;
            }

            // get table name
            string InTbl1 = this.xTable1.SelectedItem.ToString();

            // get fields of table
            TableFields tfields = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl1, md_st);
            TextVector AllFields = tfields.ColumnNames;

            if (TextVector.IfValueFoundInSet(fieldname, AllFields))
            {
                this.RedWarning.Text = "Field '" + fieldname + "' already exists in input table!";
                this.Refresh();
                return;
            }

            // check field type
            if (this.comboBox1.SelectedItem == null)
            {
                this.RedWarning.Text = "A field type must be selected!";
                this.Refresh();
                return;
            }

            // check field value
            if (this.txtBox2.Text == null || this.txtBox2.Text.Trim() == "")
            {
                this.RedWarning.Text = "A field value must be entered!";
                this.Refresh();
                return;
            }

            string FieldValue = this.txtBox2.Text.Trim();
            string FieldType = this.comboBox1.SelectedItem.ToString();

            if (FieldType == "Date Attribute")
            {
                DateTime dt;
                if (!DateTime.TryParse(FieldValue, out dt))
                {
                    this.RedWarning.Text = "Invalid value for a Date Attribute!";
                    this.Refresh();
                    return;
                }

            }
            else if (FieldType == "Integer Attribute")
            {
                int x;
                if (!int.TryParse(FieldValue, out x))
                {
                    this.RedWarning.Text = "Invalid value for an Integer Attribute!";
                    this.Refresh();
                    return;
                }
            }
            else if (FieldType == "Key Figure")
            {
                double x;
                if (!double.TryParse(FieldValue, out x))
                {
                    this.RedWarning.Text = "Invalid value for an Key Figure!";
                    this.Refresh();
                    return;
                }
            }


            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            comboBox1_st = this.comboBox1.SelectedItem.ToString();

            txtBox1_st = this.txtBox1.Text.Trim();
            txtBox2_st = this.txtBox2.Text.Trim();

            int i = 0;
            foreach (var str in this.listBox1.SelectedItems)
            {
                listBox1_st[i++] = str.ToString().Trim();
            }

            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void AssignRandomNumbers_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.listBox1.SelectedItems == null || this.listBox1.SelectedItems.Count < 1)
            {
                this.RedWarning.Text = "At least 1 field must be selected!";
                this.Refresh();
                return;
            }

            if (this.txtBox1.Text == null || this.txtBox1.Text.Trim() == "")
            {
                this.RedWarning.Text = "A lower limit must be entered!";
                this.Refresh();
                return;
            }

            double x;

            if (! double.TryParse(this.txtBox1.Text.Trim(), out x))
            {
                this.RedWarning.Text = "Lower limit must be a valid number like 5 or 5.25!";
                this.Refresh();
                return;
            }

            if (this.txtBox2.Text == null || this.txtBox2.Text.Trim() == "")
            {
                this.RedWarning.Text = "A higher limit must be entered!";
                this.Refresh();
                return;
            }

            if (! double.TryParse(this.txtBox2.Text.Trim(), out x))
            {
                this.RedWarning.Text = "Higher limit must be a valid number like 5 or 5.25!";
                this.Refresh();
                return;
            }

            if (double.Parse(this.txtBox1.Text.Trim()) > double.Parse(this.txtBox2.Text.Trim()))
            {
                this.RedWarning.Text = "Lower limit cannot be larger than upper limit!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            listBox1_st = new string[this.listBox1.SelectedItems.Count];

            int i = 0;
            foreach (var str in this.listBox1.SelectedItems)
            {
                listBox1_st[i++] = str.ToString().Trim();
            }

            txtBox1_st = this.txtBox1.Text.Trim();
            txtBox2_st = this.txtBox2.Text.Trim();

            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void SortRows_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            RtxtBox1_st = this.RtxtBox1.Text.Trim();

            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void InsertAggregateKeyFigure_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.InKeyFig1.SelectedItem == null)
            {
                this.RedWarning.Text = "A key figure of input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.listBox1.SelectedItems == null || this.listBox1.SelectedItems.Count < 1)
            {
                this.RedWarning.Text = "At least 1 reference attribute of input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.txtBox1.Text == null || this.txtBox1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of aggregate (output) key figure must be entered!";
                this.Refresh();
                return;
            }

            string OutputKeyfig = this.txtBox1.Text.Trim();

            // get table name
            string InTbl1 = this.xTable1.SelectedItem.ToString();

            // get fields of table
            TableFields tfields = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl1, md_st);

            // check if proper field name
            if (!MetaData.CheckIfProperFieldName(OutputKeyfig))
            {

                this.RedWarning.Text = "Improper field name '" + OutputKeyfig + "'!";
                this.Refresh();
                return;
            }

            // check if a field with the same name already exists
            TextVector AllFields = tfields.ColumnNames;

            if (TextVector.IfValueFoundInSet(OutputKeyfig, AllFields))
            {
                this.RedWarning.Text = "Field '" + OutputKeyfig + "' already exists in input table!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            comboBox1_st = this.comboBox1.SelectedItem.ToString();  // aggregation func
            InKeyFig1_st = this.InKeyFig1.SelectedItem.ToString();  // input key figure
            txtBox1_st = OutputKeyfig;

            listBox1_st = new string[this.listBox1.SelectedItems.Count];
            int i = 0;
            foreach (var str in this.listBox1.SelectedItems)
            {
                listBox1_st[i++] = str.ToString().Trim();
            }

            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void UserDefinedFilterFuncOnRows_OKbutton(object sender, EventArgs e) 
        {
            // validate inputs
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.RtxtBox2.Text == null || this.RtxtBox2.Text.Trim() == "")
            {
                this.RedWarning.Text = "A user-defined filter function (valid C# code) must be entered!";
                this.Refresh();
                return;
            }

            // check validity of user code
            string UserCode = this.RtxtBox2.Text;
            string ErrorStr;

            if (!HelperFunc.CheckUserDefinedFilterFunction(UserCode, out ErrorStr))
            {
                this.RedWarning.Text = ErrorStr;
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            RtxtBox2_st = this.RtxtBox2.Text;
            checkBox1_st = checkBox1.Checked;

            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void AppendRows_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "1. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xTable2.SelectedItem == null)
            {
                this.RedWarning.Text = "2. input table must be selected!";
                this.Refresh();
                return;
            }

            string Tbl1 = this.xTable1.SelectedItem.ToString();
            string Tbl2 = this.xTable2.SelectedItem.ToString();

            // get fields of table
            TableFields tfields1 = ExcelFunc_NO.ReadTableFieldsFromExcel(Tbl1, md_st);
            TableFields tfields2 = ExcelFunc_NO.ReadTableFieldsFromExcel(Tbl2, md_st);

            if (!TextVector.IsEqual(tfields1.ColumnNames, tfields2.ColumnNames, true))
            {
                this.RedWarning.Text = "Tables must have identical fields for appending them vertically!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            xTable2_st = this.xTable2.SelectedItem.ToString();
            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void AppendColumns_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "1. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xTable2.SelectedItem == null)
            {
                this.RedWarning.Text = "2. input table must be selected!";
                this.Refresh();
                return;
            }

            string Tbl1 = this.xTable1.SelectedItem.ToString();
            string Tbl2 = this.xTable2.SelectedItem.ToString();

            // get fields of table
            TableFields tfields1 = ExcelFunc_NO.ReadTableFieldsFromExcel(Tbl1, md_st);
            TableFields tfields2 = ExcelFunc_NO.ReadTableFieldsFromExcel(Tbl2, md_st);

            TextVector CommonFields = TextVector.Intersection(tfields1.ColumnNames, tfields2.ColumnNames);

            if (CommonFields != null && CommonFields.nLength > 0)
            {
                this.RedWarning.Text = "Tables must have no common fields for appending columns!" ;
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            xTable2_st = this.xTable2.SelectedItem.ToString();
            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void SimpleDateRangeFilter_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.comboBox1.SelectedItem == null)
            {
                this.RedWarning.Text = "A date field of table must be selected!";
                this.Refresh();
                return;
            }

            if (this.txtBox1.Text == null || this.txtBox1.Text.Trim() == "")
            {
                this.RedWarning.Text = "First day of range must be entered!";
                this.Refresh();
                return;
            }

            DateTime dt;
            if (!DateTime.TryParse(this.txtBox1.Text.Trim(), out dt))
            {
                this.RedWarning.Text = "FirstDayOfRange must be a valid date like 24.08.2012";
                this.Refresh();
                return;
            }

            if (this.txtBox2.Text == null || this.txtBox2.Text.Trim() == "")
            {
                this.RedWarning.Text = "Last day of range must be entered!";
                this.Refresh();
                return;
            }

            if (!DateTime.TryParse(this.txtBox2.Text.Trim(), out dt))
            {
                this.RedWarning.Text = "LastDayOfRange must be a valid date like 24.08.2012";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            comboBox1_st = this.comboBox1.SelectedItem.ToString();

            txtBox1_st = this.txtBox1.Text.Trim();
            txtBox2_st = this.txtBox2.Text.Trim();
            checkBox1_st = this.checkBox1.Checked;

            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void DateFilter_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.comboBox1.SelectedItem == null)
            {
                this.RedWarning.Text = "A date field of table must be selected!";
                this.Refresh();
                return;
            }

            if (this.txtBox1.Text == null || this.txtBox1.Text.Trim() == "")
            {
                this.RedWarning.Text = "First day of range must be entered!";
                this.Refresh();
                return;
            }

            DateTime dt;
            if (!DateTime.TryParse(this.txtBox1.Text.Trim(), out dt))
            {
                this.RedWarning.Text = "FirstDayOfRange must be a valid date like 24.08.2012";
                this.Refresh();
                return;
            }

            if (this.txtBox2.Text == null || this.txtBox2.Text.Trim() == "")
            {
                this.RedWarning.Text = "Last day of range must be entered!";
                this.Refresh();
                return;
            }

            if (!DateTime.TryParse(this.txtBox2.Text.Trim(), out dt))
            {
                this.RedWarning.Text = "LastDayOfRange must be a valid date like 24.08.2012";
                this.Refresh();
                return;
            }

            // allowed period days
            string Period = this.comboBox2.SelectedItem.ToString();
            string AllowedDaysStr = this.txtBox3.Text.Trim();

            if (AllowedDaysStr != null && AllowedDaysStr != "")
            {
                // remove all whitespaces
                AllowedDaysStr = Regex.Replace(AllowedDaysStr, @"\s+", "");

                // match string
                if (!Regex.IsMatch(AllowedDaysStr, @"[\d+,]*\d+"))
                {
                    this.RedWarning.Text = "Comma-separated allowed period-days must be entered like: 1, 15, -1";
                    this.Refresh();
                    return;
                }

                string[] AllowedDays = AllowedDaysStr.Split(',');
                int day;
                foreach (var AllowedDay in AllowedDays)
                {
                    if (!int.TryParse(AllowedDay, out day))
                    {
                        this.RedWarning.Text = "Invalid integer number as allowed period-day: " + day;
                        this.Refresh();
                        return;
                    }
                    if (Period == "Month" && Math.Abs(day) > 28)
                    {
                        this.RedWarning.Text = "Absolute value of a period-day cannot be larger than 28!";
                        this.Refresh();
                        return;
                    }
                    if (Period == "Quarter" && Math.Abs(day) > 90)
                    {
                        this.RedWarning.Text = "Absolute value of a period-day cannot be larger than 90!";
                        this.Refresh();
                        return;
                    }
                    if (Period == "Year" && Math.Abs(day) > 365)
                    {
                        this.RedWarning.Text = "Absolute value of a period-day cannot be larger than 365!";
                        this.Refresh();
                        return;
                    }
                }
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            comboBox1_st = this.comboBox1.SelectedItem.ToString();  // date field
            comboBox2_st = this.comboBox2.SelectedItem.ToString();  // period

            txtBox1_st = this.txtBox1.Text.Trim();      // first day of range
            txtBox2_st = this.txtBox2.Text.Trim();      // last day of range
            txtBox3_st = AllowedDaysStr;                // allowed period days

            // allowed week days
            listBox1_st = new string[this.listBox1.SelectedItems.Count];
            int i = 0;
            foreach (var str in this.listBox1.SelectedItems)
            {
                listBox1_st[i++] = str.ToString().Trim();
            }

            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void DateSampler_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.comboBox1.SelectedItem == null)
            {
                this.RedWarning.Text = "A date field of table must be selected!";
                this.Refresh();
                return;
            }

            if (this.txtBox1.Text == null || this.txtBox1.Text.Trim() == "")
            {
                this.RedWarning.Text = "A field name for target date must be entered!";
                this.Refresh();
                return;
            }

            // check target date
            string TargetDate = this.txtBox1.Text.Trim();

            if (!MetaData.CheckIfProperFieldName(TargetDate))
            {

                this.RedWarning.Text = "Improper field name '" + TargetDate + "'!";
                this.Refresh();
                return;
            }

            // get table name
            string InTbl1 = this.xTable1.SelectedItem.ToString();

            // get fields of table
            TableFields tfields = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl1, md_st);
            TextVector AllFields = tfields.ColumnNames;

            if (TextVector.IfValueFoundInSet(TargetDate, AllFields))
            {
                this.RedWarning.Text = "Field '" + TargetDate + "' already exists in input table!";
                this.Refresh();
                return;
            }

            // FirstDayOfRange
            if (this.txtBox2.Text == null || this.txtBox2.Text.Trim() == "")
            {
                this.RedWarning.Text = "First day of range must be entered!";
                this.Refresh();
                return;
            }

            DateTime dt;
            if (!DateTime.TryParse(this.txtBox2.Text.Trim(), out dt))
            {
                this.RedWarning.Text = "FirstDayOfRange must be a valid date like 24.08.2012";
                this.Refresh();
                return;
            }

            // LastDayOfRange
            if (this.txtBox3.Text == null || this.txtBox3.Text.Trim() == "")
            {
                this.RedWarning.Text = "Last day of range must be entered!";
                this.Refresh();
                return;
            }

            if (!DateTime.TryParse(this.txtBox3.Text.Trim(), out dt))
            {
                this.RedWarning.Text = "LastDayOfRange must be a valid date like 24.08.2012";
                this.Refresh();
                return;
            }

            // Max day-distance
            if (this.txtBox4.Text == null || this.txtBox4.Text.Trim() == "")
            {
                this.RedWarning.Text = "Max day-distance must be entered!";
                this.Refresh();
                return;
            }

            int dist;
            if (!int.TryParse(this.txtBox4.Text.Trim(), out dist))
            {
                this.RedWarning.Text = "Max day-distance must be a valid integer like 90";
                this.Refresh();
                return;
            }

            // allowed period days
            string Period = this.comboBox2.SelectedItem.ToString();
            string AllowedDaysStr = this.txtBox5.Text.Trim();

            if (AllowedDaysStr != null && AllowedDaysStr != "")
            {
                // remove all whitespaces
                AllowedDaysStr = Regex.Replace(AllowedDaysStr, @"\s+", "");

                // match string
                if (!Regex.IsMatch(AllowedDaysStr, @"[\d+,]*\d+"))
                {
                    this.RedWarning.Text = "Comma-separated allowed period-days must be entered like: 1, 15, -1";
                    this.Refresh();
                    return;
                }

                string[] AllowedDays = AllowedDaysStr.Split(',');
                int day;
                foreach (var AllowedDay in AllowedDays)
                {
                    if (!int.TryParse(AllowedDay, out day))
                    {
                        this.RedWarning.Text = "Invalid integer number as allowed period-day: " + day;
                        this.Refresh();
                        return;
                    }
                    if (Period == "Month" && Math.Abs(day) > 28)
                    {
                        this.RedWarning.Text = "Absolute value of a period-day cannot be larger than 28!";
                        this.Refresh();
                        return;
                    }
                    if (Period == "Quarter" && Math.Abs(day) > 90)
                    {
                        this.RedWarning.Text = "Absolute value of a period-day cannot be larger than 90!";
                        this.Refresh();
                        return;
                    }
                    if (Period == "Year" && Math.Abs(day) > 365)
                    {
                        this.RedWarning.Text = "Absolute value of a period-day cannot be larger than 365!";
                        this.Refresh();
                        return;
                    }
                }
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            comboBox1_st = this.comboBox1.SelectedItem.ToString();  // source date
            comboBox2_st = this.comboBox2.SelectedItem.ToString();  // period
            comboBox3_st = this.comboBox3.SelectedItem.ToString();  // search logic

            txtBox1_st = this.txtBox1.Text.Trim();  // target date
            txtBox2_st = this.txtBox2.Text.Trim();      // first day of range
            txtBox3_st = this.txtBox3.Text.Trim();      // last day of range
            txtBox4_st = this.txtBox4.Text.Trim();      // max day-distance
            txtBox5_st = AllowedDaysStr;                // allowed period days

            // allowed week days
            listBox1_st = new string[this.listBox1.SelectedItems.Count];
            int i = 0;
            foreach (var str in this.listBox1.SelectedItems)
            {
                listBox1_st[i++] = str.ToString().Trim();
            }

            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void GetPriceTable2_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "A cost table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xTable2.SelectedItem == null)
            {
                this.RedWarning.Text = "A margin table must be selected!";
                this.Refresh();
                return;
            }

            // check if cost table has a key figure named "costs"
            string InTbl1 = this.xTable1.SelectedItem.ToString();
            TableFields tfields1 = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl1, md_st);

            if (! TextVector.IfValueFoundInSet("costs", tfields1.KeyFigures))
            {
                this.RedWarning.Text = "Selected cost table must contain a key figure named costs";
                this.Refresh();
                return;
            }

            // check if margin table has a key figure named "margin"
            string InTbl2 = this.xTable2.SelectedItem.ToString();
            TableFields tfields2 = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl2, md_st);

            if (!TextVector.IfValueFoundInSet("margin", tfields2.KeyFigures))
            {
                this.RedWarning.Text = "Selected margin table must contain a key figure named margin";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();
            xTable2_st = this.xTable2.SelectedItem.ToString();

            xOutTable1_st = this.xOutTable1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void CalculateSalesCom_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "A Sales Table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xTable2.SelectedItem == null)
            {
                this.RedWarning.Text = "A Commission-Scale table must be selected!";
                this.Refresh();
                return;
            }

            if (this.comboBox1.SelectedItem == null)
            {
                this.RedWarning.Text = "A Product-Pool table must be selected!";
                this.Refresh();
                return;
            }

            if (this.comboBox2.SelectedItem == null)
            {
                this.RedWarning.Text = "A Scale-To-Pool table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of 1. output table (Commissions Per Pool) must be entered!";
                this.Refresh();
                return;
            }

            if (this.xOutTable1.Text == null || this.xOutTable1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of 1. output table (Commissions Per Pool) must be entered!";
                this.Refresh();
                return;
            }

            if (this.txtBox1.Text == null || this.txtBox1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Name of 2. output table (Commissions Per Dealer) must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output tables must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            xTable1_st = this.xTable1.SelectedItem.ToString();      // sales table
            xTable2_st = this.xTable2.SelectedItem.ToString();      // commission-scale
            comboBox1_st = this.comboBox1.SelectedItem.ToString();  // product-pool
            comboBox2_st = this.comboBox2.SelectedItem.ToString();  // scale-to-pool

            comboBox3_st = this.comboBox3.SelectedItem.ToString(); // period

            xOutTable1_st = this.xOutTable1.Text.Trim();    // commission per pool
            txtBox1_st = this.txtBox1.Text.Trim();          // commission per dealer

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        // matrix

        private void CreateRandomMatrix_OKbutton(object sender, EventArgs e)
        {
            // validate input
            if (this.txtBox1.Text == null || this.txtBox1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Row count of matrix must be entered!";
                this.Refresh();
                return;
            }

            int N;

            // check if integer
            if ( ! int.TryParse(this.txtBox1.Text.Trim(), out N) )
            {
                this.RedWarning.Text = "An integer value must be entered for row count!";
                this.Refresh();
                return;
            }

            if (this.txtBox2.Text == null || this.txtBox2.Text.Trim() == "")
            {
                this.RedWarning.Text = "Column count of matrix must be entered!";
                this.Refresh();
                return;
            }

            // check if integer
            if (! int.TryParse(this.txtBox2.Text.Trim(), out N))
            {
                this.RedWarning.Text = "An integer value must be entered for row count!";
                this.Refresh();
                return;
            }

            if (this.OutMatrix1.Text == null || this.OutMatrix1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Range Name of output  matrix must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            txtBox1_st = this.txtBox1.Text.Trim();
            txtBox2_st = this.txtBox2.Text.Trim();

            OutMatrix1_st = this.OutMatrix1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void BinaryMatrixOp_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.Matrix1.SelectedItem == null)
            {
                this.RedWarning.Text = "Range name for 1. input matrix must be selected!";
                this.Refresh();
                return;
            }

            if (this.Matrix2.SelectedItem == null)
            {
                this.RedWarning.Text = "Range name for 2. input matrix must be selected!";
                this.Refresh();
                return;
            }

            if (this.OutMatrix1.Text == null || this.OutMatrix1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Range Name of output  matrix must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            Matrix1_st = this.Matrix1.SelectedItem.ToString();
            Matrix2_st = this.Matrix2.SelectedItem.ToString();

            OutMatrix1_st = this.OutMatrix1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void SingleMatrixOp_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.Matrix1.SelectedItem == null)
            {
                this.RedWarning.Text = "Range name for 1. input matrix must be selected!";
                this.Refresh();
                return;
            }

            if (this.OutMatrix1.Text == null || this.OutMatrix1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Range Name of output  matrix must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            Matrix1_st = this.Matrix1.SelectedItem.ToString();
            OutMatrix1_st = this.OutMatrix1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        private void MatrixScalarOp_OKbutton(object sender, EventArgs e) 
        {
            // validate input
            if (this.Matrix1.SelectedItem == null)
            {
                this.RedWarning.Text = "Range name for 1. input matrix must be selected!";
                this.Refresh();
                return;
            }

            if (this.txtBox1.Text == null || this.txtBox1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Scalar value must be entered!";
                this.Refresh();
                return;
            }

            double N;

            // check if integer
            if (!double.TryParse(this.txtBox1.Text.Trim(), out N))
            {
                this.RedWarning.Text = "A scalar number like 5.25 must be entered!";
                this.Refresh();
                return;
            }

            if (this.OutMatrix1.Text == null || this.OutMatrix1.Text.Trim() == "")
            {
                this.RedWarning.Text = "Range Name of output  matrix must be entered!";
                this.Refresh();
                return;
            }

            if (this.SheetName.SelectedItem == null && (this.SheetName.Text == null || this.SheetName.Text.Trim() == ""))
            {
                this.RedWarning.Text = "A worksheet for output table must be selected or entered!";
                this.Refresh();
                return;
            }

            if (this.TopLeftCell.Text == null || this.TopLeftCell.Text.Trim() == "")
            {
                this.RedWarning.Text = "A upper-left cell address like A1 for output table must be entered!";
                this.Refresh();
                return;
            }

            // checks OK, set static variables
            Matrix1_st = this.Matrix1.SelectedItem.ToString();
            txtBox1_st = this.txtBox1.Text.Trim();
            OutMatrix1_st = this.OutMatrix1.Text.Trim();

            if (this.SheetName.SelectedItem == null)
                SheetName_st = this.SheetName.Text.Trim();
            else
                SheetName_st = this.SheetName.SelectedItem.ToString();

            TopLeftCell_st = this.TopLeftCell.Text.Trim();
            IfCancel = false;

            // close form
            this.Close();
        }

        // refresh combo-box for key figure list
        private void RefreshInKeyFigList1(object sender, EventArgs e)
        {
            this.RedWarning.Text = "";

            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "1. input table must be selected!";
                this.Refresh();
                return;
            }

            // get table name
            string InTbl1 = this.xTable1.SelectedItem.ToString();

            // get fields of table
            TableFields tfields = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl1, md_st);

            // get key figures of input table
            this.InKeyFig1.Items.Clear();

            if (tfields.KeyFiguresStr.Count() > 0)
            {
                string[] keyfigs = tfields.KeyFiguresStr;

                // fill combo-box
                foreach (var item in keyfigs)
                {
                    this.InKeyFig1.Items.Add(item);
                }

                // select first item as default
                this.InKeyFig1.SelectedIndex = 0;
            }
            else
            {
                this.RedWarning.Text = "Table '" + InTbl1 + "' has no key figures!";
            }

            this.Refresh();
        }

        // refresh combo-box for key figure list
        private void RefreshInKeyFigList2(object sender, EventArgs e)
        {
            this.RedWarning.Text = "";

            // validate input
            if (this.xTable2.SelectedItem == null)
            {
                this.RedWarning.Text = "2. input table must be selected!";
                this.Refresh();
                return;
            }

            // get table name
            string InTbl2 = this.xTable2.SelectedItem.ToString();

            // get fields of table
            TableFields tfields = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl2, md_st);

            // get key figures of input table
            this.InKeyFig2.Items.Clear();

            if (tfields.KeyFiguresStr.Count() > 0)
            {
                string[] keyfigs = tfields.KeyFiguresStr;

                // fill combo-box
                foreach (var item in keyfigs)
                {
                    this.InKeyFig2.Items.Add(item);
                }
                // select first item as default
                this.InKeyFig2.SelectedIndex = 0;
            }
            else
            {
                this.RedWarning.Text = "Table '" + InTbl2 + "' has no key figures!";
            }

            this.Refresh();
        }

        // refresh combo-box for output key figure
        private void RefreshOutKeyFigList1(object sender, EventArgs e) 
        {
            var KeyFigList = new List<string>();

            // get table names
            string InTbl1 = null;
            string InTbl2 = null;
            TableFields tfields = null;

            // get key figures of xTable1
            if (this.xTable1.SelectedItem != null)
            {
                InTbl1 = this.xTable1.SelectedItem.ToString();

                // get fields of table
                tfields = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl1, md_st);

                if (tfields.KeyFiguresStr.Count() > 0)
                {
                    KeyFigList.AddRange(tfields.KeyFiguresStr);
                }
            }

            // get key figures of xTable2
            if (this.xTable2.SelectedItem != null)
            {
                InTbl2 = this.xTable2.SelectedItem.ToString();

                // get fields of table
                tfields = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl2, md_st);

                if (tfields.KeyFiguresStr.Count() > 0)
                {
                    KeyFigList.AddRange(tfields.KeyFiguresStr);
                }
            }

            // fill combo-box
            this.OutKeyFig1.Items.Clear();

            foreach (var item in KeyFigList)
            {
                this.OutKeyFig1.Items.Add(item);
            }

            this.Refresh();
        }

        // refresh combo-box for arithmetic operation type
        private void RefreshArithmeticOpFunc(object sender, EventArgs e)
        {
            int SelectInd = this.comboBox1.SelectedIndex;

            this.FuncTitle.Text = TableArithmetics_FuncTitles[SelectInd];
            this.FuncDescription.Text = TableArithmetics_FuncDescriptions[SelectInd];
            this.Refresh();
        }

        // refresh attribute and key figure lists in list box (Aggregate Table)
        private void RefreshAttributeAndKeyfigLists(object sender, EventArgs e)
        {
            this.RedWarning.Text = "";

            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            // get table name
            string InTbl1 = this.xTable1.SelectedItem.ToString();

            // get fields of table
            TableFields tfields = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl1, md_st);

            // get attributes of input table
            this.listBox1.Items.Clear();

            if ((tfields.TextAttributes.nLength + tfields.NumAttributes.nLength) > 0)
            {
                TextVector Attributes = TextVector.Union(tfields.TextAttributes, tfields.NumAttributes);
                string[] attribs = Attributes.toArray;

                // fill combo-box
                foreach (var item in attribs)
                {
                    this.listBox1.Items.Add(item);
                }

            }
            else
            {
                this.RedWarning.Text = "Input table has no attributes!";
            }

            // get key figures of input table
            this.listBox2.Items.Clear();

            if (tfields.KeyFigures.nLength > 0)
            {
                string[] keyfigs = tfields.KeyFigures.toArray;

                // fill combo-box
                foreach (var item in keyfigs)
                {
                    this.listBox2.Items.Add(item);
                }

            }
            else
            {
                this.RedWarning.Text = "Input table has no key figures!";
            }

            this.Refresh();
        }

        // refresh explanations for user-defined function
        private void RefreshUserDefFunctionVariables(object sender, EventArgs e)
        {
            this.RedWarning.Text = "";

            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            // get table name
            string InTbl1 = this.xTable1.SelectedItem.ToString();

            // get fields of table
            TableFields tfields = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl1, md_st);

            // obtain variable info
            string TextAttribVars;
            string NumAttribVars;
            string KeyFigureVars;

            // text attributes
            if (tfields.TextAttributesStr != null && tfields.TextAttributesStr.Count() > 0)
            {
                TextAttribVars = string.Join(@"""] TA[""", tfields.TextAttributesStr);
                TextAttribVars = @"TA[""" + TextAttribVars + @"""]";
            }
            else
                TextAttribVars = "";

            // numeric attributes
            if (tfields.NumAttributesStr != null && tfields.NumAttributesStr.Count() > 0)
            {
                NumAttribVars = string.Join(@"""] NA[""", tfields.NumAttributesStr);
                NumAttribVars = @"NA[""" + NumAttribVars + @"""]";
            }
            else
                NumAttribVars = "";

            // key figures
            if (tfields.KeyFiguresStr != null && tfields.KeyFiguresStr.Count() > 0)
            {
                KeyFigureVars = string.Join(@"""] KF[""", tfields.KeyFiguresStr);
                KeyFigureVars = @"KF[""" + KeyFigureVars + @"""]";
            }
            else
                KeyFigureVars = "";

            // append to explanation
            string Expln = "You can use following variables in your function (valid C# code):\n\n"
                + "Text Attributes: " + TextAttribVars + "\n"
                + "Numeric Attributes: " + NumAttribVars + "\n"
                + "Key Figures: " + KeyFigureVars + "\n";

            this.RtxtBox1.Text = Expln;
            this.Refresh();
        }

        // refresh field list for select columns
        private void RefreshFieldList(object sender, EventArgs e)
        {
            this.RedWarning.Text = "";

            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            // get table name
            string InTbl1 = this.xTable1.SelectedItem.ToString();

            // get fields of table
            TableFields tfields = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl1, md_st);

            // get attributes of input table
            this.listBox1.Items.Clear();

            if (tfields.ColumnCount > 0)
            {
                string[] fields = tfields.ColumnNamesStr;

                // fill combo-box
                foreach (var item in fields)
                {
                    this.listBox1.Items.Add(item);
                }

            }
            else
            {
                this.RedWarning.Text = "Input table has no fields!";
            }

            this.Refresh();
        }

        // refresh key figure list for scalar operation
        private void RefreshKeyfigList(object sender, EventArgs e)
        {
            this.RedWarning.Text = "";

            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            // get table name
            string InTbl1 = this.xTable1.SelectedItem.ToString();

            // get fields of table
            TableFields tfields = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl1, md_st);

            // get attributes of input table
            this.listBox1.Items.Clear();

            if (tfields.KeyFiguresStr.Count() > 0)
            {
                string[] KeyFigs = tfields.KeyFiguresStr;

                // fill combo-box
                foreach (var item in KeyFigs)
                {
                    this.listBox1.Items.Add(item);
                }

            }
            else
            {
                this.RedWarning.Text = "Input table has no key figures!";
            }

            this.Refresh();
        }

        // refresh default field value for insert new field
        private void RefreshDefaultFieldValue(object sender, EventArgs e) 
        {
            this.RedWarning.Text = "";

            // get field type
            string FieldType = this.comboBox1.SelectedItem.ToString();

            if (FieldType == "Text Attribute")
                this.txtBox2.Text = "Red";

            else if (FieldType == "Date Attribute")
                this.txtBox2.Text = "15.10.2012";

            else if (FieldType == "Integer Attribute")
                this.txtBox2.Text = "2012";

            else if (FieldType == "Key Figure")
                this.txtBox2.Text = "5.25";

            else
                this.RedWarning.Text = "Unknown field type";

            this.Refresh();
        }

        // refresh list for key figures and numeric attributes (date & integer)
        private void RefreshNumberFieldList(object sender, EventArgs e)
        {
            this.RedWarning.Text = "";

            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            // get table name
            string InTbl1 = this.xTable1.SelectedItem.ToString();

            // get fields of table
            TableFields tfields = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl1, md_st);

            // get attributes of input table
            this.listBox1.Items.Clear();

            if ((tfields.KeyFigures.nLength + tfields.NumAttributes.nLength) > 0)
            {
                TextVector Attributes = TextVector.Union(tfields.NumAttributes, tfields.KeyFigures);
                string[] NumberFields = Attributes.toArray;

                // fill combo-box
                foreach (var item in NumberFields)
                {
                    this.listBox1.Items.Add(item);
                }
            }
            else
            {
                this.RedWarning.Text = "Input table has no numeric attributes or key figures!";
            }

            this.Refresh();
        }

        // refresh attribute list and key figure combo box
        private void RefreshAttributeListAndKeyfigComboBox(object sender, EventArgs e) 
        {
            this.RedWarning.Text = "";

            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            // get table name
            string InTbl1 = this.xTable1.SelectedItem.ToString();

            // get fields of table
            TableFields tfields = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl1, md_st);

            // get attributes of input table
            this.listBox1.Items.Clear();

            if ((tfields.TextAttributes.nLength + tfields.NumAttributes.nLength) > 0)
            {
                TextVector Attributes = TextVector.Union(tfields.TextAttributes, tfields.NumAttributes);
                string[] attribs = Attributes.toArray;

                // fill combo-box
                foreach (var item in attribs)
                {
                    this.listBox1.Items.Add(item);
                }

            }
            else
            {
                this.RedWarning.Text = "Input table has no attributes!";
            }

            // get key figures of input table
            this.InKeyFig1.Items.Clear();

            if (tfields.KeyFigures.nLength > 0)
            {
                string[] keyfigs = tfields.KeyFigures.toArray;

                // fill combo-box
                foreach (var item in keyfigs)
                {
                    this.InKeyFig1.Items.Add(item);
                }

            }
            else
            {
                this.RedWarning.Text = "Input table has no key figures!";
            }

            this.Refresh();
        }

        // refresh date field list in combo box
        private void RefreshDateFieldList(object sender, EventArgs e) 
        {
            this.RedWarning.Text = "";

            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "An input table must be selected!";
                this.Refresh();
                return;
            }

            // get table name
            string InTbl1 = this.xTable1.SelectedItem.ToString();

            // get fields of table
            TableFields tfields = ExcelFunc_NO.ReadTableFieldsFromExcel(InTbl1, md_st);

            // get attributes of input table
            this.comboBox1.Items.Clear();

            if (tfields.NumAttributesStr.Count() > 0)
            {
                string[] NumAttribs = tfields.NumAttributesStr;

                // fill combo-box
                foreach (var field in NumAttribs)
                {
                    if (md_st.GetFieldType(field) == FieldType.DateAttribute)
                        this.comboBox1.Items.Add(field);
                }
            }

            if (tfields.NumAttributesStr.Count() == 0 || this.comboBox1.Items.Count == 0)
            {
                this.RedWarning.Text = "Input table has no date fields!";
            }

            this.Refresh();
        }


        private void ParameterForm1_Load(object sender, EventArgs e)
        {
            /*
            string ProjectPath = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Directory.GetCurrentDirectory()));

            ProjectPath = @"C:\Users\Tuncali\Documents\Visual Studio 2010\FinaquantInExcel\FinaquantInExcel";
            
            string ImageLoc = ProjectPath + @"\Images\" + @"FinaquantCalcsTiger_80x80.jpg";
            this.pictureBox1.ImageLocation = ImageLoc;

            this.FuncDescription.Text = ImageLoc;
             * */
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button_OK_Click(object sender, EventArgs e)
        {
            // validate input
            if (this.xTable1.SelectedItem == null)
            {
                this.RedWarning.Text = "1. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.xTable2.SelectedItem == null)
            {
                this.RedWarning.Text = "2. input table must be selected!";
                this.Refresh();
                return;
            }

            if (this.txtBox1.Text == null || this.txtBox1.Text == "")
            {
                this.RedWarning.Text = "Name of output table must be entered!";
                this.Refresh();
                return;
            }

            string InTbl1 = this.xTable1.SelectedItem.ToString();
            string InTbl2 = this.xTable2.SelectedItem.ToString();
            string OutTblName = this.txtBox1.Text.Trim();

            // call table function
            //var xt = new ExcelTable();
            //xt.GetPriceTable(xCostTblName: InTbl1, xMarginTblName: InTbl2, xPriceTblName: OutTblName);

            // set static variables
            xTable1_st = InTbl1;
            xTable2_st = InTbl2;
            txtBox1_st = OutTblName;
            IfCancel = false;

            // close form
            this.Close();
        }

        private void button_Cancel_Click(object sender, EventArgs e)
        {
            IfCancel = true;
            this.Close();
        }

        #endregion "Event handlers"

        #region "Helper methods"

        // add combo-box to form
        private void AddComboBox(ComboBox cbox, int xLocation, int yLocation,
            ComboBoxStyle cstyle, bool IfSorted = true, string[] ValueList = null, 
            string DefaultValue = null, int DefaultIndex = -1)
        {
            cbox.Sorted = IfSorted;
            cbox.Location = new System.Drawing.Point(xLocation, yLocation);
            cbox.DropDownStyle = cstyle;

            if (ValueList != null && ValueList.Count() > 0)
            {
                foreach (var item in ValueList)
                {
                    cbox.Items.Add(item);
                }
            }

            if (DefaultIndex != -1)
            {
                cbox.SelectedIndex = DefaultIndex;
            }

            // DefaultValue is ignored if DefaultIndex != -1
            if (cstyle != ComboBoxStyle.DropDownList && DefaultValue != null
                && DefaultIndex == -1)
            {
                cbox.Text = DefaultValue;
            }

            this.Controls.Add(cbox);
        }

        // add lable to form
        private void AddLabel(Label lbl, string txt,
            int xLocation, int yLocation, int width)
        {
            lbl.Text = txt;
            lbl.Location = new System.Drawing.Point(xLocation, yLocation);
            lbl.Width = width;
            this.Controls.Add(lbl);
        }

        // add text box to form
        private void AddTextBox(TextBox txtbox,  
            int xLocation, int yLocation, int xSize, int ySize, string DefaultText = null)
        {
            txtbox.Location = new System.Drawing.Point(xLocation, yLocation);
            txtbox.Size = new System.Drawing.Size(xSize, ySize);

            if (DefaultText != null) txtbox.Text = DefaultText;

            this.Controls.Add(txtbox);
        }

        // add rich text box to form
        private void AddRichTextBox(RichTextBox Rtxtbox,
            int xLocation, int yLocation, int xSize, int ySize, string DefaultText = null, bool IfReadOnly = true)
        {
            Rtxtbox.Location = new System.Drawing.Point(xLocation, yLocation);
            Rtxtbox.Size = new System.Drawing.Size(xSize, ySize);
            Rtxtbox.ReadOnly = IfReadOnly;

            if (DefaultText != null) Rtxtbox.Text = DefaultText;

            this.Controls.Add(Rtxtbox);
        }

        // add button to form
        private void AddButton(Button btn, int xLocation, int yLocation, int xSize, int ySize,
            string name, string text)
        {
            btn.Location = new System.Drawing.Point(xLocation, yLocation);
            btn.Size = new System.Drawing.Size(xSize, ySize);
            btn.Name = name;
            btn.Text = text;
            btn.UseVisualStyleBackColor = true;
            this.Controls.Add(btn);
        }

        // add check box to form
        private void AddCheckBox(CheckBox cbx, int xLocation, int yLocation, int xSize, int ySize,
            string name, string text, bool IfCheckedByDefault = false)
        {
            cbx.AutoSize = true;
            cbx.Location = new System.Drawing.Point(xLocation, yLocation);
            cbx.Size = new System.Drawing.Size(xSize, ySize);
            cbx.Name = name;
            cbx.Text = text;
            cbx.Checked = IfCheckedByDefault;
            cbx.UseVisualStyleBackColor = true;
            this.Controls.Add(cbx);
        }

        // add list box to form
        private void AddListBox(ListBox listbx, string[] ListValues, int xLocation, int yLocation, int xSize, int ySize,
            string name, bool IfSorted = true, SelectionMode SelectMode = SelectionMode.MultiSimple, 
            bool IfFormattingEnabled = false) 
        {
            listbx.FormattingEnabled = IfFormattingEnabled;
            listbx.Location = new System.Drawing.Point(xLocation, yLocation);
            listbx.Size = new System.Drawing.Size(xSize, ySize);
            listbx.Name = name;
            listbx.SelectionMode = SelectMode;

            if (ListValues != null && ListValues.Count() > 0)
            {
                foreach (var str in ListValues)
                {
                    listbx.Items.Add(str);
                } 
            }

            listbx.Sorted = IfSorted;
            this.Controls.Add(listbx);
        }




        #endregion "Helper methods"

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void FuncDescription_TextChanged(object sender, EventArgs e)
        {

        }

    }
}
