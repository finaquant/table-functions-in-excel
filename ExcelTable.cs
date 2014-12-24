// Finaquant Analytics - http://finaquant.com/
// Copyright © Finaquant Analytics GmbH
// Email: support@finaquant.com

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Text.RegularExpressions;
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

// Calling a non-static method of class ExcelTable in VBA
/*
' Calling ExcelTbl.GetPriceTable in Excel VBA
Sub Test_GetPriceTable()
Dim ExcelTbl As Object: Set ExcelTbl = CreateObject("Finaquant_ExcelTable")

' Call .NET method with parameters
Call ExcelTbl.GetPriceTable("Cost", "Margin1")

' Call .NET method without parameters (macro)
Call ExcelTbl.GetPriceTable_macro

End Sub
*/

namespace FinaquantInExcel
{
    /// <summary>
    /// Class with application oriented table-valued functions
    /// for excel users and VBA programmers. 
    /// Input and Output parameters are generally excel tables (ListObject).
    /// All public and non-static methods in this class are available in Excel VBA.
    /// We will use the class name XTable as synonim for Excel.ListObject
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("Finaquant_ExcelTable")]
    public class ExcelTable
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public ExcelTable() { }

        // Simple test method without a return value (void) for testing excel integration
        public void AddNumbers(double x, double y)
        {
            double z = x + y;
            // Console.WriteLine("x + y = " + z);
            MessageBox.Show("x + y = " + z.ToString());
        }

        // Simple test method with a return value (double) for testing excel integration
        public double MultiplyNumbers(double x, double y)
        {
            return x * y;
        }

        /// <summary>
        /// Read an excel table (ListObject) into a MatrixTableX
        /// </summary>
        /// <param name="xTbl">Excel table (ListObject)</param>
        /// <param name="mdx">Meta data with field definitions</param>
        /// <param name="TextReplaceNull">Replacement integer value for null in excel table</param>
        /// <param name="NumReplaceNull">Replacement integer value for null in excel table</param>
        /// <param name="KeyFigReplaceNull">Replacement floating value for null in excel table</param>
        /// <returns>Table of type MatrixTableX</returns>
        public MatrixTableX ExcelToMatrixTable(XTable xTbl, MetaDataX mdx,
            string TextReplaceNull = "NULL", int NumReplaceNull = 0, double KeyFigReplaceNull = 0)
        {
            MatrixTableX tbl = new MatrixTableX();
            tbl.ReadFromExcelTable(xTbl, mdx, TextReplaceNull, NumReplaceNull, KeyFigReplaceNull);
            return tbl;
        }

        /// <summary>
        /// Create an excel table (ListObject) by writing MatrixTableX to its range in worksheet.
        /// </summary>
        /// <param name="Tbl">Table of type MatrixTableX</param>
        /// <param name="WSheet">Worksheet object</param>
        /// <param name="TableName">Name of excel table</param>
        /// <param name="CellStr">Upper-left corner of Excel Table in worksheet</param>
        /// <param name="ClearSheetContent">If true, clear whole sheet </param>
        /// <returns>Excel table (ListObject)</returns>
        /// <remarks>
        /// Calling this method from excel VBA requires MS Interop (not NetOffice interop)
        /// because this method has an excel object parameter (worksheet). 
        /// </remarks>
        public XTable MatrixTableToExcel(MatrixTableX Tbl, Excel.Worksheet WSheet, string TableName, 
            string CellStr = "A1", bool ClearSheetContent = false)
        {
            return Tbl.WriteToExcelTable(WSheet, TableName, CellStr, ClearSheetContent);
        }

        /// <summary>
        /// Create test tables in excel sheets
        /// </summary>
        public void CreateTestTables_macro()
        {
            try
            {
                if (MessageBox.Show(
                    "Creating test tables may overwrite existing tables. Is this OK for you?",
                    "Warning!", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {

                    // get active workbook
                    Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();
                    Excel.Workbook wbook = xlapp.ActiveWorkbook;
                    ExcelFunc_NO.CreateTestTables(wbook);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("CreateTestTables: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Obtain Price table from Cost and Margin tables.
        /// Excel Tables are passed as ListObject parameters.
        /// Underlying Table Function: MatrixTable.MultiplySelectedKeyFigures()
        /// </summary>
        /// <param name="xCostTbl">ListObject (excel table) for costs</param>
        /// <param name="xMarginTbl">ListObject (excel table) for margins</param>
        /// <param name="xMetaTbl">ListObject (excel table) for field definitions (metadata)</param>
        /// <param name="WSheet">Worksheet object for resultant price table</param>
        /// <param name="PriceTblName">Name of resultant price table (ListObject)</param>
        /// <param name="inKeyFig1"></param>
        /// <param name="inKeyFig1">Selected input key figure from cost table</param>
        /// <param name="inKeyFig2">Selected input key figure from margin table</param>
        /// <param name="outKeyFig">Name of resultant (output) key figure </param>
        /// <param name="StdMargin">Standard margin to be applied on other products not included in margin table</param>
        /// <param name="CellStr">Upper-left corner of the resultant table in excel sheet</param>
        /// <param name="ClearSheetContent">If true, clear the whole sheet before inserting the resultant table</param>
        /// <returns>ListObject (excel table) for resultant price table</returns>
        /// <remarks>
        /// Calling this method from excel VBA requires MS Interop (not NetOffice interop)
        /// because it has an excel object parameter (worksheet). 
        /// </remarks>
        public XTable GetPriceTableX(XTable xCostTbl, XTable xMarginTbl, XTable xMetaTbl,
            Excel.Worksheet WSheet, string PriceTblName, 
            string inKeyFig1 = "costs", string inKeyFig2 = "margin", string outKeyFig = "price",
            double StdMargin = 0.25, string CellStr = "A1", bool ClearSheetContent = true)
        {
            // Typical flow of calculation:

            // Step 1: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 2: Read excel tables (inputs) into MatrixTable objects
            MatrixTable CostTbl = ExcelToMatrixTable(xCostTbl, mdx).matrixTable;
            MatrixTable MarginTbl = ExcelToMatrixTable(xMarginTbl, mdx).matrixTable;

            // Step 3: Generate resultant (output) tables with table functions
            MatrixTable PriceTbl = MatrixTable.MultiplySelectedKeyFigures(
                CostTbl, MarginTbl + 1, inKeyFig1, inKeyFig2, outKeyFig,
                MultiplyRestWith: (StdMargin + 1), JokerMatchesAllvalues: true);


            // Step 4: Write output tables (in this case PriceTbl) into excel tables
            return MatrixTableToExcel(new MatrixTableX(PriceTbl), WSheet, PriceTblName);
        }

        /// <summary>
        /// Obtain Price table by multiplying Cost and Margin tables:
        /// PriceTable = CostTable * (MarginTable + 1).
        /// Underlying Table Function: MatrixTable.MultiplySelectedKeyFigures()
        /// </summary>
        /// <param name="xCostTblName">Name of ListObject (excel table) for costs</param>
        /// <param name="xMarginTblName">Name of ListObject (excel table) for margins</param>
        /// <param name="xMetaTblName">Name of ListObject (excel table) for field definitions.</param>
        /// <param name="WorkbookFullName">Full name of excel workbook (if null, active workbook is assumed)</param>
        /// <param name="xPriceTblName">Name of ListObject (excel table) for the resultant (output) prices</param>
        /// <param name="TargetSheetName">Sheet name for writing resultant table</param>
        /// <param name="inKeyFig1">Selected input key figure from cost table</param>
        /// <param name="inKeyFig2">Selected input key figure from margin table</param>
        /// <param name="outKeyFig">Name of resultant (output) key figure </param>
        /// <param name="StdMargin">Standard margin to be applied on other products not included in margin table</param>
        /// <param name="CellStr">Upper-left corner of the resultant table in excel sheet</param>
        /// <param name="ClearSheetContent">If true, clear the whole sheet before inserting the resultant table</param>
        /// <remarks>
        /// Excel Tables are passed by their names (string).
        /// Example without explicit declaration of metadata (field definitions).
        /// Set xMetaTblName to null if there is no metadata table with field definitions.
        /// </remarks>
        public void GetPriceTable(string xCostTblName, string xMarginTblName, string xMetaTblName = null,
            string WorkbookFullName = null, string xPriceTblName = "Price", string TargetSheetName = "PriceTable",
            string inKeyFig1 = "costs", string inKeyFig2 = "margin", string outKeyFig = "price",
            double StdMargin = 0.25, string CellStr = "A1", bool ClearSheetContent = true)
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xCostTbl = ExcelFunc_NO.GetListObject(wbook, xCostTblName);
            XTable xMarginTbl = ExcelFunc_NO.GetListObject(wbook, xMarginTblName);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) 
                xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) 
                mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable CostTbl = ExcelToMatrixTable(xCostTbl, mdx).matrixTable;
            MatrixTable MarginTbl = ExcelToMatrixTable(xMarginTbl, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable PriceTbl = MatrixTable.MultiplySelectedKeyFigures(
                CostTbl, MarginTbl + 1, inKeyFig1, inKeyFig2, outKeyFig,
                MultiplyRestWith: (StdMargin + 1), JokerMatchesAllvalues: true);

            // Step 5: Write output tables (in this case PriceTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);
            MatrixTableToExcel(new MatrixTableX(PriceTbl), wsheet, xPriceTblName);
            wsheet.Activate();
        }

        /// <summary>
        /// Obtain price table from cost and margin tables; shows input box for getting parameter values
        /// </summary>
        public void GetPriceTable_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for GetPriceTable";

                string prompt = "Please select a single-column range with 3 input values for following parameters:\n\n"
                    + "1) Name of excel table (ListObject) for costs\n"
                    + "2) Name of excel table (ListObject) for margins\n"
                    + "3) Name of excel table (ListObject) for resultant price table\n";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 3, out ParameterValues))
                {
                    // assign parameter values
                    string xCostTblName = (string)ParameterValues[0];
                    string xMarginTblName = (string)ParameterValues[1];
                    string xPriceTblName = (string)ParameterValues[2];

                    GetPriceTable(xCostTblName, xMarginTblName, 
                        WorkbookFullName: null,
                        xPriceTblName: xPriceTblName
                        );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("GetPriceTable: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Combine two input tables that have some common attributes (text or numeric).
        /// Underlying Table Function: MatrixTable.CombineTablesFirstMatch()
        /// </summary>
        /// <param name="xTb1Name">Name of 1. excel table (ListObject)</param>
        /// <param name="xTbl2Name">Name of 2. excel table (ListObject)</param>
        /// <param name="xMetaTblName">Name of excel table (ListObject) for field definitions.</param>
        /// <param name="wbook">Excel Workbook object</param>
        /// <param name="xCombinedTblName">Name of resultant combined table</param>
        /// <param name="TargetSheetName">Sheet name for writing resultant table</param>
        /// <param name="JokerMatchesAllvalues">If true, joker value matches all possible attribute values</param>
        /// <param name="TextJoker">Joker (match-all) value for text attributes</param>
        /// <param name="NumJoker">Joker (match-all) value for numeric attributes</param>
        /// <returns>Excel Table (ListObject) object for resultant combined table</returns>
        /// <remarks>
        /// Calling this method from excel VBA requires MS Interop (not NetOffice interop)
        /// because this method has an excel object parameter (workbook). 
        /// </remarks>
        public XTable CombineTablesX(string xTb1Name, string xTbl2Name, string xMetaTblName,
            Excel.Workbook wbook, string xCombinedTblName = "Combined", string TargetSheetName = "CombinedTable", 
            bool JokerMatchesAllvalues = true, string TextJoker = "ALL", int NumJoker = 0)
        {
            // Typical flow of calculation:

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTb1Name);
            XTable xTbl2 = ExcelFunc_NO.GetListObject(wbook, xTbl2Name);
            XTable xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;
            MatrixTable Tbl2 = ExcelToMatrixTable(xTbl2, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable CombinedTbl = MatrixTable.CombineTablesFirstMatch(
                Tbl1, Tbl2, JokerMatchesAllvalues, TextJoker, NumJoker);

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);
            return MatrixTableToExcel(new MatrixTableX(CombinedTbl), wsheet, xCombinedTblName);
        }

        /// <summary>
        /// Combine two input tables that have some common attributes (text or numeric).
        /// Underlying Table Function: MatrixTable.CombineTablesFirstMatch()
        /// </summary>
        /// <param name="xTb1Name">Name of 1. excel table (ListObject)</param>
        /// <param name="xTbl2Name">Name of 2. excel table (ListObject)</param>
        /// <param name="xMetaTblName">Name of excel table (ListObject) for field definitions.</param>
        /// <param name="WorkbookFullName">File path of Excel Workbook</param>
        /// <param name="xCombinedTblName">Name of resultant combined table</param>
        /// <param name="TargetSheetName">Sheet name for writing resultant table</param>
        /// <param name="JokerMatchesAllvalues">If true, joker value matches all possible attribute values</param>
        /// <param name="TextJoker">Joker (match-all) value for text attributes</param>
        /// <param name="NumJoker">Joker (match-all) value for numeric attributes</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        /// <remarks>
        /// Excel Tables are passed by their names (string); no excel object parameters
        /// </remarks>
        public void CombineTables(string xTb1Name, string xTbl2Name, string xMetaTblName = null,
                string WorkbookFullName = null, string xCombinedTblName = "Combined", string TargetSheetName = "CombinedTable",
                bool JokerMatchesAllvalues = true, string TextJoker = "ALL", int NumJoker = 0, string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // TEST 
            // MessageBox.Show("Step 0: Get current application and active workbook");

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // TEST 
            // MessageBox.Show("Step 1: Get ListObjects, workook: " + wbook.FullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTb1Name);
            XTable xTbl2 = ExcelFunc_NO.GetListObject(wbook, xTbl2Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // TEST 
            // MessageBox.Show("Step 2: Get meta data with field definitions");

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // TEST 
            // MessageBox.Show("Step 3: Read excel tables (inputs) into MatrixTable objects");

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;
            MatrixTable Tbl2 = ExcelToMatrixTable(xTbl2, mdx).matrixTable;

            // TEST 
            // MessageBox.Show("Step 4: Generate resultant (output) tables with table functions");

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable CombinedTbl = MatrixTable.CombineTablesFirstMatch(
                Tbl1, Tbl2, JokerMatchesAllvalues, TextJoker, NumJoker);

            //TEST
            // CombinedTbl.ViewTable("Resultant Combined Table");

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(CombinedTbl), wsheet, xCombinedTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Combine two input tables that have some common attributes (text or numeric).
        /// Shows input box for getting parameters.
        /// Underlying Table Function: MatrixTable.CombineTables()
        /// </summary>
        public void CombineTables_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for CombineTables";

                string prompt = "Please select a single-column range with 2 input values for following parameters:\n\n"
                    + "1) Name of 1. excel table (ListObject)\n"
                    + "2) Name of 2. excel table (ListObject)\n";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 2, out ParameterValues))
                {
                    // assign parameter values
                    string xTb1Name = (string)ParameterValues[0];
                    string xTbl2Name = (string)ParameterValues[1];

                    CombineTables(xTb1Name, xTbl2Name);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("CombineTables: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Combine two input tables that have some common attributes (text or numeric).
        /// Shows input box for getting parameters.
        /// Underlying Table Function: MatrixTable.CombineTables()
        /// </summary>
        public void CombineTables_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for CombineTables";
                string FuncTitle = "Combine Tables";
                string FuncDescr = "Combine two input tables that have some common attributes (text or numeric).";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_CombineTables(FormTitle, FuncTitle, FuncDescr, "Combined", "CombinedTable", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTb1Name = ParameterForm1.xTable1_st;
                string xTbl2Name = ParameterForm1.xTable2_st;
                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // combine tables
                CombineTables(xTb1Name, xTbl2Name,
                    xCombinedTblName: xOutTable, TargetSheetName: SheetName, 
                    TopLeftCell: TopLeftCell);

            }
            catch (Exception ex)
            {
                MessageBox.Show("CombineTables: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Add two input tables (table addition) with selected key figures.
        /// Underlying table function: MatrixTable.AddSelectedKeyFigures()
        /// </summary>
        /// <param name="xTbl1Name">Name of 1. input table (ListObject)</param>
        /// <param name="xTbl2Name">Name of 2. input table (ListObject)</param>
        /// <param name="InputKeyfig1">Selected key figure from 1. input table</param>
        /// <param name="InputKeyfig2">Selected key figure from 2. input table</param>
        /// <param name="OutputKeyfig">Output (resultant) key figure</param>
        /// <param name="xMetaTblName">Name of excel table (ListObject) with field definitions</param>
        /// <param name="WorkbookFullName">File path for workbook</param>
        /// <param name="xResultantTblName">Name of resultant output table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void AddTables(string xTbl1Name, string xTbl2Name,
            string InputKeyfig1, string InputKeyfig2, string OutputKeyfig,
             string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "Result", string TargetSheetName = "ResultTable", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);
            XTable xTbl2 = ExcelFunc_NO.GetListObject(wbook, xTbl2Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;
            MatrixTable Tbl2 = ExcelToMatrixTable(xTbl2, mdx).matrixTable;

            // TEST
            // Tbl1.ViewTable(xTbl1Name);
            // Tbl2.ViewTable(xTbl2Name);

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl = MatrixTable.AddSelectedKeyFigures(
                Tbl1, Tbl2, InputKeyfig1, InputKeyfig2, OutputKeyfig,
                AddToRest: 0.0, JokerMatchesAllvalues: true,
                TextJoker: "ALL", NumJoker: 0);

            // TEST
            // ResultantTbl.ViewTable(xResultantTblName);

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Add two input tables (table addition) with selected key figures.
        /// Calls input range box for getting parameter values.
        /// Underlying table function: MatrixTable.AddSelectedKeyFigures()
        /// </summary>
        public void AddTables_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for AddTables";

                string prompt = "Please select a single-column range with 5 input values for following parameters:\n\n"
                    + "1) Name of 1. input table\n"
                    + "2) Name of 2. input table\n"
                    + "3) Selected key figure from 1. input table\n"
                    + "4) Selected key figure from 2. input table\n"
                    + "5) Resultant (output) key figure\n";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 5, out ParameterValues))
                {
                    // assign parameter values
                    string xTb1Name = (string)ParameterValues[0];
                    string xTbl2Name = (string)ParameterValues[1];

                    string InputKeyfig1 = (string)ParameterValues[2];
                    string InputKeyfig2 = (string)ParameterValues[3];
                    string OutputKeyfig = (string)ParameterValues[4];

                    AddTables(xTb1Name, xTbl2Name,
                        InputKeyfig1, InputKeyfig2, OutputKeyfig);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("AddTables: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Add two input tables (table addition) with selected key figures.
        /// Calls input range box for getting parameter values.
        /// Underlying table function: MatrixTable.AddSelectedKeyFigures()
        /// </summary>
        public void AddTables_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for AddTables";
                string FuncTitle = "Add Tables";
                string FuncDescr = "Table Addition: Add two input tables with selected key figures.\n\n"
                    + "All the attributes of 2. input table (numeric & text) must be contained in 1. input table.";

		        using (ParameterForm1 myform = new ParameterForm1())
		        {
                    myform.PrepForm_TableArithmetics(FormTitle, FuncTitle, FuncDescr, "Result", "ResultTable", "A1");
			        myform.ShowDialog();
		        }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTb1Name = ParameterForm1.xTable1_st;
                string xTbl2Name = ParameterForm1.xTable2_st;
                string InputKeyfig1 = ParameterForm1.InKeyFig1_st;
                string InputKeyfig2 = ParameterForm1.InKeyFig2_st;
                string xOutTable = ParameterForm1.xOutTable1_st;
                string OutputKeyfig = ParameterForm1.OutKeyFig_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // add tables
                AddTables(xTb1Name, xTbl2Name, InputKeyfig1, InputKeyfig2, OutputKeyfig,
                xResultantTblName: xOutTable, TargetSheetName: SheetName, TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("AddTables: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Subtract 2. input table from 1. input table with selected key figures.
        /// Underlying table function: MatrixTable.SubtractSelectedKeyFigures()
        /// </summary>
        /// <param name="xTbl1Name">Name of 1. input table (ListObject)</param>
        /// <param name="xTbl2Name">Name of 2. input table (ListObject)</param>
        /// <param name="InputKeyfig1">Selected key figure from 1. input table</param>
        /// <param name="InputKeyfig2">Selected key figure from 2. input table</param>
        /// <param name="OutputKeyfig">Output (resultant) key figure</param>
        /// <param name="xMetaTblName">Name of excel table (ListObject) with field definitions</param>
        /// <param name="WorkbookFullName">File path for workbook</param>
        /// <param name="xResultantTblName">Name of resultant output table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void SubtractTable(string xTbl1Name, string xTbl2Name,
            string InputKeyfig1, string InputKeyfig2, string OutputKeyfig, 
             string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "Result", string TargetSheetName = "ResultTable", string TopLeftCell = "A1")
        {

            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);
            XTable xTbl2 = ExcelFunc_NO.GetListObject(wbook, xTbl2Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;
            MatrixTable Tbl2 = ExcelToMatrixTable(xTbl2, mdx).matrixTable;

            // TEST
            // Tbl1.ViewTable(xTbl1Name);
            // Tbl2.ViewTable(xTbl2Name);

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl = MatrixTable.SubtractSelectedKeyFigures(
                Tbl1, Tbl2, InputKeyfig1, InputKeyfig2, OutputKeyfig,
                SubtractFromRest: 0.0, JokerMatchesAllvalues: true,
                TextJoker: "ALL", NumJoker: 0);

            // TEST
            // ResultantTbl.ViewTable(xResultantTblName);

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Divide 1. input table by 2. input table with selected key figures.
        /// Underlying table function: MatrixTable.DivideSelectedKeyFigures()
        /// </summary>
        /// <param name="xTbl1Name">Name of 1. input table (ListObject)</param>
        /// <param name="xTbl2Name">Name of 2. input table (ListObject)</param>
        /// <param name="InputKeyfig1">Selected key figure from 1. input table</param>
        /// <param name="InputKeyfig2">Selected key figure from 2. input table</param>
        /// <param name="OutputKeyfig">Output (resultant) key figure</param>
        /// <param name="xMetaTblName">Name of excel table (ListObject) with field definitions</param>
        /// <param name="WorkbookFullName">File path for workbook</param>
        /// <param name="xResultantTblName">Name of resultant output table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void DivideTable(string xTbl1Name, string xTbl2Name, 
            string InputKeyfig1, string InputKeyfig2, string OutputKeyfig,
             string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "Result", string TargetSheetName = "ResultTable", string TopLeftCell = "A1")
        {

            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);
            XTable xTbl2 = ExcelFunc_NO.GetListObject(wbook, xTbl2Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;
            MatrixTable Tbl2 = ExcelToMatrixTable(xTbl2, mdx).matrixTable;

            // TEST
            // Tbl1.ViewTable(xTbl1Name);
            // Tbl2.ViewTable(xTbl2Name);

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl = MatrixTable.DivideSelectedKeyFigures(
                Tbl1, Tbl2, InputKeyfig1, InputKeyfig2, OutputKeyfig,
                DivideRestBy: 1.0, JokerMatchesAllvalues: true,
                TextJoker: "ALL", NumJoker: 0);

            // TEST
            // ResultantTbl.ViewTable(xResultantTblName);

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Table Arithmetics: Addition, Multiplication, Subtraction or Division
        /// </summary>
        public void TableArithmetics_macro() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for TableArithmetics";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_TableArithmeticsAll(FormTitle, "Result", "ResultTable", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTb1Name = ParameterForm1.xTable1_st;
                string xTbl2Name = ParameterForm1.xTable2_st;
                string InputKeyfig1 = ParameterForm1.InKeyFig1_st;
                string InputKeyfig2 = ParameterForm1.InKeyFig2_st;
                string xOutTable = ParameterForm1.xOutTable1_st;
                string OutputKeyfig = ParameterForm1.OutKeyFig_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;
                string ArithmeticOperation = ParameterForm1.comboBox1_st;

                if (ArithmeticOperation == ParameterForm1.TableArithmetics_FuncTitles[0])
                {
                    AddTables(xTb1Name, xTbl2Name, InputKeyfig1, InputKeyfig2, OutputKeyfig,
                        xResultantTblName: xOutTable, TargetSheetName: SheetName, TopLeftCell: TopLeftCell);
                }
                else if (ArithmeticOperation == ParameterForm1.TableArithmetics_FuncTitles[1])
                {
                    MultiplyTables(xTb1Name, xTbl2Name, InputKeyfig1, InputKeyfig2, OutputKeyfig,
                        xResultantTblName: xOutTable, TargetSheetName: SheetName, TopLeftCell: TopLeftCell);
                }
                else if (ArithmeticOperation == ParameterForm1.TableArithmetics_FuncTitles[2])
                {
                    SubtractTable(xTb1Name, xTbl2Name, InputKeyfig1, InputKeyfig2, OutputKeyfig,
                        xResultantTblName: xOutTable, TargetSheetName: SheetName, TopLeftCell: TopLeftCell);
                }
                else if (ArithmeticOperation == ParameterForm1.TableArithmetics_FuncTitles[3])
                {
                    DivideTable(xTb1Name, xTbl2Name, InputKeyfig1, InputKeyfig2, OutputKeyfig,
                        xResultantTblName: xOutTable, TargetSheetName: SheetName, TopLeftCell: TopLeftCell);
                }
                else
                {
                    // normally, should not happen
                    throw new Exception("Undefined arithmetic operation type! \n"
                        + "ArithmeticOperation = " + ArithmeticOperation + "\n ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("TableArithmetics: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Multiply two input tables (table multiplication) with selected key figures.
        /// Underlying table function: MatrixTable.MultiplySelectedKeyFigures()
        /// </summary>
        /// <param name="xTbl1Name">Name of 1. input table (ListObject)</param>
        /// <param name="xTbl2Name">Name of 2. input table (ListObject)</param>
        /// <param name="InputKeyfig1">Selected key figure from 1. input table</param>
        /// <param name="InputKeyfig2">Selected key figure from 2. input table</param>
        /// <param name="OutputKeyfig">Output (resultant) key figure</param>
        /// <param name="xMetaTblName">Name of excel table (ListObject) with field definitions</param>
        /// <param name="WorkbookFullName">File path for workbook</param>
        /// <param name="xResultantTblName">Name of resultant output table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void MultiplyTables(string xTbl1Name, string xTbl2Name,
            string InputKeyfig1, string InputKeyfig2, string OutputKeyfig,
             string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "Result", string TargetSheetName = "ResultTable", string TopLeftCell = "A1")
        {

            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);
            XTable xTbl2 = ExcelFunc_NO.GetListObject(wbook, xTbl2Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;
            MatrixTable Tbl2 = ExcelToMatrixTable(xTbl2, mdx).matrixTable;

            // TEST
            // Tbl1.ViewTable(xTbl1Name);
            // Tbl2.ViewTable(xTbl2Name);

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl = MatrixTable.MultiplySelectedKeyFigures(
                Tbl1, Tbl2, InputKeyfig1, InputKeyfig2, OutputKeyfig,
                MultiplyRestWith: 1.0, JokerMatchesAllvalues: true,
                TextJoker: "ALL", NumJoker: 0);

            // TEST
            // ResultantTbl.ViewTable(xResultantTblName);

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Multiply two input tables (table multiplication) with selected key figures.
        /// Calls input range box for getting parameter values.
        /// Underlying table function: MatrixTable.MultiplySelectedKeyFigures()
        /// </summary>
        public void MultiplyTables_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for MultiplyTables";

                string prompt = "Please select a single-column range with 5 input values for following parameters:\n\n"
                    + "1) Name of 1. input table\n"
                    + "2) Name of 2. input table\n"
                    + "3) Selected key figure from 1. input table\n"
                    + "4) Selected key figure from 2. input table\n"
                    + "5) Resultant (output) key figure\n";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 5, out ParameterValues))
                {
                    // assign parameter values
                    string xTb1Name = (string)ParameterValues[0];
                    string xTbl2Name = (string)ParameterValues[1];

                    string InputKeyfig1 = (string)ParameterValues[2];
                    string InputKeyfig2 = (string)ParameterValues[3];
                    string OutputKeyfig = (string)ParameterValues[4];

                    MultiplyTables(xTb1Name, xTbl2Name,
                        InputKeyfig1, InputKeyfig2, OutputKeyfig);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("MultiplyTables: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Simple pro-rate distribution; returns table with distributed amounts.
        /// The source amount in SourceTable is disrubuted to target amounts w.r.t. keys (or ratios) given in KeyTable.
        /// Important: Both Source and Key tables must have exactly one key figure.
        /// </summary>
        /// <param name="xSourceTblName">Name of source table with source amounts to be distributed</param>
        /// <param name="xDistrKeyTblName">Name if key table with key amounts (or ratios) used for the distribution</param>
        /// <param name="ResultantKeyFig">Name of key figure with the resultant distributed amounts</param>
        /// <param name="KeySumKeyFig">Name of the key figure with the sum of key amounts</param>
        /// <param name="xMetaTblName">Name of metadata table with field definitions</param>
        /// <param name="WorkbookFullName">File path of workbook</param>
        /// <param name="xResultantTblName">Name of resultant combined table with distributed amounts</param>
        /// <param name="TargetSheetName">Target sheet for resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void DistributeTable(string xSourceTblName, string xDistrKeyTblName,
            string ResultantKeyFig, string KeySumKeyFig,
            string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "Distrib", string TargetSheetName = "DistribTable", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xSourceTblName);
            XTable xTbl2 = ExcelFunc_NO.GetListObject(wbook, xDistrKeyTblName);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // TEST
            // MessageBox.Show("Step 3: Read excel tables (inputs) into MatrixTable objects");

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;
            MatrixTable Tbl2 = ExcelToMatrixTable(xTbl2, mdx).matrixTable;

            // TEST
            // MessageBox.Show("Step 4: Generate resultant (output) tables with table functions");

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl = AppUtility.SimpleDistribution(Tbl1, Tbl2,
                ResultantKeyFig, KeySumKeyFig);

            // TEST
            // MessageBox.Show("Step 5: Write output tables (in this case ResultantTbl) into excel tables");
            // ResultantTbl.ViewTable(xResultantTblName);

            // Step 5: Write output tables (in this case ResultantTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Simple pro-rate distribution; returns table with distributed amounts.
        /// </summary>
        public void DistributeTable_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for DistributeTable";

                string prompt = "Please select a single-column range with 4 input values for following parameters:\n\n"
                    + "1) Name of source table\n"
                    + "2) Name of key table with distribution ratios\n"
                    + "3) Name of resultant key figure\n"
                    + "4) Name of key figure for sum of key values\n";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 4, out ParameterValues))
                {
                    // assign parameter values
                    string xTb1Name = (string)ParameterValues[0];
                    string xTbl2Name = (string)ParameterValues[1];

                    string ResultantKeyFig = (string)ParameterValues[2];
                    string KeySumKeyFig = (string)ParameterValues[3];

                    DistributeTable(xTb1Name, xTbl2Name,
                        ResultantKeyFig, KeySumKeyFig);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("DistributeTable: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Simple pro-rate distribution; returns table with distributed amounts.
        /// </summary>
        public void DistributeTable_macro2()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for DistributeTable";
                string FuncTitle = "Distribute Table";
                string FuncDescr = "Simple pro-rate distribution; returns table with distributed amounts.\n\n"
                    + "The source amount in SourceTable is disrubuted to target amounts w.r.t. keys (or ratios) given in KeyTable.\n"
                    + "Important: Both Source and Key tables must have exactly one key figure.";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_DistributeTable(FormTitle, FuncTitle, FuncDescr,
                        "distributed_amount", "sum_of_keys", "Result", "ResultTable", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTb1Name = ParameterForm1.xTable1_st;
                string xTbl2Name = ParameterForm1.xTable2_st;
                string ResultKeyFig = ParameterForm1.txtBox1_st;
                string KeySumsKeyFig = ParameterForm1.txtBox2_st; 

                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // call table function
                DistributeTable(xTb1Name, xTbl2Name, ResultKeyFig, KeySumsKeyFig, null, null,
                    xOutTable, SheetName, TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("AddTables: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Create a new table with all possible combinations of given field values in input table.
        /// If a value set is not explicity given for a field, single default value is assumed as value set.
        /// </summary>
        /// <param name="xTblName">Name of excel table (ListObject) with field values</param>
        /// <param name="xMetaTblName">Name of metadata table with field definitions</param>
        /// <param name="WorkbookFullName">File path of workbook</param>
        /// <param name="DefaultTextVal">Default text attribute value</param>
        /// <param name="DefaultNumVal">Default numeric attribute value</param>
        /// <param name="DefaultKeyFigVal">Default key figure value</param>
        /// <param name="xResultantTblName">Name of resultant excel table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for inserting resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void CombinateFieldValues(string xTblName, 
            string xMetaTblName = null, string WorkbookFullName = null, 
            string DefaultTextVal = "", int DefaultNumVal = 0, double DefaultKeyFigVal = 0,
            string xResultantTblName = "FieldComb", string TargetSheetName = "FieldCombTable", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTblName);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;

            // TEST
            // Tbl1.ViewTable("Table with field values; now get combination table");

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl = MatrixTable.CombinateFieldValues_A(Tbl1,
                DefaultTextVal, DefaultNumVal, DefaultKeyFigVal);

            // TEST
            // ResultantTbl.ViewTable("Combination table; now write it into excel");

            // Step 5: Write output tables (in this case ResultantTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);
            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Create a new table with all possible combinations of given field values in input table.
        /// </summary>
        public void CombinateFieldValues_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for CombinateFieldValues";

                string prompt = "Please select a single-column range with 1 input values for following parameters:\n\n"
                    + "1) Name of input table with field values\n";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 1, out ParameterValues))
                {
                    // assign parameter values
                    string xTb1Name = (string)ParameterValues[0];

                    CombinateFieldValues(xTb1Name);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("CombinateFieldValues: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Create a new table with all possible combinations of given field values in input table.
        /// </summary>
        public void CombinateFieldValues_macro2()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for CombinateFieldValues";
                string FuncTitle = "Combinate Field Values";
                string FuncDescr = "reate a new table with all possible combinations of given field values in input table.";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_CombinateFields(FormTitle, FuncTitle, FuncDescr,
                        "FieldComb", "FieldCombTable", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTb1Name = ParameterForm1.xTable1_st;
                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // call table function
                CombinateFieldValues(xTb1Name, null, null,
                    xResultantTblName: xOutTable, TargetSheetName: SheetName, 
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("CombinateFieldValues: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Filter input table BaseTbl with a condition table CondTbl.
        /// if ExcludeMatchedRows = false (default case) output table includes only rows that match with
        /// some row(s) of condition table CondTbl.
        /// </summary>
        /// <param name="xBaseTblName">Input base table to be filtered</param>
        /// <param name="xCondTblName">Condition table defining filtering criteria</param>
        /// <param name="ExcludeMatchedRows">If true, exclude matched rows from input table</param>
        /// <param name="JokerMatchesAllvalues">If true, JokerValue matches all other values</param>
        /// <param name="TextJoker">Joker (match-all) value for text attributes</param>
        /// <param name="NumJoker">Joker (match-all) value for numeric attributes</param>
        /// <param name="xMetaTblName">Name of metadata table with field definitions</param>
        /// <param name="WorkbookFullName">File path of workbook</param>
        /// <param name="xResultantTblName">Name of resultant excel table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for inserting resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void FilterTable(string xBaseTblName, string xCondTblName,
            bool ExcludeMatchedRows = false, bool JokerMatchesAllvalues = true,
            string TextJoker = "ALL", int NumJoker = 0,
            string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "Filtered", string TargetSheetName = "FilteredTable",
            string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xBaseTblName);
            XTable xTbl2 = ExcelFunc_NO.GetListObject(wbook, xCondTblName);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;
            MatrixTable Tbl2 = ExcelToMatrixTable(xTbl2, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl = MatrixTable.FilterTableA(Tbl1, Tbl2,
                ExcludeMatchedRows, JokerMatchesAllvalues, TextJoker, NumJoker);

            // Step 5: Write output tables (in this case ResultantTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName);
            wsheet.Activate();
        }

        /// <summary>
        /// Filter input table BaseTbl with a condition table CondTbl.
        /// if ExcludeMatchedRows = false (default case) output table includes only rows that match with
        /// some row(s) of condition table CondTbl.
        /// </summary>
        public void FilterTable_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for FilterTable";

                string prompt = "Please select a single-column range with 2 input values for following parameters:\n\n"
                    + "1) Name of base table to be filtered\n"
                    + "2) Name of condition table defining filtering criteria\n";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 2, out ParameterValues))
                {
                    // assign parameter values
                    string xTb1Name = (string)ParameterValues[0];
                    string xTb2Name = (string)ParameterValues[1];

                    FilterTable(xTb1Name, xTb2Name);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("FilterTable: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Filter input table BaseTbl with a condition table CondTbl.
        /// if ExcludeMatchedRows = false (default case) output table includes only rows that match with
        /// some row(s) of condition table CondTbl.
        /// </summary>
        public void FilterTable_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for FilterTable";
                string FuncTitle = "Filter Table";
                string FuncDescr = "Filter 1. input table (Base Table) with 2. input table (Condition Table).";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_FilterTable(FormTitle, FuncTitle, FuncDescr, "Filtered", "FilteredTable", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTbl1Name = ParameterForm1.xTable1_st;
                string xTbl2Name = ParameterForm1.xTable2_st;
                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;
                bool ExcludeMatchedRows = ParameterForm1.checkBox1_st; 

                // add tables
                FilterTable(xTbl1Name, xTbl2Name,
                    ExcludeMatchedRows: ExcludeMatchedRows,
                    xResultantTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("AddTables: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Cartesian Multiplication of table rows: Generate a new table 
        /// as all possible row combinations of input tables. 
        /// There must be no common fields among input tables.
        /// </summary>
        /// <param name="xMetaTblName">Name of metadata table with field definitions</param>
        /// <param name="WorkbookFullName">File path of workbook</param>
        /// <param name="xResultantTblName">Name of resultant excel table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for inserting resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        /// <param name="xTableNames">Names of excel tables</param>
        public void CombinateRows(string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "CombRows", string TargetSheetName = "CombRowsTable",
            string TopLeftCell = "A1", params string[] xTableNames)
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable[] xTblArr = new XTable[xTableNames.Count()];

            string xTableName;

            for (int i = 0; i < xTableNames.Count(); i++)
            {
                xTableName = xTableNames[i];
                xTblArr[i] = ExcelFunc_NO.GetListObject(wbook, xTableName);
            }

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable[] TblArr = new MatrixTable[xTableNames.Count()];

            for (int i = 0; i < xTblArr.Count(); i++)
            {
                TblArr[i] = ExcelToMatrixTable(xTblArr[i], mdx).matrixTable;
            }

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl = MatrixTable.CombinateTableRows(TblArr);

            // Step 5: Write output tables (in this case ResultantTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Cartesian Multiplication of table rows: Generate a new table 
        /// as all possible row combinations of input tables. 
        /// There must be no common fields among input tables.
        /// </summary>
        public void CombinateRows_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for CombinateRows";

                string prompt = "Please select a single-column range with the names of 2 or more excel tables (ListObject) whose rows will be combinated:\n\n";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 2, out ParameterValues, false))
                {
                    var xTableNames = new string[ParameterValues.Count];

                    // assign parameter values
                    for (int i = 0; i < ParameterValues.Count; i++)
                    {
                        xTableNames[i] = (string)ParameterValues[i];
                    }

                    CombinateRows(xTableNames: xTableNames);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("CombinateRows: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Cartesian Multiplication of table rows: Generate a new table 
        /// as all possible row combinations of input tables. 
        /// There must be no common fields among input tables.
        /// </summary>
        public void CombinateRows_macro2()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for CombinateRows";
                string FuncTitle = "Combinate Rows of Tables";
                string FuncDescr = "Cartesian Multiplication of Table Rows: Generate a new table "
                    + "with all possible row combinations of input tables.\n\n"
                    + "Important: There must be no common fields among input tables.";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_CombinateRows(FormTitle, FuncTitle, FuncDescr, "RowComb", "RowCombTable", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string[] SelectedTables = ParameterForm1.listBox1_st;
                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // call table function
                CombinateRows(xTableNames: SelectedTables, xResultantTblName: xOutTable,
                    TargetSheetName: SheetName, TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("CombinateRows: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Aggregate all selected key figures of table with the selected (default) aggregation function. 
        /// Resultant aggregated table contains selected attributes and key figures.
        /// </summary>
        /// <param name="xTbl1Name">Input table to be aggregated over given attributes</param>
        /// <param name="IncludedAttributes">Included attributes in the resultant aggregated table</param>
        /// <param name="IncludedKeyFigures">Included key figures in the resultant aggregated table</param>
        /// <param name="AggregateFunc">Aggregation function for all selected key figures</param>
        /// <param name="xMetaTblName">Name of metadata table with field definitions</param>
        /// <param name="WorkbookFullName">File path of workbook</param>
        /// <param name="xResultantTblName">Name of resultant excel table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for inserting resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void AggregateTable(string xTbl1Name, string[] IncludedAttributes,
            string[] IncludedKeyFigures, string AggregateFunc = "sum", 
             string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "AggregatedTbl", string TargetSheetName = "AggregatedTbl", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;

            // TEST
            // Tbl1.ViewTable(xTbl1Name);
            // Tbl2.ViewTable(xTbl2Name);

            // Step 4: Generate resultant (output) tables with table functions

            // get excluded attributes
            if (IncludedAttributes == null || IncludedAttributes.Count() < 1)
                throw new Exception("Null-valued or empty string vector IncludedAttributes!");

            TextVector InclAttr = new TextVector(IncludedAttributes);
            TextVector ExclTextAttr = TextVector.SetDifference(Tbl1.TextAttributeFields, InclAttr);
            TextVector ExclNumAttr = TextVector.SetDifference(Tbl1.NumAttributeFields, InclAttr);
            TextVector InclKeyFig = new TextVector(IncludedKeyFigures);

            // get aggregation option from string
            AggregateOption AggrOp = AggregateOption.nSum;
            string AggrOpStr = AggregateFunc.Trim().ToLower();

            switch (AggrOpStr)
            {
                case "sum":
                    AggrOp = AggregateOption.nSum;
                    break;

                case "avg":
                    AggrOp = AggregateOption.nAvg;
                    break;

                case "min":
                    AggrOp = AggregateOption.nMin;
                    break;

                case "max":
                    AggrOp = AggregateOption.nMax;
                    break;

                default:        // normally, should not happen
                    throw new Exception("Undefined aggregation function!");
            }

            MatrixTable ResultantTbl = MatrixTable.AggregateTable(
                Tbl1, ExclTextAttr, ExclNumAttr, InclKeyFig, null, AggrOp);

            // TEST
            // ResultantTbl.ViewTable(xResultantTblName);

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Aggregate all selected key figures of table with the selected (default) aggregation function. 
        /// Resultant aggregated table contains selected attributes and key figures.
        /// </summary>
        public void AggregateTable_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for AggregateTable";
                string FuncTitle = "Aggregate Table";
                string FuncDescr = "Aggregate all selected key figures of input table with the selected aggregation function. "
                    + "Resultant aggregated table contains selected attributes and key figures.\n\n"
                    + "NOTE: In this particular aggregation function the same aggregation function (sum, avg, min or max) is applied on all selected key figures. "
                    + "See table functions of the underlying .NET library Finaquant Calcs if you want to apply different aggregation functions for each key figure of an input table.";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_AggregateTable(FormTitle, FuncTitle, FuncDescr, "AggregatedTbl", "AggregatedTbl", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTbl1Name = ParameterForm1.xTable1_st;
                string AggrFunc = ParameterForm1.comboBox1_st;
                string[] InclAttributes = ParameterForm1.listBox1_st;
                string[] InclKeyfigs = ParameterForm1.listBox2_st;

                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;
                bool ExcludeMatchedRows = ParameterForm1.checkBox1_st;

                // aggregate table
                AggregateTable(xTbl1Name, InclAttributes, InclKeyfigs, AggrFunc,
                    xResultantTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("AggregateTable: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Row Transformer: Apply a user-defined transformation function on every row of input table.
        /// </summary>
        /// <param name="xTbl1Name">Input table</param>
        /// <param name="UserCode">A valid piece of code in C# as user-defined function</param>
        /// <param name="xMetaTblName">Name of metadata table with field definitions</param>
        /// <param name="WorkbookFullName">File path of workbook</param>
        /// <param name="xResultantTblName">Name of resultant excel table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for inserting resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void ApplyUserDefinedTransformFuncOnRows(string xTbl1Name, string UserCode,
            string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "ResultTbl", string TargetSheetName = "ResultTable",
            string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            HelperFunc.UserDefinedTransformFunction(UserCode);
            MatrixTable ResultantTbl = MatrixTable.TransformRowsDic(Tbl1, HelperFunc.ApplyUserDefinedTransformFuncOnTableRow);

            // Step 5: Write output tables (in this case ResultantTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Row Transformer: Apply a user-defined transformation function on every row of input table.
        /// </summary>
        public void ApplyUserDefinedTransformFuncOnRows_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for ApplyUserDefinedTransformFuncOnRows";
                string FuncTitle = "Apply User-Defined Transformation Function on Table Rows";
                string FuncDescr = "Apply user-defined function (valid C# code) entered here on every row of input table (Row Transformer). "
                    + @"User-defined function can contain anything including if statements and other structures, provided that it is a valid C# code. " 
                    + @"See code examples below assuming that input table contains text attributes ""category"" and ""product"", numeric attributes ""year"" and ""date"", key figures "
                    + @" ""costs"", ""margin"" and ""price"":" + "\n\n"
                    + @"1) KF[""price""] = (1 + KF[""margin""]) * KF[""price""];" + "\n\n"
                    + @"2) if (TA[""category""] == ""Computers"" && NA[""year""] == 2010)" + "\n" + @"KF[""price""] = 1.3 * KF[""price""];" + "\n" + @"else KF[""price""] = (1 + KF[""margin""]) * KF[""price""];" + "\n\n"
                    + @"3) if (TA[""product""] == ""HPX Laptop X7"")" + "\n" + @"TA[""category""] = ""Laptop"";";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_UserDefinedTransformFuncOnRows(FormTitle, FuncTitle, FuncDescr, "ResultTbl", "ResultTbl", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTbl1Name = ParameterForm1.xTable1_st;
                string UserCode = ParameterForm1.RtxtBox2_st;
                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // aggregate table
                ApplyUserDefinedTransformFuncOnRows(xTbl1Name, UserCode,
                    xResultantTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ApplyUserDefinedTransformFuncOnRows: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Return subtable with selected columns of input table.
        /// </summary>
        /// <param name="xTbl1Name">Input table</param>
        /// <param name="SelectedFields">Selected fields of input table</param>
        /// <param name="IfExcludeFields">If true, exclude selected fields</param>
        /// <param name="xMetaTblName">Name of metadata table with field definitions</param>
        /// <param name="WorkbookFullName">File path of workbook</param>
        /// <param name="xResultantTblName">Name of resultant excel table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for inserting resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void SelectColumns(string xTbl1Name, string[] SelectedFields,
            bool IfExcludeFields = false,
            string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "Subtable", string TargetSheetName = "Subtable", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl;

            if (IfExcludeFields)
                ResultantTbl = MatrixTable.ExcludeColumns(Tbl1, new TextVector(SelectedFields));
            else
                ResultantTbl = MatrixTable.SelectColumns(Tbl1, new TextVector(SelectedFields));

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Return subtable with selected columns of input table.
        /// </summary>
        public void SelectColumns_macro2()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for SelectColumns";
                string FuncTitle = "Select Columns of Input Table";
                string FuncDescr = "Return a subtable with selected fields of input table. Selected fields are excluded if the checkbox below is checked; otherwise, selected fields are included.";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_SelectColumns(FormTitle, FuncTitle, FuncDescr, "Subtable", "Subtable", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTbl1Name = ParameterForm1.xTable1_st;
                string[] SelectedFields = ParameterForm1.listBox1_st;
                bool IfExclude = ParameterForm1.checkBox1_st;
                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // aggregate table
                SelectColumns(xTbl1Name, SelectedFields, IfExclude,
                    xResultantTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("SelectColumns: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Apply selected scalar arithmetical operation (Addition, Multiplication, Subtraction, Division)
        /// on selected key figures of input table, like adding a scalar value to selected key figures.
        /// </summary>
        /// <param name="xTbl1Name">Input table</param>
        /// <param name="Operation">Addition, Multiplication, Subtraction or Division</param>
        /// <param name="SelectedKeyFigures">Array of selected key figures</param>
        /// <param name="ScalarValue">Scalar key figure value like 2.5</param>
        /// <param name="xMetaTblName">Name of metadata table with field definitions</param>
        /// <param name="WorkbookFullName">File path of workbook</param>
        /// <param name="xResultantTblName">Name of resultant excel table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for inserting resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void ScalarOperation(string xTbl1Name, string Operation, 
            string[] SelectedKeyFigures, double ScalarValue,
            string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "ResultTbl", string TargetSheetName = "ResultTbl", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl = Tbl1;

            foreach (var keyfig in SelectedKeyFigures)
            {

                if (Operation == "Addition")
                    ResultantTbl = MatrixTable.AddScalarToSelectedKeyFigure(ResultantTbl, ScalarValue, keyfig, keyfig);
                else if (Operation == "Multiplication")
                    ResultantTbl = MatrixTable.MultiplySelectedKeyFigureByScalar(ResultantTbl, ScalarValue, keyfig, keyfig);
                else if (Operation == "Subtraction")
                    ResultantTbl = MatrixTable.SubtractScalarFromSelectedKeyFigure(ResultantTbl, ScalarValue, keyfig, keyfig);
                else if (Operation == "Division")
                    ResultantTbl = MatrixTable.DivideSelectedKeyFigureByScalar(ResultantTbl, ScalarValue, keyfig, keyfig);
                else
                    throw new Exception("Invalid operation name " + Operation);
            }

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Apply selected scalar arithmetical operation (Addition, Multiplication, Subtraction, Division)
        /// on selected key figures of input table, like adding a scalar value to selected key figures.
        /// </summary>
        public void ScalarOperation_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for ScalarOperation";
                string FuncTitle = "Scalar Arithmetic Operation";
                string FuncDescr = "Add, Multiply, Subtract or Divide selected key figures of input table by a scalar number like 2.5";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_ScalarOperation(FormTitle, FuncTitle, FuncDescr, "ResultTbl", "ResultTbl", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTbl1Name = ParameterForm1.xTable1_st;
                string[] SelectedKeyFigures = ParameterForm1.listBox1_st;
                double ScalarNumber = double.Parse(ParameterForm1.txtBox1_st);
                string Operation = ParameterForm1.comboBox1_st;

                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // aggregate table
                ScalarOperation(xTbl1Name, Operation, SelectedKeyFigures, ScalarNumber,
                    xResultantTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ScalarOperation: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Round selected key figures to given number of digits after decimal point.
        /// </summary>
        /// <param name="xTbl1Name">Input table</param>
        /// <param name="SelectedKeyFigures">Array of selected key figures</param>
        /// <param name="RoundDigits">Number of digits after decimal point</param>
        /// <param name="xMetaTblName">Name of metadata table with field definitions</param>
        /// <param name="WorkbookFullName">File path of workbook</param>
        /// <param name="xResultantTblName">Name of resultant excel table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for inserting resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void RoundNumbers(string xTbl1Name,
            string[] SelectedKeyFigures, int RoundDigits,
            string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "RoundedTbl", string TargetSheetName = "RoundedTbl", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl = Tbl1;

            foreach (var keyfig in SelectedKeyFigures)
            {
                ResultantTbl = MatrixTable.Round(ResultantTbl, RoundDigits, keyfig);
            }

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Round selected key figures to given number of digits after decimal point.
        /// </summary>
        public void RoundNumbers_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for RoundNumbers";
                string FuncTitle = "Round Selected Key Figures";
                string FuncDescr = "Round selected key figures of input table to given number of digits after decimal point.";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_RoundNumbers(FormTitle, FuncTitle, FuncDescr, "RoundTbl", "RoundTbl", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTbl1Name = ParameterForm1.xTable1_st;
                string[] SelectedKeyFigures = ParameterForm1.listBox1_st;
                int RoundDigits = int.Parse(ParameterForm1.comboBox1_st);

                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // aggregate table
                RoundNumbers(xTbl1Name, SelectedKeyFigures, RoundDigits,
                    xResultantTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("RoundNumbers: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Insert a new field with a constant value into input table
        /// </summary>
        /// <param name="xTbl1Name">Input table</param>
        /// <param name="FieldName">Field name</param>
        /// <param name="ftype">Field type</param>
        /// <param name="FieldValue">Constant field value for all rows of table</param>
        /// <param name="xMetaTblName">Name of metadata table with field definitions</param>
        /// <param name="WorkbookFullName">File path of workbook</param>
        /// <param name="xResultantTblName">Name of resultant excel table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for inserting resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void InsertNewField(string xTbl1Name, string FieldName, FieldType ftype, string FieldValue,
            string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "UpdatedTbl", string TargetSheetName = "UpdatedTbl", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            md.AddFieldIfNew(FieldName, ftype);
            MatrixTable ResultantTbl = null;

            switch (ftype)
            {
                case FieldType.DateAttribute:
                    DateTime dt = DateTime.Parse(FieldValue);
                    ResultantTbl = MatrixTable.InsertNewColumn(Tbl1, FieldName, DateFunctions.DateToNumber(dt));
                    break;

                case FieldType.IntegerAttribute:
                    int n = int.Parse(FieldValue);
                    ResultantTbl = MatrixTable.InsertNewColumn(Tbl1, FieldName, n);
                    break;

                case FieldType.KeyFigure:
                    double x = double.Parse(FieldValue);
                    ResultantTbl = MatrixTable.InsertNewColumn(Tbl1, FieldName, x);
                    break;

                case FieldType.TextAttribute:
                    ResultantTbl = MatrixTable.InsertNewColumn(Tbl1, FieldName, FieldValue);
                    break;

                default:
                    break;
            }

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Insert a new field with a constant value into input table
        /// </summary>
        public void InsertNewField_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for InsertNewField";
                string FuncTitle = "Insert A New Field into Table";
                string FuncDescr = "Insert a new field into input table. The entered constant field value must be a valid value for the selected field type.\n"
                    + " Examples for valid values: Red for Text Attribute, 15.10.2012 for Date Attribute, 2012 for Integer Attribute, and 5.25 for Key Figure.";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_InsertNewField(FormTitle, FuncTitle, FuncDescr, "UpdatedTbl", "UpdatedTbl", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTbl1Name = ParameterForm1.xTable1_st;
                string FieldTypeStr = ParameterForm1.comboBox1_st;

                FieldType ftype;

                if (FieldTypeStr == "Text Attribute")
                    ftype = FieldType.TextAttribute;

                else if (FieldTypeStr == "Date Attribute")
                    ftype = FieldType.DateAttribute;

                else if (FieldTypeStr == "Integer Attribute")
                    ftype = FieldType.IntegerAttribute;

                else if (FieldTypeStr == "Key Figure")
                    ftype = FieldType.KeyFigure;

                else
                    ftype = FieldType.Undefined;

                string FieldName = ParameterForm1.txtBox1_st;
                string FieldValue = ParameterForm1.txtBox2_st;

                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // aggregate table
                InsertNewField(xTbl1Name, FieldName, ftype, FieldValue,
                    xResultantTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("InsertNewField: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Assign uniformly distributed random values to selected key figures or numeric attributes (of date or integer type) of input table,
        /// within lower and upper limits.
        /// </summary>
        /// <param name="xTbl1Name">Input table</param>
        /// <param name="SelectedFields">Selected fields</param>
        /// <param name="LowerLimit">Lower limit to random numbers</param>
        /// <param name="UpperLimit">Upper limit to random numbers</param>
        /// <param name="xMetaTblName">Name of metadata table with field definitions</param>
        /// <param name="WorkbookFullName">File path of workbook</param>
        /// <param name="xResultantTblName">Name of resultant excel table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for inserting resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void AssignRandomNumbers(string xTbl1Name,
            string[] SelectedFields, double LowerLimit, double UpperLimit, 
            string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "ResultTbl", string TargetSheetName = "ResultTbl", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl = Tbl1;

            foreach (var field in SelectedFields)
            {
                if (md.GetFieldType(field) == FieldType.KeyFigure)
                    ResultantTbl = MatrixTable.AssignRandomValues(ResultantTbl, field, LowerLimit, UpperLimit);
                
                else   // date or integer attribute
                    ResultantTbl = MatrixTable.AssignRandomValues(ResultantTbl, field, (int)Math.Round(LowerLimit, 0), (int)Math.Round(UpperLimit, 0));
            }

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Assign uniformly distributed random values to selected key figures or numeric attributes (of date or integer type) of input table,
        /// within lower and upper limits.
        /// </summary>
        public void AssignRandomNumbers_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for AssignRandomNumbers";
                string FuncTitle = "Assign Random Numbers to Selected Fields";
                string FuncDescr = "Assign uniformly distributed random values to selected key figures and numeric attributes (of date or integer type) of input table, within lower and upper limits.";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_AssignRandomNumbers(FormTitle, FuncTitle, FuncDescr, "ResultTbl", "ResultTbl", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTbl1Name = ParameterForm1.xTable1_st;
                string[] SelectedFields = ParameterForm1.listBox1_st;

                double LowerLimit = double.Parse(ParameterForm1.txtBox1_st);
                double UpperLimit = double.Parse(ParameterForm1.txtBox2_st);

                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // call table function
                AssignRandomNumbers(xTbl1Name, SelectedFields, LowerLimit, UpperLimit,
                    xResultantTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("AssignRandomNumbers: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Sort rows of table w.r.t. given fields and sort options (ASC/DESC).
        /// Sort after all attributes in order and ASC if SortStr is null or empty string "".
        /// </summary>
        /// <param name="xTbl1Name">Input table</param>
        /// <param name="SortStr">Comma separated field names and sort option, like "category ASC, product DESC"</param>
        /// <param name="xMetaTblName">Name of metadata table with field definitions</param>
        /// <param name="WorkbookFullName">File path of workbook</param>
        /// <param name="xResultantTblName">Name of resultant excel table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for inserting resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void SortRows(string xTbl1Name, string SortStr,
            string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "ResultTbl", string TargetSheetName = "ResultTbl", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl = MatrixTable.SortRows(Tbl1, SortStr);

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Sort rows of table w.r.t. given fields and sort options (ASC/DESC).
        /// Sort after all attributes in order and ASC if SortStr is null or empty string "".
        /// </summary>
        public void SortRows_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for SortRows";
                string FuncTitle = "Sort Rows of Input Table";
                string FuncDescr = "Sort rows of table w.r.t. given fields and sort options (ASC/DESC). "
                    + "Comma-separated field names and sort options are entered like: category ASC, product DESC, ..\n"
                    + "Sort after all attributes in order and ASC (ascending) if the text box for field names & sort options is left empty. ";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_SortRows(FormTitle, FuncTitle, FuncDescr, "SortedTbl", "SortedTbl", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTbl1Name = ParameterForm1.xTable1_st;
                string SortStr = ParameterForm1.RtxtBox1_st;

                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // call table function
                SortRows(xTbl1Name, SortStr,
                    xResultantTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("SortRows: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Insert a new (output) key figure which is aggregate of selected (input) key figure
        /// w.r.t. selected reference attributes.
        /// For example, KF sales_per_category as aggregate of KF sales w.r.t. attribute category
        /// </summary>
        /// <param name="xTbl1Name">Input table</param>
        /// <param name="ReferenceAttributes">Reference attributes</param>
        /// <param name="InputKeyFig">Input key figure</param>
        /// <param name="OutputKeyFig">Output (aggregate) key figure</param>
        /// <param name="AggregateFunc">Aggregate function</param>
        /// <param name="xMetaTblName">Name of metadata table with field definitions</param>
        /// <param name="WorkbookFullName">File path of workbook</param>
        /// <param name="xResultantTblName">Name of resultant excel table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for inserting resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void InsertAggregateKeyFigure(string xTbl1Name, string[] ReferenceAttributes,
            string InputKeyFig, string OutputKeyFig, string AggregateFunc = "sum",
            string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "UpdatedTbl", string TargetSheetName = "UpdatedTbl", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions

            // get excluded attributes
            /*
            if (IncludedAttributes == null || IncludedAttributes.Count() < 1)
                throw new Exception("Null-valued or empty string vector IncludedAttributes!");

            TextVector InclAttr = new TextVector(IncludedAttributes);
            TextVector ExclTextAttr = TextVector.SetDifference(Tbl1.TextAttributeFields, InclAttr);
            TextVector ExclNumAttr = TextVector.SetDifference(Tbl1.NumAttributeFields, InclAttr);
            TextVector InclKeyFig = new TextVector(IncludedKeyFigures);
             * */

            // get aggregation option from string
            AggregateOption AggrOp = AggregateOption.nSum;
            string AggrOpStr = AggregateFunc.Trim().ToLower();

            switch (AggrOpStr)
            {
                case "sum":
                    AggrOp = AggregateOption.nSum;
                    break;

                case "avg":
                    AggrOp = AggregateOption.nAvg;
                    break;

                case "min":
                    AggrOp = AggregateOption.nMin;
                    break;

                case "max":
                    AggrOp = AggregateOption.nMax;
                    break;

                default:        // normally, should not happen
                    throw new Exception("Undefined aggregation function!");
            }

            // get text vector from string array
            TextVector RefAttributes = new TextVector(ReferenceAttributes);

            MatrixTable ResultantTbl = MatrixTable.AggregateSelectedKeyFigure_B(
                Tbl1, RefAttributes, InputKeyFig, OutputKeyFig, AggrOp);

            // TEST
            // ResultantTbl.ViewTable(xResultantTblName);

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Insert a new (output) key figure which is aggregate of selected (input) key figure
        /// w.r.t. selected reference attributes.
        /// For example, KF sales_per_category as aggregate of KF sales w.r.t. attribute category
        /// </summary>
        public void InsertAggregateKeyFigure_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for InsertAggregateKeyFigure";
                string FuncTitle = "Insert a New Aggregate Key Figure into Table";
                string FuncDescr = "Insert a new (output) key figure which is the aggregate of selected (input) key figure w.r.t. selected reference attributes. "
                    + "For example, KF sales_per_category as aggregate of KF sales w.r.t. attribute category";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_InsertAggregateKeyFigure(FormTitle, FuncTitle, FuncDescr, "UpdatedTbl", "UpdatedTbl", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTbl1Name = ParameterForm1.xTable1_st;
                string AggrFunc = ParameterForm1.comboBox1_st;
                string[] RefAttributes = ParameterForm1.listBox1_st;
                string InputKeyfig = ParameterForm1.InKeyFig1_st;
                string OutputKeyfig = ParameterForm1.txtBox1_st;

                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // aggregate table
                InsertAggregateKeyFigure(xTbl1Name, RefAttributes, InputKeyfig, OutputKeyfig, AggrFunc,
                    xResultantTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("InsertAggregateKeyFigure: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Apply a user-defined filter function (or statement) that returns true or false, 
        /// on every row of input table.
        /// </summary>
        /// <param name="xTbl1Name">Input table</param>
        /// <param name="UserCode">A valid piece of code in C# as user-defined function</param>
        /// <param name="IfExclude">If true, row is exluded when user function returns true</param>
        /// <param name="xMetaTblName">Name of metadata table with field definitions</param>
        /// <param name="WorkbookFullName">File path of workbook</param>
        /// <param name="xResultantTblName">Name of resultant excel table (ListObject)</param>
        /// <param name="TargetSheetName">Target sheet for inserting resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void ApplyUserDefinedFilterFuncOnRows(string xTbl1Name, string UserCode, bool IfExclude = false,
            string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "FilteredTbl", string TargetSheetName = "FilteredTbl",
            string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            HelperFunc.UserDefinedFilterFunction(UserCode);
            MatrixTable ResultantTbl = MatrixTable.FilterRowsDic(Tbl1, HelperFunc.ApplyUserDefinedFilterFuncOnTableRow, IfExclude);

            // Step 5: Write output tables (in this case ResultantTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Row Filter: Apply a user-defined filter function on every row of input table.
        /// </summary>
        public void ApplyUserDefinedFilterFuncOnRows_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for ApplyUserDefinedFilterFuncOnRows";
                string FuncTitle = "Apply User-Defined Filter Function on Table Rows";
                string FuncDescr = "Include or exclude table rows with a user-defined filter function (valid C# code) that returns true or false. "
                    + @"Depending on exclude option, a row is included or excluded when user-defined filter function returns true. "
                    + @"See code examples below assuming that input table contains text attributes ""category"" and ""product"", numeric attributes ""year"" and ""date"", key figures "
                    + @" ""costs"", ""margin"" and ""price"":" + "\n\n"
                    + @"1) return KF[""price""] >= 1.25 * KF[""costs""];" + "\n\n"
                    + @"2) return (TA[""category""] == ""Computers"" && NA[""year""] == 2010);";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_UserDefinedFilterFuncOnRows(FormTitle, FuncTitle, FuncDescr, "ResultTbl", "ResultTbl", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTbl1Name = ParameterForm1.xTable1_st;
                string UserCode = ParameterForm1.RtxtBox2_st;
                bool IfExclude = ParameterForm1.checkBox1_st;
                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // aggregate table
                ApplyUserDefinedFilterFuncOnRows(xTbl1Name, UserCode, IfExclude,
                    xResultantTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ApplyUserDefinedFilterFuncOnRows: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Append table2 to table1 vertically. Two tables must have identical fields.
        /// </summary>
        /// <param name="xTb1Name">Name of 1. excel table (ListObject)</param>
        /// <param name="xTbl2Name">Name of 2. excel table (ListObject)</param>
        /// <param name="xMetaTblName">Name of excel table (ListObject) for field definitions.</param>
        /// <param name="WorkbookFullName">File path of Excel Workbook</param>
        /// <param name="xCombinedTblName">Name of resultant combined table</param>
        /// <param name="TargetSheetName">Sheet name for writing resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void AppendRows(string xTb1Name, string xTbl2Name, string xMetaTblName = null,
                string WorkbookFullName = null, string xCombinedTblName = "ResultTbl", string TargetSheetName = "ResultTbl",
                string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTb1Name);
            XTable xTbl2 = ExcelFunc_NO.GetListObject(wbook, xTbl2Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;
            MatrixTable Tbl2 = ExcelToMatrixTable(xTbl2, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultTbl = MatrixTable.AppendRowsToTable(Tbl1, Tbl2);

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultTbl), wsheet, xCombinedTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Append table2 to table1 vertically. Two tables must have identical fields.
        /// </summary>
        public void AppendRows_macro2()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for AppendRows";
                string FuncTitle = "Append Tables Vertically";
                string FuncDescr = "Append table2 to table1 vertically. Two input tables must have identical fields for this vertical append.";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_AppendRows(FormTitle, FuncTitle, FuncDescr, "ResultTbl", "ResultTbl", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTb1Name = ParameterForm1.xTable1_st;
                string xTbl2Name = ParameterForm1.xTable2_st;
                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // call func
                AppendRows(xTb1Name, xTbl2Name,
                    xCombinedTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("AppendRows: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Append table2 to table1 horizontally. Two tables must have distinct fields with identical number of rows.
        /// </summary>
        /// <param name="xTb1Name">Name of 1. excel table (ListObject)</param>
        /// <param name="xTbl2Name">Name of 2. excel table (ListObject)</param>
        /// <param name="xMetaTblName">Name of excel table (ListObject) for field definitions.</param>
        /// <param name="WorkbookFullName">File path of Excel Workbook</param>
        /// <param name="xCombinedTblName">Name of resultant combined table</param>
        /// <param name="TargetSheetName">Sheet name for writing resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void AppendColumns(string xTb1Name, string xTbl2Name, string xMetaTblName = null,
                string WorkbookFullName = null, string xCombinedTblName = "ResultTbl", string TargetSheetName = "ResultTbl",
                string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTb1Name);
            XTable xTbl2 = ExcelFunc_NO.GetListObject(wbook, xTbl2Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;
            MatrixTable Tbl2 = ExcelToMatrixTable(xTbl2, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultTbl = MatrixTable.AppendColumnsToTable(Tbl1, Tbl2);

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultTbl), wsheet, xCombinedTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Append table2 to table1 horizontally. Two tables must have distinct fields with identical number of rows.
        /// </summary>
        public void AppendColumns_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for AppendColumns";
                string FuncTitle = "Append Tables Horizontally";
                string FuncDescr = "Append table2 to table1 horizontally. Two input tables must have distinct fields with identical number of rows.";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_AppendColumns(FormTitle, FuncTitle, FuncDescr, "ResultTbl", "ResultTbl", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTb1Name = ParameterForm1.xTable1_st;
                string xTbl2Name = ParameterForm1.xTable2_st;
                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // call func
                AppendColumns(xTb1Name, xTbl2Name,
                    xCombinedTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("AppendColumns: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Simple date range filter. 
        /// Selects (or deselects if ExcludeDateRange = true) all rows with dates within the given range.
        /// </summary>
        /// <param name="xTbl1Name">Name of input table</param>
        /// <param name="DateField">Date field of table which is the basis for filtering</param>
        /// <param name="FirstDayOfRange">First day of the range</param>
        /// <param name="LastDayOfRange">Last day of the range</param>
        /// <param name="ExcludeDateRange">If true, exclude all rows within date range</param>
        /// <param name="xMetaTblName">Name of excel table (ListObject) for field definitions.</param>
        /// <param name="WorkbookFullName">File path of Excel Workbook</param>
        /// <param name="xResultantTblName">Name of resultant combined table</param>
        /// <param name="TargetSheetName">Sheet name for writing resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void SimpleDateRangeFilter(string xTbl1Name, string DateField, 
            int FirstDayOfRange, int LastDayOfRange, bool ExcludeDateRange = false,
            string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "FilteredTbl", string TargetSheetName = "FilteredTbl", 
            string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl = MatrixTable.DateRangeFilter(Tbl1, DateField,
                FirstDayOfRange, LastDayOfRange, ExcludeDateRange);

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Simple date range filter. 
        /// Selects (or deselects if ExcludeDateRange = true) all rows with dates within the given range.
        /// </summary>
        public void SimpleDateRangeFilter_macro2()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for DateRangeFilter";
                string FuncTitle = "Date Range Filter";
                string FuncDescr = "Selects (or deselects if exclude option is checked) rows of input table within the given date range. "
                    + "First and last day of range must be entered with a valid date format like 25.08.2012";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_SimpleDateRangeFilter(FormTitle, FuncTitle, FuncDescr, "FilteredTbl", "FilteredTbl", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTb1Name = ParameterForm1.xTable1_st;
                string DateField = ParameterForm1.comboBox1_st;
                int FirstDayOfRange = DateFunctions.DateToNumber(DateTime.Parse(ParameterForm1.txtBox1_st));
                int LastDayOfRange = DateFunctions.DateToNumber(DateTime.Parse(ParameterForm1.txtBox2_st));
                bool IfExclude = ParameterForm1.checkBox1_st;

                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // call func
                SimpleDateRangeFilter(xTb1Name, DateField, FirstDayOfRange, LastDayOfRange,IfExclude,
                    xResultantTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("DateRangeFilter: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Filter days considering date range as well as period (month, quarter, year) and week days.
        /// </summary>
        /// <param name="xTbl1Name">Name of input table</param>
        /// <param name="DateField">Date attribute of input table as basis for date filtering</param>
        /// <param name="Period">Month, Quarter or Year</param>
        /// <param name="FirstDayOfRange">First day of the range</param>
        /// <param name="LastDayOfRange">Last day of the range</param>
        /// <param name="AllowedPeriodDays">Vector with allowed period-days; empty vector means all period-days are allowed. -1 means last day of year.</param>
        /// <param name="AllowedWeekDays">Vector with allowed week-days (1-7); empty vector means all week-days are allowed</param>
        /// <param name="xMetaTblName">Name of excel table (ListObject) for field definitions.</param>
        /// <param name="WorkbookFullName">File path of Excel Workbook</param>
        /// <param name="xResultantTblName">Name of resultant combined table</param>
        /// <param name="TargetSheetName">Sheet name for writing resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void DateFilter(string xTbl1Name, string DateField, string Period, 
            int FirstDayOfRange, int LastDayOfRange,
            NumVector AllowedPeriodDays, NumVector AllowedWeekDays,
            string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "FilteredTbl", string TargetSheetName = "FilteredTbl",
            string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl = null;

            if (Period == "Month")
            {
                ResultantTbl = MatrixTable.FilterDatesMonthly(Tbl1, DateField,
                    FirstDayOfRange, LastDayOfRange, AllowedPeriodDays, AllowedWeekDays);
            }
            else if (Period == "Quarter")
            {
                ResultantTbl = MatrixTable.FilterDatesQuarterly(Tbl1, DateField,
                    FirstDayOfRange, LastDayOfRange, AllowedPeriodDays, AllowedWeekDays);
            }
            else if (Period == "Year")
            {
                ResultantTbl = MatrixTable.FilterDatesYearly(Tbl1, DateField,
                    FirstDayOfRange, LastDayOfRange, AllowedPeriodDays, AllowedWeekDays);
            }
            else // normally, should not happen
                throw new Exception("Unknown perid " + Period + "\n");

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Filter days considering date range as well as period (month, quarter, year) and week days.
        /// </summary>
        public void DateFilter_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for DateFilter";
                string FuncTitle = "Date Filter";
                string FuncDescr = "Filters rows of input table based on selected date field and period (month, quarter, year). "
                    + "Includes or excludes days considering date range as well as allowed period and week days. "
                    + "No selection of week-days means all week-days are permitted. Similarly, all period-days will be allowed if no perid-days are entered.\n\n"
                    + "First and last day of range must be entered with a valid date format like 25.08.2012\n\n"
                    + "Allowed period-days must be separated with commas, like: 1, 15, -1\n"
                    + ".. where -1 means last day of selected period. Allowed ranges for perid-days are (-28 - 28) for Month, (-90 - 90) for Quarter, (-365 - 365) for Year.\n";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_DateFilter(FormTitle, FuncTitle, FuncDescr, "FilteredTbl", "FilteredTbl", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTb1Name = ParameterForm1.xTable1_st;
                string DateField = ParameterForm1.comboBox1_st;
                string Period = ParameterForm1.comboBox2_st;
                int FirstDayOfRange = DateFunctions.DateToNumber(DateTime.Parse(ParameterForm1.txtBox1_st));
                int LastDayOfRange = DateFunctions.DateToNumber(DateTime.Parse(ParameterForm1.txtBox2_st));

                // allowed week days
                string[] AllowedWeekDaysArr = ParameterForm1.listBox1_st;
                NumVector AllowedWeekDays = new NumVector();
                List<int> AllowedWeekDayList = new List<int>();

                Dictionary<string, int> AllowedWeekDayDic = new Dictionary<string, int>();
                AllowedWeekDayDic["Monday"] = 1;
                AllowedWeekDayDic["Tuesday"] = 2;
                AllowedWeekDayDic["Wednesday"] = 3;
                AllowedWeekDayDic["Thursday"] = 4;
                AllowedWeekDayDic["Friday"] = 5;
                AllowedWeekDayDic["Saturday"] = 6;
                AllowedWeekDayDic["Sunday"] = 7;

                if (AllowedWeekDaysArr != null && AllowedWeekDaysArr.Count() > 0)
                {
                    foreach (var AllowedWeekDay in AllowedWeekDaysArr)
                    {
                        AllowedWeekDayList.Add(AllowedWeekDayDic[AllowedWeekDay]);
                    }
                    AllowedWeekDays = new NumVector(AllowedWeekDayList.ToArray());
                }

                // allowed period days
                string AllowedPeriodDaysStr = ParameterForm1.txtBox3_st;
                NumVector AllowedPeriodDays = new NumVector();
                List<int> AllowedPeriodDayList = new List<int>();

                if (AllowedPeriodDaysStr != null && AllowedPeriodDaysStr != "")
                {
                    string[] AllowedDays = AllowedPeriodDaysStr.Split(',');

                    foreach (var day in AllowedDays)
                    {
                        AllowedPeriodDayList.Add(int.Parse(day));
                    }
                    AllowedPeriodDays = new NumVector(AllowedPeriodDayList.ToArray());
                }

                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // call func
                DateFilter(xTb1Name, DateField, Period, FirstDayOfRange, LastDayOfRange,
                    AllowedPeriodDays, AllowedWeekDays,
                    xResultantTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("DateFilter: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Sample dates for allowed period (month, quarter, year) and week days with the given search logic and return 
        /// a subtable with the source and target days.
        /// </summary>
        /// <param name="xTbl1Name">Name of input table</param>
        /// <param name="SourceDate">Source date field of the input table to be sampled</param>
        /// <param name="TargetDate">Target date field contained by the output table; it must not exist in input table.</param>
        /// <param name="Period">Sampling period; month, quarter or year</param>
        /// <param name="FirstDayOfRange">First day of the range</param>
        /// <param name="LastDayOfRange">Last day of the range</param>
        /// <param name="search_logic">Search logic</param>
        /// <param name="MaxDistance">Maximum allowed distance in days between the source date and target date</param>
        /// <param name="AllowedPeriodDays">Vector with allowed period-days; empty vector means all period-days are allowed. -1 means last day of year.</param>
        /// <param name="AllowedWeekDays">Vector with allowed week-days (1-7); empty vector means all week-days are allowed</param>
        /// <param name="xMetaTblName">Name of excel table (ListObject) for field definitions.</param>
        /// <param name="WorkbookFullName">File path of Excel Workbook</param>
        /// <param name="xResultantTblName">Name of resultant combined table</param>
        /// <param name="TargetSheetName">Sheet name for writing resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void DateSampler(string xTbl1Name, string SourceDate, string TargetDate,  
            string Period, int FirstDayOfRange, int LastDayOfRange,
            SearchLogic search_logic, int MaxDistance,  
            NumVector AllowedPeriodDays, NumVector AllowedWeekDays,
            string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "SampledTbl", string TargetSheetName = "SampledTbl",
            string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xTbl1Name);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable ResultantTbl = null;

            if (Period == "Month")
            {
                ResultantTbl = MatrixTable.SampleDatesMonthly_A(Tbl1, SourceDate, TargetDate,
                    FirstDayOfRange, LastDayOfRange, search_logic, MaxDistance,
                    AllowedPeriodDays, AllowedWeekDays);
            }
            else if (Period == "Quarter")
            {
                ResultantTbl = MatrixTable.SampleDatesQuarterly_A(Tbl1, SourceDate, TargetDate,
                    FirstDayOfRange, LastDayOfRange, search_logic, MaxDistance,
                    AllowedPeriodDays, AllowedWeekDays);
            }
            else if (Period == "Year")
            {
                ResultantTbl = MatrixTable.SampleDatesYearly_A(Tbl1, SourceDate, TargetDate,
                    FirstDayOfRange, LastDayOfRange, search_logic, MaxDistance,
                    AllowedPeriodDays, AllowedWeekDays);
            }
            else // normally, should not happen
                throw new Exception("Unknown perid " + Period + "\n");

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(ResultantTbl), wsheet, xResultantTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Sample dates for allowed period (month, quarter, year) and week days with the given search logic and return 
        /// a subtable with the source and target days.
        /// </summary>
        public void DateSampler_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for DateSampler";
                string FuncTitle = "Sample Table for Target Dates";
                string FuncDescr = "Samples dates for allowed period (month, quarter, year) and week days with the given search logic and returns "
                    + "a subtable with the source and target days. "
                    + "No selection of week-days means all week-days are permitted. Similarly, all period-days will be allowed if no perid-days are entered.\n\n"
                    + "First and last day of range must be entered with a valid date format like 25.08.2012\n\n"
                    + "Allowed period-days must be separated with commas, like: 1, 15, -1\n"
                    + ".. where -1 means last day of selected period. Allowed ranges for perid-days are (-28 - 28) for Month, (-90 - 90) for Quarter, (-365 - 365) for Year.\n";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_DateSampler(FormTitle, FuncTitle, FuncDescr, "SampledTbl", "SampledTbl", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTb1Name = ParameterForm1.xTable1_st;
                string SourceDate = ParameterForm1.comboBox1_st;
                string TargetDate = ParameterForm1.txtBox1_st;
                string Period = ParameterForm1.comboBox2_st;
                int FirstDayOfRange = DateFunctions.DateToNumber(DateTime.Parse(ParameterForm1.txtBox2_st));
                int LastDayOfRange = DateFunctions.DateToNumber(DateTime.Parse(ParameterForm1.txtBox3_st));
                int MaxDistance = int.Parse(ParameterForm1.txtBox4_st);

                // search logic
                string SearchLogicStr = ParameterForm1.comboBox3_st;
                SearchLogic SLogic = SearchLogic.Previous;
                string[] SearchOpts= new string[] { "Previous date", "Next date", "Nearest date", "Exact date" };

                if (SearchLogicStr == SearchOpts[0])
                    SLogic = SearchLogic.Previous;
                else if (SearchLogicStr == SearchOpts[1])
                    SLogic = SearchLogic.Next;
                else if (SearchLogicStr == SearchOpts[2])
                    SLogic = SearchLogic.Nearest;
                else if (SearchLogicStr == SearchOpts[3])
                    SLogic = SearchLogic.Exact;

                // allowed week days
                string[] AllowedWeekDaysArr = ParameterForm1.listBox1_st;
                NumVector AllowedWeekDays = new NumVector();
                List<int> AllowedWeekDayList = new List<int>();

                Dictionary<string, int> AllowedWeekDayDic = new Dictionary<string, int>();
                AllowedWeekDayDic["Monday"] = 1;
                AllowedWeekDayDic["Tuesday"] = 2;
                AllowedWeekDayDic["Wednesday"] = 3;
                AllowedWeekDayDic["Thursday"] = 4;
                AllowedWeekDayDic["Friday"] = 5;
                AllowedWeekDayDic["Saturday"] = 6;
                AllowedWeekDayDic["Sunday"] = 7;

                if (AllowedWeekDaysArr != null && AllowedWeekDaysArr.Count() > 0)
                {
                    foreach (var AllowedWeekDay in AllowedWeekDaysArr)
                    {
                        AllowedWeekDayList.Add(AllowedWeekDayDic[AllowedWeekDay]);
                    }
                    AllowedWeekDays = new NumVector(AllowedWeekDayList.ToArray());
                }

                // allowed period days
                string AllowedPeriodDaysStr = ParameterForm1.txtBox5_st;
                NumVector AllowedPeriodDays = new NumVector();
                List<int> AllowedPeriodDayList = new List<int>();

                if (AllowedPeriodDaysStr != null && AllowedPeriodDaysStr != "")
                {
                    string[] AllowedDays = AllowedPeriodDaysStr.Split(',');

                    foreach (var day in AllowedDays)
                    {
                        AllowedPeriodDayList.Add(int.Parse(day));
                    }
                    AllowedPeriodDays = new NumVector(AllowedPeriodDayList.ToArray());
                }

                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // call func
                DateSampler(xTb1Name, SourceDate, TargetDate, Period, 
                    FirstDayOfRange, LastDayOfRange, SLogic, MaxDistance,
                    AllowedPeriodDays, AllowedWeekDays,
                    xResultantTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("DateSampler: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Get price table with costs, margins and prices (user-defined table function example)
        /// </summary>
        /// <param name="xCostTable">Name of cost table with key figure named "costs"</param>
        /// <param name="xMarginTable">Name of margin table with key figure named "margin"</param>
        /// <param name="xMetaTblName">Name of excel table (ListObject) for field definitions.</param>
        /// <param name="WorkbookFullName">File path of Excel Workbook</param>
        /// <param name="xResultTblName">Name of resultant combined table</param>
        /// <param name="TargetSheetName">Sheet name for writing resultant table</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void GetPriceTable2(string xCostTable, string xMarginTable, string xMetaTblName = null,
            string WorkbookFullName = null, string xResultTblName = "PriceTbl", 
            string TargetSheetName = "PriceTbl", string TopLeftCell = "A1") 
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xTbl1 = ExcelFunc_NO.GetListObject(wbook, xCostTable);
            XTable xTbl2 = ExcelFunc_NO.GetListObject(wbook, xMarginTable);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable Tbl1 = ExcelToMatrixTable(xTbl1, mdx).matrixTable;
            MatrixTable Tbl2 = ExcelToMatrixTable(xTbl2, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable PriceTbl = UserFunc.GetPriceTable(Tbl1, Tbl2);

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, TargetSheetName, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(PriceTbl), wsheet, xResultTblName, TopLeftCell);
            wsheet.Activate();
        }

        /// <summary>
        /// Get price table with costs, margins and prices (user-defined table function example)
        /// </summary>
        public void GetPriceTable2_macro2()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for GetPriceTable2";
                string FuncTitle = "Get Price Table";
                string FuncDescr = "Get price table with costs, margins and prices. "
                    + "Cost table must contain a key figure named costs, and margin table must contain a key figure named margin. "
                    + "Resultant price table contains all three key figures: costs, margin and price";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_GetPriceTable2(FormTitle, FuncTitle, FuncDescr, "PriceTbl", "PriceTbl", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string xTb1Name = ParameterForm1.xTable1_st;
                string xTbl2Name = ParameterForm1.xTable2_st;

                string xOutTable = ParameterForm1.xOutTable1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // combine tables
                GetPriceTable2(xTb1Name, xTbl2Name,
                    xResultTblName: xOutTable, TargetSheetName: SheetName,
                    TopLeftCell: TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("GetPriceTable2: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Calculate sales commissions per product pool and dealer with tiered commission rates.
        /// See related article: http://finaquant.com/commission-calculation-with-finaquant-calcs/3607
        /// </summary>
        /// <param name="Period">Calculation and payment period, 'month' or 'quarter'</param>
        /// <param name="xSalesTable">Input table with with sales transaction data for each dealer</param>
        /// <param name="xComScaleTable">Input table with tiered rates for each commission scale</param>
        /// <param name="xScaleToPoolTable">Input table that maps commission scales to product pools, with additional scale logic (class or level)</param>
        /// <param name="xProductPoolTable">Input table that maps categories and products to a product pool for each dealer</param>
        /// <param name="xMetaTblName">Name of excel table (ListObject) for field definitions.</param>
        /// <param name="WorkbookFullName">File path of Excel Workbook</param>
        /// <param name="xComPerPoolTbl">Output table: Sales Commissions per Product Pool</param>
        /// <param name="xComPerDealerTbl">Output table: Sales Commissions per Dealer</param>
        /// <param name="SheetSalesCom">Sheet name for writing resultant tables</param>
        /// <param name="TopLeftCell">Cell address of upper-left corner for output table</param>
        public void CalculateSalesCommissions(string Period, string xSalesTable, string xComScaleTable,
            string xScaleToPoolTable, string xProductPoolTable,
            string xMetaTblName = null, string WorkbookFullName = null,
            string xComPerPoolTbl = "SalesComPerPool", 
            string xComPerDealerTbl = "SalesComPerDealer",
            string SheetSalesCom = "SalesCom", string TopLeftCell = "A1") 
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Get ListObjects
            XTable xSalesTbl = ExcelFunc_NO.GetListObject(wbook, xSalesTable);
            XTable xComScaleTbl = ExcelFunc_NO.GetListObject(wbook, xComScaleTable);
            XTable xScaleToPoolTbl = ExcelFunc_NO.GetListObject(wbook, xScaleToPoolTable);
            XTable xProductPoolTbl = ExcelFunc_NO.GetListObject(wbook, xProductPoolTable);

            XTable xMetaTbl = null;
            if (xMetaTblName != null && xMetaTblName != String.Empty) xMetaTbl = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Step 2: Get meta data with field definitions
            MetaDataX mdx = new MetaDataX();
            if (xMetaTblName != null && xMetaTblName != String.Empty) mdx.ReadFieldsFromExcelTable(xMetaTbl);
            MetaData md = mdx.metaData;

            // Step 3: Read excel tables (inputs) into MatrixTable objects
            MatrixTable SalesTbl = ExcelToMatrixTable(xSalesTbl, mdx).matrixTable;
            MatrixTable ComScaleTbl = ExcelToMatrixTable(xComScaleTbl, mdx).matrixTable;
            MatrixTable ScaleToPoolTbl = ExcelToMatrixTable(xScaleToPoolTbl, mdx).matrixTable;
            MatrixTable ProductPoolTbl = ExcelToMatrixTable(xProductPoolTbl, mdx).matrixTable;

            // Step 4: Generate resultant (output) tables with table functions
            MatrixTable CommissionsPerPool, CommissionsPerDealer;

            UserFunc.CalculateSalesCommissions(Period, SalesTbl, ComScaleTbl, ScaleToPoolTbl, ProductPoolTbl,
                out CommissionsPerPool, out CommissionsPerDealer);

            // Step 5: Write output tables (in this case CombinedTbl) into excel tables
            Excel.Worksheet wsheet = ExcelFunc_NO.GetWorksheet(wbook, SheetSalesCom, AddSheetIfNotFound: true);

            MatrixTableToExcel(new MatrixTableX(CommissionsPerPool), wsheet, xComPerPoolTbl, TopLeftCell);
            
            MatrixTableToExcel(new MatrixTableX(CommissionsPerDealer), wsheet, xComPerDealerTbl,
                ExcelFunc_NO.ShiftCell(wsheet, TopLeftCell, 0, CommissionsPerPool.ColumnCount + 2));
            wsheet.Activate();
        }

        /// <summary>
        /// Calculate sales commissions per product pool and dealer with tiered commission rates.
        /// See related article: http://finaquant.com/commission-calculation-with-finaquant-calcs/3607
        /// </summary>
        public void CalculateSalesCommissions_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;
                ParameterForm1.md_st = new MetaData();

                string FormTitle = "Parameter Form for CalculateSalesCommissions";
                string FuncTitle = "Calculate Sales Commissions";
                string FuncDescr = "Calculate sales commissions per product pool and per dealer, given four input tables: "
                    + "Sales table, Commission-Scale table, Product-Pool table, Scale-To-Pool table. "
                    + "Two output tables will be generated: Commissions per Product Pool, and Commissions per Dealer";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_CalculateSalesCom(FormTitle, FuncTitle, FuncDescr, "CommissionsPerPool", "CommissionsPerDealer", "SalesCommissions", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values

                // input tables
                string SalesTbl = ParameterForm1.xTable1_st;
                string ComScaleTbl = ParameterForm1.xTable2_st;
                string ProductPoolTbl = ParameterForm1.comboBox1_st;
                string ScaleToPoolTbl = ParameterForm1.comboBox2_st;

                // period
                string period = ParameterForm1.comboBox3_st;

                // output tables
                string ComPerPool = ParameterForm1.xOutTable1_st;
                string ComPerDealer= ParameterForm1.txtBox1_st;

                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // calculate commissions
                CalculateSalesCommissions(period, SalesTbl, ComScaleTbl, ScaleToPoolTbl, ProductPoolTbl,
                    xComPerPoolTbl: ComPerPool, xComPerDealerTbl: ComPerDealer, 
                    SheetSalesCom: SheetName, TopLeftCell: TopLeftCell); 
            }
            catch (Exception ex)
            {
                MessageBox.Show("CalculateSalesCommissions: " + ex.Message + "\n");
            }
        }

    }

    // requires "using ExcelDna.ComInterop"
    [ComVisible(false)]
    class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ComServer.DllRegisterServer();
        }
        public void AutoClose()
        {
            ComServer.DllUnregisterServer();
        }
    }

    [ComVisible(true)]
    public class Ribbon : CustomUI.ExcelRibbon
    {
        public void RunControlID(CustomUI.IRibbonControl ctl)
        {
            ((Excel.Application)ExcelDnaUtil.Application).Run(ctl.Id);
        }

        public void RunControlIDWithTag(CustomUI.IRibbonControl ctl)
        {
            ((Excel.Application)ExcelDnaUtil.Application).Run(ctl.Id, ctl.Tag);
        }
    }

}
