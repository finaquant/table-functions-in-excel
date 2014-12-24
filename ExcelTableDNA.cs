// Finaquant Analytics - http://finaquant.com/
// Copyright © Finaquant Analytics GmbH
// Email: support@finaquant.com

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using FinaquantCalcs;

// NetOffice: Microsoft Office integration without version limitations
// see: http://netoffice.codeplex.com/
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.GlobalHelperModules;
using NetOffice.ExcelApi.Enums;
using XTable = NetOffice.ExcelApi.ListObject;

// Excel DNA: Integrate .NET into Excel
// http://exceldna.codeplex.com/
using ExcelDna.Integration;
using ExcelDna.ComInterop;
using CustomUI = ExcelDna.Integration.CustomUI;

// Calling a static method of class ExcelTableDNA in Excel VBA
/*
'Call static .net method MultiplyThem() in VBA
Sub Test_MultiplyThem()
Dim x As Double

x = Application.Run("MultiplyThem", 2, 3)
Debug.Print "Run: " & Application.Run("MultiplyThem", 2, 3)
Debug.Print "x = " & x

End Sub
*/

// Notes
/* Method's return type must be 'void' for [ExcelCommand(..)], otherwise
 * .. it doesn't appear in command menu in add-in tab.
 * 
 */

namespace FinaquantInExcel
{
    public class ExcelTableDNA
    {
        // [ExcelCommand(MenuName = "Test Functions", MenuText = "Add Numbers")]
        public static void AddNumbers(double x, double y)
        {
            double z = x + y;
            // Console.WriteLine("x + y = " + z);
            MessageBox.Show("x + y = " + z.ToString());
        }

        // [ExcelFunction(Description = "Multiplies two numbers", Category = "Test Functions")]
        public static double MultiplyNumbers(double x, double y)
        {
            return x * y;
        }

        /// <summary>
        /// Show user license form
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Help My License")]
        public static void HelpCalcsLicense()
        {
            HelpCalcs.ShowLicenseInformation();
        }

        /// <summary>
        /// Show product information form
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Help My Product")]
        public static void HelpCalcsProduct()
        {
            HelpCalcs.ProductInformation();
        }

        /// <summary>
        /// Create test tables in excel
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Create Test Tables")]
        public static void CreateTestTables_macro()
        {
            var xt = new ExcelTable();
            xt.CreateTestTables_macro();
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
        public static void GetPriceTable(string xCostTblName, string xMarginTblName, string xMetaTblName = null,
            string WorkbookFullName = null, string xPriceTblName = "Price", string TargetSheetName = "PriceTable",
            string inKeyFig1 = "costs", string inKeyFig2 = "margin", string outKeyFig = "price",
            double StdMargin = 0.25, string CellStr = "A1", bool ClearSheetContent = true)
        {
            {
                try
                {
                    var xt = new ExcelTable();
                    xt.GetPriceTable( xCostTblName, xMarginTblName, xMetaTblName,
                        WorkbookFullName, xPriceTblName, TargetSheetName,
                        inKeyFig1, inKeyFig2, outKeyFig, StdMargin, CellStr, ClearSheetContent);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ExcelTableDNA.GetPriceTable: " + ex.Message + "\n");
                }
            }

        }

        /// <summary>
        /// Obtain price table from cost and margin tables; shows input box for getting parameter values
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Get Price Table (InputBox)")]
        public static void GetPriceTable_macro()
        {
            var xt = new ExcelTable();
            xt.GetPriceTable_macro();
        }

        /// <summary>
        /// Combine two input tables that have some common attributes (text or numeric).
        /// Underlying Table Function: MatrixTable.CombineTables()
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
        /// <remarks>
        /// Excel Tables are passed by their names (string); no excel object parameters.
        /// Set xMetaTblName to null if there is no excel table for field definitions.
        /// </remarks>
        public static void CombineTables(string xTb1Name, string xTbl2Name, string xMetaTblName = null,
                string xCombinedTblName = "Combined", string TargetSheetName = "CombinedTable",
                bool JokerMatchesAllvalues = true, string TextJoker = "ALL", int NumJoker = 0)
        {
            try
            {
                var xt = new ExcelTable();
                xt.CombineTables(xTb1Name, xTbl2Name, xMetaTblName,
                    null, xCombinedTblName, TargetSheetName,
                    JokerMatchesAllvalues, TextJoker, NumJoker);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ExcelTableDNA.CombineTables2: " + ex.Message);
            }
        }

        /// <summary>
        /// Combine two input tables that have some common attributes (text or numeric).
        /// Shows input box for getting parameters.
        /// Underlying Table Function: MatrixTable.CombineTables()
        /// </summary>
        //[ExcelCommand(MenuName = "Table Functions", MenuText = "Combine Tables (InputBox)")]
        public static void CombineTables_macro()
        {
            var xt = new ExcelTable();
            xt.CombineTables_macro();
        }

        /// <summary>
        /// Combine two input tables that have some common attributes (text or numeric).
        /// Shows input box for getting parameters.
        /// Underlying Table Function: MatrixTable.CombineTables()
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Combine Tables")]
        public static void CombineTables_macro2() 
        {
            var xt = new ExcelTable();
            xt.CombineTables_macro2();
        }

        /// <summary>
        /// Add two input tables (table addition) with selected key figures.
        /// Calls input range box for getting parameter values.
        /// Underlying table function: MatrixTable.AddSelectedKeyFigures()
        /// </summary>
        //[ExcelCommand(MenuName = "Table Functions", MenuText = "Add Tables (InputBox)")]
        public static void AddTables_macro()
        {
            var xt = new ExcelTable();
            xt.AddTables_macro();
        }

        /// <summary>
        /// Add two input tables (table addition) with selected key figures.
        /// Calls input range box for getting parameter values.
        /// Underlying table function: MatrixTable.AddSelectedKeyFigures()
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Add Tables")]
        public static void AddTables_macro2()
        {
            var xt = new ExcelTable();
            xt.AddTables_macro2();
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
        public void MultiplyTables(string xTbl1Name, string xTbl2Name,
            string InputKeyfig1, string InputKeyfig2, string OutputKeyfig,
             string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "Result", string TargetSheetName = "ResultTable")
        {
            try
            {
                var xt = new ExcelTable();

                xt.MultiplyTables( xTbl1Name, xTbl2Name,
                    InputKeyfig1, InputKeyfig2, OutputKeyfig,
                    xMetaTblName, WorkbookFullName,
                    xResultantTblName, TargetSheetName);
            }
            catch (Exception ex)
            {
                throw new Exception("ExcelTableDNA.CombineTables2: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Multiply two input tables (table multiplication) with selected key figures.
        /// Calls input range box for getting parameter values.
        /// Underlying table function: MatrixTable.MultiplySelectedKeyFigures()
        /// </summary>
        //[ExcelCommand(MenuName = "Table Functions", MenuText = "Multiply Tables (InputBox)")]
        public static void MultiplyTables_macro()
        {
            var xt = new ExcelTable();
            xt.MultiplyTables_macro();
        }

        /// <summary>
        /// Table Arithmetics: Addition, Multiplication, Subtraction or Division
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Table Arithmetics")]
        public static void TableArithmetics_macro()
        {
            var xt = new ExcelTable();
            xt.TableArithmetics_macro();
        }

        /// <summary>
        /// Simple pro-rate distribution; returns table with distributed amounts.
        /// The source amount in SourceTable is disrubuted to target amounts w.r.t. keys (or ratios) given in KeyTable.
        /// Important: Both Source and Key tables must have exactly one key figure.
        /// </summary>
        /// <param name="xSourceTblName">Name of source table with source amounts to be distributed</param>
        /// <param name="DistrKeyTblName">Name if key table with key amounts (or ratios) used for the distribution</param>
        /// <param name="ResultantKeyFig">Name of key figure with the resultant distributed amounts</param>
        /// <param name="KeySumKeyFig">Name of the key figure with the sum of key amounts</param>
        /// <param name="xMetaTblName">Name of metadata table with field definitions</param>
        /// <param name="WorkbookFullName">File path of workbook</param>
        /// <param name="xResultantTblName">Name of resultant combined table with distributed amounts</param>
        /// <param name="TargetSheetName">Target sheet for resultant table</param>
        public static void DistributeTable(string xSourceTblName, string DistrKeyTblName,
            string ResultantKeyFig, string KeySumKeyFig,
            string xMetaTblName = null, string WorkbookFullName = null,
            string xResultantTblName = "Distrib", string TargetSheetName = "DistribTable")
        {
            try
            {
                var xt = new ExcelTable();
                xt.DistributeTable( xSourceTblName,  DistrKeyTblName,
                     ResultantKeyFig,  KeySumKeyFig,
                     xMetaTblName,  WorkbookFullName,
                     xResultantTblName,  TargetSheetName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ExcelTableDNA.DistributeTable: " + ex.Message);
            }
        }

        /// <summary>
        /// Simple pro-rate distribution; returns table with distributed amounts.
        /// </summary>
        //[ExcelCommand(MenuName = "Table Functions", MenuText = "Distribute Table (InputBox)")]
        public static void DistributeTable_macro()
        {
            var xt = new ExcelTable();
            xt.DistributeTable_macro();
        }

        /// <summary>
        /// Simple pro-rate distribution; returns table with distributed amounts.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Distribute Table")]
        public static void DistributeTable_macro2()
        {
            var xt = new ExcelTable();
            xt.DistributeTable_macro2();
        }

        /// <summary>
        /// Create a new table with all possible combinations of given field values in input table.
        /// </summary>
        //[ExcelCommand(MenuName = "Table Functions", MenuText = "Combinate Field Values (InputBox)")]
        public static void CombinateFieldValues_macro()
        {
            var xt = new ExcelTable();
            xt.CombinateFieldValues_macro();
        }

        /// <summary>
        /// Create a new table with all possible combinations of given field values in input table.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Combinate Field Values")]
        public static void CombinateFieldValues_macro2()
        {
            var xt = new ExcelTable();
            xt.CombinateFieldValues_macro2();
        }

        /// <summary>
        /// Filter input table BaseTbl with a condition table CondTbl.
        /// if ExcludeMatchedRows = false (default case) output table includes only rows that match with
        /// some row(s) of condition table CondTbl.
        /// </summary>
        //[ExcelCommand(MenuName = "Table Functions", MenuText = "Filter Table (InputBox)")]
        public static void FilterTable_macro()
        {
            var xt = new ExcelTable();
            xt.FilterTable_macro();
        }

        /// <summary>
        /// Filter input table BaseTbl with a condition table CondTbl.
        /// if ExcludeMatchedRows = false (default case) output table includes only rows that match with
        /// some row(s) of condition table CondTbl.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Filter Table")]
        public static void FilterTable_macro2() 
        {
            var xt = new ExcelTable();
            xt.FilterTable_macro2();
        }

        /// <summary>
        /// Cartesian Multiplication of table rows: Generate a new table 
        /// with all possible row combinations of input tables. 
        /// There must be no common fields among input tables.
        /// </summary>
        //[ExcelCommand(MenuName = "Table Functions", MenuText = "Combinate Rows (InputBox)")]
        public static void CombinateRows_macro()
        {
            var xt = new ExcelTable();
            xt.CombinateRows_macro();
        }

        /// <summary>
        /// Cartesian Multiplication of table rows: Generate a new table 
        /// with all possible row combinations of input tables. 
        /// There must be no common fields among input tables.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Combinate Rows")]
        public static void CombinateRows_macro2()
        {
            var xt = new ExcelTable();
            xt.CombinateRows_macro2();
        }

        /// <summary>
        /// Aggregate selected key figures of table with the selected (default) aggregation function. 
        /// Resultant aggregated table contains selected attributes and key figures.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Aggregate Table")]
        public static void AggregateTable_macro2()
        {
            var xt = new ExcelTable();
            xt.AggregateTable_macro2();
        }

        /// <summary>
        /// Row Transformer: Apply a user-defined function on every row of input table.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Transform Rows with UDF")]
        public static void ApplyUserDefinedTransformFuncOnRows_macro2()
        {
            var xt = new ExcelTable();
            xt.ApplyUserDefinedTransformFuncOnRows_macro2();
        }

        /// <summary>
        /// Return subtable with selected columns of input table.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Select Columns")]
        public static void SelectColumns_macro2()
        {
            var xt = new ExcelTable();
            xt.SelectColumns_macro2();
        }

        /// <summary>
        /// Apply selected scalar arithmetical operation (Addition, Multiplication, Subtraction, Division)
        /// on selected key figures of input table, like adding a scalar value to selected key figures.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Scalar Arithmetic Operation")]
        public static void ScalarOperation_macro2()
        {
            var xm = new ExcelTable();
            xm.ScalarOperation_macro2();
        }

        /// <summary>
        /// Round selected key figures to given number of digits after decimal point.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Round Key Figures")]
        public static void RoundNumbers_macro2()
        {
            var xm = new ExcelTable();
            xm.RoundNumbers_macro2();
        }

        /// <summary>
        /// Insert a new field with a constant value into input table
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Insert New Field")]
        public static void InsertNewField_macro2()
        {
            var xm = new ExcelTable();
            xm.InsertNewField_macro2();
        }

        /// <summary>
        /// Assign uniformly distributed random values to selected key figures or numeric attributes (of date or integer type) of input table,
        /// within lower and upper limits.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Assign Random Numbers")]
        public static void AssignRandomNumbers_macro2()
        {
            var xm = new ExcelTable();
            xm.AssignRandomNumbers_macro2();
        }

        /// <summary>
        /// Sort rows of table w.r.t. given fields and sort options (ASC/DESC).
        /// Sort after all attributes in order and ASC if SortStr is null or empty string "".
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Sort Rows")]
        public static void SortRows_macro2()
        {
            var xm = new ExcelTable();
            xm.SortRows_macro2();
        }

        /// <summary>
        /// Insert a new (output) key figure which is aggregate of selected (input) key figure
        /// w.r.t. selected reference attributes.
        /// For example, KF sales_per_category as aggregate of KF sales w.r.t. attribute category
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Insert Aggregate Key Figure")]
        public static void InsertAggregateKeyFigure_macro2()
        {
            var xm = new ExcelTable();
            xm.InsertAggregateKeyFigure_macro2();
        }

        /// <summary>
        /// Row Filter: Apply a user-defined filter function on every row of input table.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Filter Rows with UDF")]
        public static void ApplyUserDefinedFilterFuncOnRows_macro2()
        {
            var xm = new ExcelTable();
            xm.ApplyUserDefinedFilterFuncOnRows_macro2();
        }

        /// <summary>
        /// Append table2 to table1 vertically. Two tables must have identical fields.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Append Tables Vertically")]
        public static void AppendRows_macro2()
        {
            var xm = new ExcelTable();
            xm.AppendRows_macro2();
        }

        /// <summary>
        /// Append table2 to table1 horizontally. Two tables must have distinct fields with identical number of rows.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Append Tables Horizontally")]
        public static void AppendColumns_macro2()
        {
            var xm = new ExcelTable();
            xm.AppendColumns_macro2();
        }

        /// <summary>
        /// Simple date range filter. 
        /// Selects (or deselects if ExcludeDateRange = true) all rows with dates within the given range.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Date Range Filter")]
        public static void SimpleDateRangeFilter_macro2()
        {
            var xm = new ExcelTable();
            xm.SimpleDateRangeFilter_macro2();
        }

        /// <summary>
        /// Filter days considering date range as well as period (month, quarter, year) and week days.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Date Filter")]
        public static void DateFilter_macro2()
        {
            var xm = new ExcelTable();
            xm.DateFilter_macro2();
        }

        /// <summary>
        /// Sample dates for allowed period (month, quarter, year) and week days with the given search logic and return 
        /// a subtable with the source and target days.
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Date Sampler")]
        public static void DateSampler_macro2()
        {
            var xm = new ExcelTable();
            xm.DateSampler_macro2();
        }

        /// <summary>
        /// Get price table with costs, margins and prices (user-defined table function example)
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Get Price Table (User Defined)")]
        public static void GetPriceTable2_macro2()
        {
            var xm = new ExcelTable();
            xm.GetPriceTable2_macro2();
        }

        /// <summary>
        /// Calculate sales commissions per product pool and dealer with tiered commission rates.
        /// See related article: http://finaquant.com/commission-calculation-with-finaquant-calcs/3607
        /// </summary>
        [ExcelCommand(MenuName = "Table Functions", MenuText = "Calculate Sales Commissions (User Defined)")]
        public static void CalculateSalesCommissions_macro2()
        {
            var xm = new ExcelTable();
            xm.CalculateSalesCommissions_macro2();
        }

    }
}
