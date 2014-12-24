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
using FinaquantCalcs;
using System.Reflection;

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

// useful links for Excel Table (ListObject):
// http://www.jkp-ads.com/articles/Excel2007TablesVBA.asp
// http://excelandaccess.wordpress.com/2013/01/07/working-with-tables-in-vba/
// http://stackoverflow.com/questions/3070123/how-to-loop-though-a-table-and-access-row-items-by-their-column-header
// http://msdn.microsoft.com/en-us/library/microsoft.office.tools.excel.listobject.aspx
//
// If there is no header row, HeaderRowRange returns null:
// http://msdn.microsoft.com/en-us/library/microsoft.office.tools.excel.listobject.headerrowrange.aspx
//
// Determining the Type of a Variant
// http://www.java2s.com/Code/VBA-Excel-Access-Word/Data-Type/DeterminingtheTypeofaVariant.htm

namespace FinaquantInExcel
{

    /// <summary>
    /// Class with table-valued functions (Table Functions) for excel VBA. 
    /// A MatrixTableX object represents a data table. 
    /// All public and non-static methods in this class are available in Excel VBA.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("Finaquant_MatrixTableX")]
    public class MatrixTableX
    {
        private MatrixTable _tbl = null; 

    #region "MatrixTableX constructors"

        /// <summary>
        /// Initialize an empty table
        /// </summary>
        public MatrixTableX() { _tbl = new MatrixTable(); }

        /// <summary>
        /// Initialize a new table
        /// </summary>
        internal MatrixTableX(MatrixTable tbl)
        {
            this._tbl = tbl;
        }

        /// <summary>
        /// Initialize a new table
        /// </summary>
        /// <param name="MatrixTableObj">MatrixTable object</param>
        public MatrixTableX(object MatrixTableObj)
        {
            // see is and as operators
            // http://msdn.microsoft.com/en-us/library/cc488006.aspx
            // http://stackoverflow.com/questions/3786361/difference-between-is-and-as-keyword

            if (MatrixTableObj is MatrixTable)
                this._tbl = (MatrixTable)MatrixTableObj;
            else
                throw new Exception("MatrixTableX: Input parameter MatrixTableObj must be a MatrixTable instance!");
        }

        #endregion "MatrixTableX constructors"

    #region "MatrixTableX properties" 

        /// <summary>
        /// Get and Set underlying MatrixTable object
        /// </summary>
        internal MatrixTable matrixTable
        {
            get { return this._tbl; }
        }

        /// <summary>
        /// Get underlying MatrixTable instance (object)
        /// </summary>
        public object MatrixTableObj
        {
            get { return this._tbl; }
        }

        /// <summary>
        /// Return true if table has no fields at all
        /// </summary>
        public bool IsEmpty
        {
            get { return this._tbl.IsEmpty; }
        }

        /// <summary>
        /// Return all ordered field names (text, numeric, key figure)
        /// </summary>
        public string[] ColumnNames
        {
            get
            {
                // concatenate arrays in order
                var strl = new List<string>();

                if (! this._tbl.TextAttributeFields.IsEmpty && this._tbl.TextAttributeFields.toArray != null) 
                    strl.AddRange(this._tbl.TextAttributeFields.toArray);

                if (! this._tbl.NumAttributeFields.IsEmpty && this._tbl.NumAttributeFields.toArray != null)
                    strl.AddRange(this._tbl.NumAttributeFields.toArray);

                if (! this._tbl.KeyFigureFields.IsEmpty && this._tbl.KeyFigureFields.toArray != null)
                    strl.AddRange(this._tbl.KeyFigureFields.toArray);

                return strl.ToArray();
            }
        }

        /// <summary>
        /// Number of rows in table
        /// </summary>
        public int RowCount
        {
            get { return this._tbl.RowCount; }
        }

        /// <summary>
        /// Number of fields (columns) in table
        /// </summary>
        public int ColumnCount
        {
            get { return this._tbl.ColumnCount; }
        }

        /// <summary>
        /// Get ordered field names of text attributes
        /// </summary>
        public string[] TextAttributeFields
        {
            get { return this._tbl.TextAttributeFields.toArray; }
        }

        /// <summary>
        /// Get ordered field names of numeric attributes
        /// </summary>
        public string[] NumAttributeFields
        {
            get { return this._tbl.NumAttributeFields.toArray; }
        }

        /// <summary>
        /// Get ordered field names of key figures
        /// </summary>
        public string[] KeyFigureFields
        {
            get { return this._tbl.KeyFigureFields.toArray; }
        }

        /// <summary>
        /// Get value matrix for text attributes; columns match ordered fields
        /// </summary>
        public string[,] TextAttribValues
        {
            get { return this._tbl.TextAttribValues.toArray; }
        }

        /// <summary>
        /// Get value matrix for numeric attributes; columns match ordered fields
        /// </summary>
        public int[,] NumAttribValues
        {
            get { return this._tbl.NumAttribValues.toArray; }
        }

        /// <summary>
        /// Get value matrix for key figures; columns match ordered fields
        /// </summary>
        public double[,] KeyFigValues
        {
            get { return this._tbl.KeyFigValues.toArray; }
        }

        /// <summary>
        /// Return true if table has unique attribute rows; no recurring rows with same attribute value combinations
        /// </summary>
        public bool IsUniqueAttributeRows
        {
            get { return this._tbl.IsUniqueAttributeRows; }
        }

        #endregion "MatrixTableX properties"

    #region "MatrixTableX methods"

    /// <summary>
    /// Call all methods of class MatrixTable by name.
    /// With ParameterTypes to resolve ambiguous method names.
    /// </summary>
    /// <param name="MethodName">Method's name</param>
    /// <param name="ParameterTypeNames">Names of parameter types, like 'System.Double' or 'MatrixTable' (set to null if not required)</param>
    /// <param name="Parameters">Parameters of method in correct order</param>
    /// <returns>Outputs of method; null if method's return type is void</returns>
    /// <remarks>
    /// See: MethodBase.Invoke Method (Object, Object[])
    /// http://msdn.microsoft.com/en-us/library/a89hcwhh%28v=vs.110%29.aspx
    /// Set ParameterTypes to null if parameter types (signature) 
    /// are not required to resolve method name ambiguities.
    /// </remarks>
    public object CallMethodByName(string MethodName, string[] ParameterTypeNames,
        params object[] Parameters)
    {
        return ExcelFunc_NO.CallAnyMethodByName("MatrixTable", this.MatrixTableObj, MethodName, 
            ParameterTypeNames, Parameters);
    }

    /// <summary>
    /// Format table data as a printable string.
    /// Overwrites standard object method.
    /// </summary>
    public override string ToString()
    {
        return this._tbl.ToString();
    }

    /// <summary>
    /// Clone table
    /// </summary>
    public MatrixTableX Clone()
    {
        return new MatrixTableX(this._tbl.Clone());
    }

    /// <summary>
    /// Get column names of an excel table
    /// </summary>
    /// <param name="xtbl">Excel table of type ListObject</param>
    /// <returns>Column names</returns>
    public string[] GetColumnNames(Excel.ListObject xtbl)
    {
        if (xtbl == null) return null;
        int cols = xtbl.ListColumns.Count;  // # columns

        if (xtbl.ListColumns == null || cols == 0) return null;

        var ColumnNames = new List<string>();

        for (int i = 0; i < cols; i++)
        {
            // ColumnNames.Add(xtbl.ListColumns.Item[i + 1].Name);
            ColumnNames.Add(xtbl.ListColumns[i + 1].Name);
        }
        return ColumnNames.ToArray();
    }

    /// <summary>
    /// Get value of table element by column name and row index
    /// </summary>
    /// <param name="xtbl">Excel table of type ListObject</param>
    /// <param name="ColName">Column name</param>
    /// <param name="RowInd">Row index</param>
    /// <returns>Value of table element</returns>
    public object GetElementValueByColName(Excel.ListObject xtbl, string ColName, int RowInd)
    {
        // check input parameters
        if (xtbl == null)
            throw new Exception("MatrixTableX.GetTableElementByColName: Null-valued input table!");

        if (ColName == null || ColName == "")
            throw new Exception("MatrixTableX.GetTableElementByColName: Null or empty string for column name!");

        if (RowInd < 0 || RowInd > xtbl.ListRows.Count)
            throw new Exception("MatrixTableX.GetTableElementByColName: Row index is either negative, or out of bounds!");

        if (xtbl.HeaderRowRange == null)
            throw new Exception("MatrixTableX.GetTableElementByColName: Excel table must have a header with column names!");

        // if ( xtbl.ListColumns.get_Item(ColName) == null)

        if (xtbl.ListColumns[ColName] == null)
            throw new Exception("MatrixTableX.GetTableElementByColName: A field named " + ColName + " is not found in table!");

        // input checks OK
        try
        {
            // return ((Excel.Range)xtbl.Range.get_Item(RowInd, xtbl.ListColumns.get_Item(ColName).Index)).Value2;
            return (Excel.Range)xtbl.Range[RowInd, xtbl.ListColumns[ColName].Index].Value2;
        }
        catch (Exception ex)
        {
            throw new Exception("MatrixTableXL.GetTableElementByColName: " + ex.Message);
        }
    }

    /// <summary>
    /// Get value of table element by row and column index
    /// </summary>
    /// <param name="xtbl">Excel table of type ListObject</param>
    /// <param name="RowInd">Row index</param>
    /// <param name="ColInd">Column index</param>
    /// <returns>Value of table element</returns>
    public object GetElementValueByColInd(Excel.ListObject xtbl, int RowInd, int ColInd)
    {
        // check input parameters
        if (xtbl == null)
            throw new Exception("MatrixTableX.GetElementValueByColInd: Null-valued input table!");

        if (ColInd < 0 || ColInd > xtbl.ListColumns.Count)
            throw new Exception("MatrixTableX.GetElementValueByColInd: Column index is either negative, or out of bounds!");

        if (RowInd < 0 || RowInd > xtbl.ListRows.Count)
            throw new Exception("MatrixTableX.GetElementValueByColInd: Row index is either negative, or out of bounds!");

        if (xtbl.HeaderRowRange == null)
            throw new Exception("MatrixTableX.GetElementValueByColInd: Excel table must have a header with column names!");

        // input checks OK
        try
        {
            // return ((Excel.Range)xtbl.Range.get_Item(RowInd, ColInd)).Value2;
            return (Excel.Range)xtbl.Range[RowInd, ColInd].Value2;
        }
        catch (Exception ex)
        {
            throw new Exception("MatrixTableX.GetElementValueByColInd: " + ex.Message);
        }
    }

    /// <summary>
    /// Convert an excel table (ListObject) to a MatrixTableX
    /// </summary>
    /// <param name="xtbl">Excel table</param>
    /// <param name="mdx">Meta data</param>
    /// <param name="TextReplaceNull">Replacement integer value for null in excel table</param>
    /// <param name="NumReplaceNull">Replacement integer value for null in excel table</param>
    /// <param name="KeyFigReplaceNull">Replacement floating value for null in excel table</param>
    public void ReadFromExcelTable(Excel.ListObject xtbl, MetaDataX mdx,
        string TextReplaceNull = "NULL", int NumReplaceNull = 0, double KeyFigReplaceNull = 0)
    {
        // check input parameters

        if (mdx == null )
            throw new Exception("MatrixTableX.ReadFromExcelTable: Null-valued meta data object!");

        // call primary method
        ReadFromExcelTable(xtbl, mdx.metaData, TextReplaceNull, NumReplaceNull, KeyFigReplaceNull);
    }

    /// <summary>
    /// Convert an excel table (ListObject) to a MatrixTableX
    /// </summary>
    /// <param name="xtbl">Excel table</param>
    /// <param name="md">Meta data</param>
    /// <param name="TextReplaceNull">Replacement integer value for null in excel table</param>
    /// <param name="NumReplaceNull">Replacement integer value for null in excel table</param>
    /// <param name="KeyFigReplaceNull">Replacement floating value for null in excel table</param>
    internal void ReadFromExcelTable(Excel.ListObject xtbl, MetaData md,
        string TextReplaceNull = "NULL", int NumReplaceNull = 0, double KeyFigReplaceNull = 0)
    {
        // check input parameters
        if (xtbl == null)
            throw new Exception("MatrixTableX.ReadFromExcelTable: Null-valued input table!");

        if (md == null || md.IsEmpty)
            throw new Exception("MatrixTableX.ReadFromExcelTable: Null-valued or empty meta data object!");

        if (xtbl.HeaderRowRange == null)
            throw new Exception("MatrixTableX.ReadFromExcelTable: Excel table must have a header with column names!");

        if (xtbl.ListColumns == null || xtbl.ListColumns.Count == 0)
            throw new Exception("MatrixTableX.ReadFromExcelTable: Empty table without any columns!");

        try
        {
            // read excel range into MatrixTable
            this._tbl = ExcelFunc_NO.ReadMatrixTableFromRange(xtbl.Range, md);
        }
        catch (Exception ex)
        {
            throw new Exception("MatrixTableX.ReadFromExcelTable: " + ex.Message);
        }
    }

    /// <summary>
    /// Convert an excel table (ListObject) to a MatrixTableX
    /// </summary>
    /// <param name="wbook">Excel Workbook</param>
    /// <param name="xTblName">Name of ListObject (excel table object)</param>
    /// <param name="mdx">Meta data</param>
    /// <param name="TextReplaceNull">Replacement integer value for null in excel table</param>
    /// <param name="NumReplaceNull">Replacement integer value for null in excel table</param>
    /// <param name="KeyFigReplaceNull">Replacement floating value for null in excel table</param>
    public void ReadFromExcelTable2(Excel.Workbook wbook, string xTblName, MetaDataX mdx,
        string TextReplaceNull = "NULL", int NumReplaceNull = 0, double KeyFigReplaceNull = 0)
    {
        if (xTblName == null && xTblName == "")
            throw new Exception("MetaDataX.ReadFromExcelTable2: Null or empty string xTblName!\n");
        try
        {
            // Find excel table (ListObject) by its name and return its ListObject instance
            Excel.ListObject xTblMeta = ExcelFunc_NO.GetListObject(wbook, xTblName);

            // read table data into MatrixTableX
            ReadFromExcelTable(xTblMeta, mdx, TextReplaceNull, NumReplaceNull, KeyFigReplaceNull);
        }
        catch (Exception ex)
        {
            throw new Exception("MetaDataX.ReadFromExcelTable2: " + ex.Message);
        }
    }

    /// <summary>
    /// Write MatrixTableX into an Excel Table (ListObject).
    /// </summary>
    /// <param name="WSheet">Excel Workbook</param>
    /// <param name="xTableName">Name of Excel Table</param>
    /// <param name="CellStr">Upper-left corner of Excel Table in worksheet</param>
    /// <param name="ClearSheetContent">If true, clear whole sheet </param>
    /// <returns>Excel table (ListObject)</returns>
    /// <remarks>
    /// If it doesn't exist already, a new worksheet with the given name will be created.
    /// </remarks>
    public Excel.ListObject WriteToExcelTable(Excel.Worksheet WSheet, string xTableName, 
        string CellStr = "A1", bool ClearSheetContent = false)
    {

        // PARAMETER CHECKS
        if (WSheet == null)
            throw new Exception("MatrixTableX.WriteToExcelTable: Null-valued Worksheet object!");

        if (xTableName == null || xTableName == "")
            throw new Exception("MatrixTableX.WriteToExcelTable: Null or empty string TableName");

        // parameter checks OK..
        return ExcelFunc_NO.WriteTableToExcel(WSheet, this._tbl, xTableName,
                CellStr, ClearSheetContent);
    }

    /// <summary>
    /// Write MatrixTableX into an Excel Table (ListObject).
    /// </summary>
    /// <param name="wbook">Excel Workbook object</param>
    /// <param name="TargetSheetName">Name of target sheet</param>
    /// <param name="xTableName">Name of Excel Table</param>
    /// <param name="CellStr">Upper-left corner of Excel Table in worksheet</param>
    /// <param name="ClearSheetContent">If true, clear whole sheet </param>
    /// <returns>Excel table (ListObject)</returns>
    /// <remarks>
    /// If it doesn't exist already, a new worksheet with the given name will be created.
    /// </remarks>
    public Excel.ListObject WriteToExcelTable2(Excel.Workbook wbook, string TargetSheetName, string xTableName,
        string CellStr = "A1", bool ClearSheetContent = false)
    {
        if (TargetSheetName == null || TargetSheetName == "")
            throw new Exception("MetaDataX.WriteToExcelTable2: Null or empty string TargetSheetName!\n");
        try
        {
            return ExcelFunc_NO.WriteTableToExcel(wbook, this._tbl, TargetSheetName, xTableName,
                CellStr, ClearSheetContent);
        }
        catch (Exception ex)
        {
            throw new Exception("MetaDataX.WriteToExcelTable2: " + ex.Message);
        }
    }

    /// <summary>
    /// View table in GridViewer
    /// </summary>
    /// <param name="HeaderStr">Text displayed at the header of viewer</param>
    public void ViewTable(string HeaderStr)
    {
        if (this.IsEmpty)
            throw new Exception("MatrixTableX.ViewTable: Empty table without any fields!");

        MatrixTable.View_MatrixTable(this._tbl, HeaderStr);
    }

    /// <summary>
    /// Check equality of two tables with given tolerance for key figures
    /// </summary>
    /// <param name="tbl2">Second table</param>
    /// <param name="KeyFigTolerance">Key figure tolerance</param>
    /// <param name="IgnoreFieldOrder">If true, field order of tables need not be same for equality</param>
    /// <param name="IgnoreRowOrder">If true, row order of tables need not be same for equality</param>
    /// <returns>True if tables are equal</returns>
    public bool IsEqual(MatrixTableX tbl2, double KeyFigTolerance = 0.000001, bool IgnoreFieldOrder = true,
        bool IgnoreRowOrder = true)
    {
        return this._tbl.IsEqual(tbl2._tbl, KeyFigTolerance, IgnoreFieldOrder, IgnoreRowOrder);
    }

    /// <summary>
    /// Get field type: 
    /// TextAttribute, IntegerAttribute, DateAttribute, KeyFigure or Undefined
    /// </summary>
    /// <param name="FieldName">Field name</param>
    /// <returns>Field type</returns>
    public string GetFieldType(string FieldName)
    {
        return this.GetFieldType(FieldName).ToString();
    }

    /// <summary>
    /// Return true if two tables have exactly the same fields and field types
    /// </summary>
    /// <param name="tbl2">Table 2</param>
    /// <param name="IgnoreFieldOrder">If true, ignore field order for checking field equality</param>
    /// <returns>True if two tables have the same set of fields</returns>
    public bool IfIdenticalFields(MatrixTableX tbl2, bool IgnoreFieldOrder = true)
    {
        return MatrixTable.IfIdenticalFields(this._tbl, tbl2._tbl, IgnoreFieldOrder);
    }

    /// <summary>
    /// Return true if given field exists as a column in table
    /// </summary>
    /// <param name="FieldName">Field name</param>
    /// <returns>True if field exists</returns>
    public bool IfFieldExists(string FieldName)
    {
        return MatrixTable.IfFieldExistsInTable(this._tbl, FieldName);
    }

    /// <summary>
    /// Get value of a field with given row index
    /// </summary>
    /// <param name="FieldName">Name of field</param>
    /// <param name="RowInd">Row position index</param>
    /// <returns>Field value</returns>
    public Object GetFieldValue(string FieldName, int RowInd)
    {
        return this._tbl.GetFieldValue(FieldName, RowInd);
    }

    /// <summary>
    /// Set value of a field with given row index
    /// </summary>
    /// <param name="FieldName">Name of field</param>
    /// <param name="RowInd">Row position index</param>
    /// <param name="FieldValue">Field value</param>
    public void SetFieldValue(string FieldName, int RowInd, object FieldValue)
    {
        this._tbl.SetFieldValue(FieldName, RowInd, FieldValue);
    }

    /// <summary>
    /// Get subtable with selected columns
    /// </summary>
    /// <param name="ColumnNames">Names of selected columns</param>
    /// <returns>Subtable with selected columns</returns>
    public MatrixTableX SelectColumns(string[] ColumnNames)
    {
        return new MatrixTableX(this._tbl.PartitionColumn(TextVector.CreateVectorWithElements(ColumnNames)));
    }

    /// <summary>
    /// Get subtable with excluded columns
    /// </summary>
    /// <param name="ColumnNames">Names of columns to be excluded</param>
    /// <returns>Subtable with excluded columns</returns>
    public MatrixTableX ExcludeColumns(string[] ColumnNames)
    {
        return new MatrixTableX(this._tbl.ExcludeColumns(TextVector.CreateVectorWithElements(ColumnNames)));
    }

    /// <summary>
    /// Get subtable with selected rows
    /// </summary>
    /// <param name="RowIndices">Row indices (begin from 0)</param>
    /// <returns>Subtable with selected rows</returns>
    public MatrixTableX SelectRows(int[] RowIndices)
    {
        return new MatrixTableX(this._tbl.PartitionRow(NumVector.CreateVectorWithElements(RowIndices)));
    }

    /// <summary>
    /// Append table vertically; both tables must have identical columns
    /// </summary>
    /// <param name="tbl2">Table to be appended</param>
    /// <returns>Table with rows of tbl1 (this instance) and tbl2</returns>
    public MatrixTableX AppendRows(MatrixTableX tbl2)
    {
        return new MatrixTableX(MatrixTable.AppendRowsToTable(this._tbl, tbl2._tbl));
    }

    /// <summary>
    /// Append tables horizontally; tables must have distinct fields
    /// with identical number of rows.
    /// </summary>
    /// <param name="tbl2">Table to be appended</param>
    /// <returns>Table with columns of tbl1 (this instance) and tbl2</returns>
    public MatrixTableX AppendColumns(MatrixTableX tbl2)
    {
        return new MatrixTableX(MatrixTable.AppendColumnsToTable(this._tbl, tbl2._tbl));
    }

    /// <summary>
    /// Sort rows of table w.r.t. given fields and sort options (ASC/DESC).
    /// Sort after all attributes (in order and ASC) if SortStr is null or empty string.
    /// </summary>
    /// <param name="SortStr">Comma separated field names and sort option, like "field1 ASC, field2 DESC"</param>
    /// <returns>Sorted table</returns>
    public MatrixTableX SortRows(string SortStr = null)
    {
        return new MatrixTableX(this._tbl.SortRows(SortStr));
    }

    /// <summary>
    /// Reshuffle rows of table N times by swapping randomly selected row pairs
    /// </summary>
    /// <param name="N">Swap N times</param>
    /// <param name="nRandomSeed">Seed number for replicating pseudo-random numbers</param>
    /// <returns>Table with reshuffled rows</returns>
    public MatrixTableX ShuffleRows(int N, int nRandomSeed = 100)
    {
        return new MatrixTableX(MatrixTable.ShuffleRows(this._tbl, N, nRandomSeed));
    }

    /// <summary>
    /// Filter input table (this instance) with a condition table CondTbl.
    /// if ExcludeMatchedRows = false (default case) output table includes only rows that match with
    /// some row(s) of condition table CondTbl.
    /// </summary>
    /// <param name="CondTbl">Condition table</param>
    /// <param name="ExcludeMatchedRows">If true, exclude matched rows from input table</param>
    /// <param name="JokerMatchesAllvalues">If true, JokerValue matches all other values</param>
    /// <param name="TextJoker">Joker (match-all) value for text attributes</param>
    /// <param name="NumJoker">Joker (match-all) value for numeric attributes</param>
    /// <returns>Filtered table</returns>
    public MatrixTableX FilterTableA(MatrixTableX CondTbl,
        bool ExcludeMatchedRows = false, bool JokerMatchesAllvalues = false,
        string TextJoker = "ALL", int NumJoker = 0)
    {
        return new MatrixTableX(MatrixTable.FilterTableA(this._tbl, CondTbl._tbl,
            ExcludeMatchedRows, JokerMatchesAllvalues, TextJoker, NumJoker));
    }

    /// <summary>
    /// Aggregate all key figures with the same aggregation function.
    /// </summary>
    /// <param name="AggrOpt">Aggregation option (sum, min, max, avg)</param>
    /// <returns>Aggregated table with distinct rows and aggregated key figures</returns>
    /// <remarks>Applies default option "sum" if an invalid option like "avt" is entered.</remarks>
    public MatrixTableX AggregateAllKeyFigures(string AggrOpt = "sum")
    {
        string AggrOptS = AggrOpt.ToLower();
        AggregateOption agrop = AggregateOption.nSum;

        switch (AggrOptS)
        {
            case "sum":
                agrop = AggregateOption.nSum;
                break;
            case "min":
                agrop = AggregateOption.nMin;
                break;
            case "max":
                agrop = AggregateOption.nMax;
                break;
            case "avg":
                agrop = AggregateOption.nAvg;
                break;
            default:
                break;
        }
        return new MatrixTableX(MatrixTable.AggregateAllKeyFigures(this._tbl, null, agrop));
    }

    /// <summary>
    /// Add all key figures of table2 to common key figures of table1 (this instance) with matching attribute rows.
    /// </summary>
    /// <param name="tbl2">input table 2</param>
    /// <param name="AddToRest">Add this amount to rows of table1 that have no match in table2</param>
    /// <param name="JokerMatchesAllvalues">If true, JokerValue matches all other values</param>
    /// <param name="TextJoker">Joker (match-all) value for text attributes</param>
    /// <param name="NumJoker">Joker (match-all) value for numeric attributes</param>
    /// <returns>Output table with added key figures</returns>
    /// <remarks>
    /// All fields of table2 must be contained by table1; i.e. fields of table2 must be a subset of fields of table1.
    /// </remarks>
    public MatrixTableX AddAllKeyFigures(MatrixTableX tbl2, double AddToRest = 0.0,
        bool JokerMatchesAllvalues = false, string TextJoker = "ALL", int NumJoker = 0)
    {
        return new MatrixTableX(MatrixTable.AddAllKeyFigures(this._tbl, tbl2._tbl, AddToRest,
            JokerMatchesAllvalues, TextJoker, NumJoker));
    }

    /// <summary>
    /// Add selected key figure of table2 to selected key figure of table1 (this instance) with matching attribute rows
    /// </summary>
    /// <param name="tbl2">input table 2</param>
    /// <param name="InputKeyFigTbl1">Selected key figure of table1</param>
    /// <param name="InputKeyFigTbl2">Selected key figure of table2</param>
    /// <param name="OutputKeyFig">Resultant key figure</param>
    /// <param name="AddToRest">Add this amount to rows of table1 that have no match in table2</param>
    /// <param name="JokerMatchesAllvalues">If true, JokerValue matches all other values</param>
    /// <param name="TextJoker">Joker (match-all) value for text attributes</param>
    /// <param name="NumJoker">Joker (match-all) value for numeric attributes</param>
    /// <returns>Output table with added key figures</returns>
    public MatrixTableX AddSelectedKeyFigures(MatrixTableX tbl2,
        string InputKeyFigTbl1, string InputKeyFigTbl2, string OutputKeyFig, double AddToRest = 0,
        bool JokerMatchesAllvalues = false, string TextJoker = "ALL", int NumJoker = 0)
    {
        return new MatrixTableX(MatrixTable.AddSelectedKeyFigures(this._tbl, tbl2._tbl,
            InputKeyFigTbl1, InputKeyFigTbl2, OutputKeyFig, AddToRest,
            JokerMatchesAllvalues, TextJoker, NumJoker));
    }

    /// <summary>
    /// Multiply all key figures of table1 (this instance) by common key figures of table2 with matching attribute rows.
    /// </summary>
    /// <param name="tbl1">input table 1</param>
    /// <param name="tbl2">input table 2</param>
    /// <param name="MultiplyRestWith">Multiply with this amount the rows of table1 that have no match in table2</param>
    /// <param name="JokerMatchesAllvalues">If true, JokerValue matches all other values</param>
    /// <param name="TextJoker">Joker (match-all) value for text attributes</param>
    /// <param name="NumJoker">Joker (match-all) value for numeric attributes</param>
    /// <returns>Output table with multiplied key figures</returns>
    /// <remarks>
    /// All fields of table2 must be contained by table1; i.e. fields of table2 must be a subset of fields of table1
    /// </remarks>
    public MatrixTableX MultiplyAllKeyFigures(MatrixTableX tbl2, double MultiplyRestWith = 1.0,
        bool JokerMatchesAllvalues = false, string TextJoker = "ALL", int NumJoker = 0)
    {
        return new MatrixTableX(MatrixTable.MultiplyAllKeyFigures(this._tbl, tbl2._tbl, MultiplyRestWith,
            JokerMatchesAllvalues, TextJoker, NumJoker));
    }

    /// <summary>
    /// Multiply selected key figure of table1 (this instance) with selected key figure of table2 with matching attribute rows
    /// </summary>
    /// <param name="tbl2">input table 2</param>
    /// <param name="InputKeyFigTbl1">Selected key figure of table1</param>
    /// <param name="InputKeyFigTbl2">Selected key figure of table2</param>
    /// <param name="OutputKeyFig">Resultant key figure</param>
    /// <param name="MultiplyRestWith">Multiply this amount with rows of table1 that have no match in table2</param>
    /// <param name="JokerMatchesAllvalues">If true, JokerValue matches all other values</param>
    /// <param name="TextJoker">Joker (match-all) value for text attributes</param>
    /// <param name="NumJoker">Joker (match-all) value for numeric attributes</param>
    /// <returns>Output table with multiplied key figures</returns>
    public MatrixTableX MultiplySelectedKeyFigures(MatrixTableX tbl2,
        string InputKeyFigTbl1, string InputKeyFigTbl2, string OutputKeyFig, double MultiplyRestWith = 0,
        bool JokerMatchesAllvalues = false, string TextJoker = "ALL", int NumJoker = 0)
    {
        return new MatrixTableX(MatrixTable.MultiplySelectedKeyFigures(this._tbl, tbl2._tbl,
            InputKeyFigTbl1, InputKeyFigTbl2, OutputKeyFig, MultiplyRestWith,
            JokerMatchesAllvalues, TextJoker, NumJoker));
    }

    /// <summary>
    /// Combine fields of two tables by matching common attributes. 
    /// Output table contains only the matched rows of Table1 (this instance), plus added fields from Table2.
    /// </summary>
    /// <param name="tbl2">Input table 2</param>
    /// <param name="JokerMatchesAllvalues">If true, JokerValue matches all other values</param>
    /// <param name="TextJoker">Joker (match-all) value for text attributes</param>
    /// <param name="NumJoker">Joker (match-all) value for numeric attributes</param>
    /// <returns>Output table containing only the matched rows of Table1, plus added fields from Table2</returns>
    public MatrixTableX CombineTables(MatrixTableX tbl2,
        bool JokerMatchesAllvalues = false, string TextJoker = "ALL", int NumJoker = 0)
    {
        return new MatrixTableX(MatrixTable.CombineTables(this._tbl, tbl2._tbl,
            JokerMatchesAllvalues, TextJoker, NumJoker));
    }

    /// <summary>
    /// Subtract all key figures of table2 from common key figures of table1 (this instance) with matching attribute rows.
    /// </summary>
    /// <param name="tbl2">input table 2</param>
    /// <param name="SubtractFromRest">Subtract this amount from rows of table1 that have no match in table2</param>
    /// <param name="JokerMatchesAllvalues">If true, JokerValue matches all other values</param>
    /// <param name="TextJoker">Joker (match-all) value for text attributes</param>
    /// <param name="NumJoker">Joker (match-all) value for numeric attributes</param>
    /// <returns>Output table with added key figures</returns>
    /// <remarks>
    /// All fields of table2 must be contained by table1; i.e. fields of table2 must be a subset of fields of table1.
    /// </remarks>
    private MatrixTableX SubtractAllKeyFigures(MatrixTableX tbl2, double SubtractFromRest = 0.0,
        bool JokerMatchesAllvalues = false, string TextJoker = "ALL", int NumJoker = 0)
    {
        return new MatrixTableX(MatrixTable.SubtractAllKeyFigures(this._tbl, tbl2._tbl, SubtractFromRest,
            JokerMatchesAllvalues, TextJoker, NumJoker));
    }

    // test function for workbook
    public string GetSheets(Excel.Workbook WBook)
    {
        if (WBook == null)
            throw new Exception("MatrixTableX.GetSheets: Null-valued Workbook object!");
            
        Excel.Application xlapp = null;

        // get workbook's application instance
        try
        {
            xlapp = WBook.Application;
        }
        catch(Exception ex) 
        {
            throw new Exception("MatrixTableX.GetSheets: xlapp " + ex.Message);
        }

        string str = "";

        try
        {
            // get all existing worksheets
            foreach (var wsheet in WBook.Worksheets)
            {
                str += ((Excel.Worksheet)wsheet).Name + ",";
            }
            return str;
        }
        catch (Exception ex)
        {
            throw new Exception("MatrixTableX.GetSheets: Sheets " + ex.Message);
        }
    }

    #endregion "MatrixTableX methods"

    #region "OLD and Obsolete Methods"

        /// <summary>
        /// Convert an excel table (ListObject) to a MatrixTableX
        /// </summary>
        /// <param name="xtbl">Excel table</param>
        /// <param name="md">Meta data</param>
        /// <param name="TextReplaceNull">Replacement integer value for null in excel table</param>
        /// <param name="NumReplaceNull">Replacement integer value for null in excel table</param>
        /// <param name="KeyFigReplaceNull">Replacement floating value for null in excel table</param>
        [Obsolete]
        private void FromExcelTable_OLD(Excel.ListObject xtbl, MetaData md,
            string TextReplaceNull = "NULL", int NumReplaceNull = 0, double KeyFigReplaceNull = 0)
        {
            // check input parameters
            if (xtbl == null)
                throw new Exception("MatrixTableX.ConvertToMatrixTableX: Null-valued input table!");

            if (md == null || md.IsEmpty)
                throw new Exception("MatrixTableX.ConvertToMatrixTableX: Null-valued or empty meta data object!");

            if (xtbl.HeaderRowRange == null)
                throw new Exception("MatrixTableX.ConvertToMatrixTableX: Excel table must have a header with column names!");

            if (xtbl.ListColumns == null || xtbl.ListColumns.Count == 0)
                throw new Exception("MatrixTableX.ConvertToMatrixTableX: Empty table without any columns!");

            // init variables
            var tfields = TableFields.CreateEmptyTableFields(md);
            FieldType ftype;
            string fieldname;
            var FieldList = new List<string>();

            try
            {

                // add fields to table structure
                for (int i = 1; i <= xtbl.ListColumns.Count; i++)
                {
                    // check master data
                    fieldname = xtbl.ListColumns[i].Name.ToLower();
                    ftype = MetaData.GetFieldType(md, fieldname);

                    // check if there is a recurring field name
                    if (FieldList.Contains(fieldname))
                        throw new Exception("MatrixTableX.ConvertToMatrixTableX: Recurring field name! " + fieldname + "\n");
                    else
                        FieldList.Add(fieldname);

                    // add fields
                    switch (ftype)
                    {
                        case FieldType.TextAttribute:
                        case FieldType.IntegerAttribute:
                        case FieldType.DateAttribute:
                        case FieldType.KeyFigure:

                            TableFields.AddNewField(tfields, fieldname);
                            break;

                        case FieldType.Undefined:
                            throw new Exception("MatrixTableX.ConvertToMatrixTableX: Excel table field " + fieldname +
                                " is not defined in meta data!\n");
                    }
                }   // for i

                // init variables
                var tblout = MatrixTable.CreateEmptyTable(tfields);     // table without rows
                string sval = "EMPTY";
                int ival = 0;
                double dval = 0.0;
                TableRow trow;
                object val;     // element value


                // add rows
                for (int i = 2; i <= xtbl.ListRows.Count + 1; i++)
                {
                    trow = TableRow.CreateDefaultTableRow(tfields);

                    for (int j = 1; j <= xtbl.ListColumns.Count; j++)
                    {
                        // check master data
                        fieldname = xtbl.ListColumns[j].Name.ToLower();
                        ftype = MetaData.GetFieldType(md, fieldname);
                        val = ((Excel.Range)xtbl.Range[i, j]).Value2;

                        switch (ftype)
                        {
                            case FieldType.TextAttribute:
                                if (val == null)
                                    sval = TextReplaceNull;
                                else
                                    sval = val.ToString();
                                TableRow.SetTextAttributeValue(trow, fieldname, sval);
                                break;

                            case FieldType.IntegerAttribute:
                            case FieldType.DateAttribute:

                                if (val == null)
                                    ival = NumReplaceNull;
                                else
                                {
                                    // throw an error if element value can't be converted to integer
                                    try
                                    {
                                        ival = Convert.ToInt32(val.ToString());
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("Value " + val + " cannot be converted to integer! " +
                                            "Row: " + i + " Column: " + fieldname + "\n" + ex.Message);
                                    }
                                    TableRow.SetNumAttributeValue(trow, fieldname, ival);
                                }
                                break;

                            case FieldType.KeyFigure:

                                if (val == null)
                                    dval = KeyFigReplaceNull;
                                else
                                {
                                    // throw an error if element value can't be converted to double
                                    try
                                    {
                                        dval = Convert.ToDouble(val.ToString());
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("Value " + val + " cannot be converted to double! " +
                                            "Row: " + i + " Column: " + fieldname + "\n" + ex.Message);
                                    }
                                    TableRow.SetKeyFigureValue(trow, fieldname, dval);
                                }

                                break;
                        }
                    }   // for j
                    tblout = MatrixTable.AddRowToTable(tblout, trow);
                }

                // assign field value
                this._tbl = tblout;
            }
            catch (Exception ex)
            {
                throw new Exception("MatrixTableX.ConvertToMatrixTableX: \n" + ex.Message);
            }
        }

        /// <summary>
        /// Convert a MatrixTableX to an Excel Table (ListObject).
        /// </summary>
        /// <param name="WBook">Excel Workbook</param>
        /// <param name="SheetName">Name of worksheet</param>
        /// <param name="xTableName">Name of Excel Table</param>
        /// <param name="CellStr">Upper-left corner of Excel Table in worksheet</param>
        /// <param name="ClearSheetContent">If true, clear whole sheet </param>
        /// <returns>Excel table (ListObject)</returns>
        /// <remarks>
        /// If it doesn't exist already, a new worksheet with the given name will be created.
        /// </remarks>
        [Obsolete]
        private Excel.ListObject ToExcelTable_OLD(Excel.Workbook WBook, string SheetName,
            string xTableName, string CellStr = "A1", bool ClearSheetContent = false)
        {
            // How to: Add ListObject Controls to Worksheets
            // http://msdn.microsoft.com/en-us/library/eyfs6478.aspx
            // Binding a ListObject to a .NET List
            // http://www.clear-lines.com/blog/post/Binding-a-ListObject-to-a-NET-List.aspx
            // Binding a ListObject to a DataTable: ListObject.SetDataBinding Method (Object)
            // http://msdn.microsoft.com/en-us/library/c5f64c2x.aspx
            // How to: Access Office Interop Objects by Using Visual C# 2010 Features
            // http://msdn.microsoft.com/en-us/library/dd264733.aspx

            // PARAMETER CHECKS
            if (WBook == null)
                throw new Exception("MatrixTableX.ExportToExcelTable: Null-valued Workbook object!");

            if (SheetName == null || SheetName == "")
                throw new Exception("MatrixTableX.ExportToExcelTable: Null or empty string SheetName");

            // parameter checks OK..
            Excel.Application xlapp = null;
            Excel.Worksheet wsheet = null;
            Excel.Range range = null;

            try
            {
                // get application instance
                xlapp = WBook.Application;

                // check if sheet exists
                bool SheetExists = false;

                foreach (Excel.Worksheet sheet in WBook.Sheets)
                {
                    if (sheet.Name.Equals(SheetName))
                    {
                        SheetExists = true;
                        break;
                    }
                }

                // create a new sheet if it does not already exist
                if (SheetExists)
                    wsheet = (Excel.Worksheet)WBook.Sheets[SheetName];
                else
                {
                    wsheet = (Excel.Worksheet)WBook.Sheets.Add();
                    wsheet.Name = SheetName;
                }

                if (ClearSheetContent)
                    wsheet.Cells.ClearContents();  // clear sheet content

                // get upper left corner of range defined by CellStr
                range = (Excel.Range)wsheet.Range(CellStr).Cells[1, 1];   // .get_Range(1, 1);

                // Write table to sheet
                ExcelFunc_NO.WriteTableToExcel(wsheet, this._tbl, range.Address);

                // derive range for table, +1 row for table header
                range = wsheet.Range(CellStr).Resize(this.RowCount + 1, this.ColumnCount);

                // add ListObject to sheet

                // ListObjects.AddEx Method 
                // http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.listobjects.addex%28v=office.14%29.aspx

                /*
                Excel.ListObject tbl = (Excel.ListObject)wsheet.ListObjects.Add(
	                sourceType: XlListObjectSourceType.xlSrcRange,
	                source: range, 
	                linkSource: Type.Missing,
	                xlListObjectHasHeaders: XlYesNoGuess.xlYes); // ?

                tbl.Name = TableName;
                 * */
                Excel.ListObject tbl = ExcelFunc_NO.AddListObject(range, xTableName);

                // return excel table (ListObject)
                return (Excel.ListObject)tbl;
            }
            catch (Exception ex)
            {
                throw new Exception("MatrixTableX.ExportToExcelTable: " + ex.Message);
            }
            finally
            {
                try
                {
                    ExcelFunc_NO.releaseObject(wsheet);
                    ExcelFunc_NO.releaseObject(WBook);
                    ExcelFunc_NO.releaseObject(xlapp);
                    ExcelFunc_NO.releaseObject(range);
                }
                catch { }
            }
        }

        #endregion "OLD and Obsolete Methods"

    }    

}
