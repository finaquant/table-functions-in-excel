// Finaquant Analytics - http://finaquant.com/
// Copyright © Finaquant Analytics GmbH
// Email: support@finaquant.com

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Text.RegularExpressions;
using FinaquantCalcs;

using System.Windows.Forms;
using System.Data.Odbc;
using System.Data.OleDb;
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

namespace FinaquantInExcel
{

    /// <summary>
    /// Helper methods for excel integration, based on NetOffice (NO) assemblies
    /// </summary>
    public class ExcelFunc_NO
    {
        /// <summary>
        /// Read an excel range into a table of type MatrixTable.
        /// Excel range must have a header row with field names.
        /// </summary>
        /// <param name="range">Excel range with header and table data</param>
        /// <param name="md">Meta data object with field definitions</param>
        /// <param name="TextReplaceNull">Replace null strings in DataTable with this value</param>
        /// <param name="NumReplaceNull">Replace null integers in DataTable with this value</param>
        /// <param name="KeyFigReplaceNull">Replace null numbers in DataTable with this value</param>
        /// <returns>Table of type MatrixTable</returns>
        public static MatrixTable ReadMatrixTableFromRange(Excel.Range range, MetaData md,
            string TextReplaceNull = "", int NumReplaceNull = 0, double KeyFigReplaceNull = 0)
        {
            // PARAMETER CHECKS
            if (range == null)
                throw new Exception("ReadTableFromExcelRange: Null-valued excel range!");
            if (md == null)
                throw new Exception("ReadTableFromExcelRange: Null-valued meta data object.");

            // parameter checks OK ...
            Excel.Worksheet wsheet = null;
            Excel.Range xrange = null;

            try
            {
                // get worksheet from range
                wsheet = (Excel.Worksheet)range.Worksheet;

                // init variables
                var tblfields = TableFields.CreateEmptyTableFields(md);
                object CellValue;
                string fieldname;
                var fieldnames = new List<string>();
                FieldType ftype;

                // read column names from header line into MatrixTable
                for (int i = 0; i < range.Columns.Count; i++)
                {
                    // all field names are stored in small letters in MetaData
                    fieldname = ((string)range.Cells[1, i + 1].Value2).ToLower();
                    fieldnames.Add(fieldname);

                    ftype = MetaData.GetFieldType(md, fieldname);

                    switch (ftype)
                    {
                        case FieldType.TextAttribute:
                        case FieldType.IntegerAttribute:
                        case FieldType.DateAttribute:
                        case FieldType.KeyFigure:

                            // add field to table structure
                            TableFields.AddNewField(tblfields, fieldname);
                            break;

                        case FieldType.Undefined:            // undefined field type

                            // get field type from the first cell under column name
                            FieldType ft = GetFieldTypeOfExcelCell(range.Cells[2, i + 1]);

                            if (ft != FieldType.Undefined)
                                TableFields.AddNewField(tblfields, fieldname, ft);
                            else
                                throw new Exception("Undefined field type! Field '" + fieldname + "' is not defined in MetaData "
                                  + "and its field type could not be determined from the first cell of the column.");

                            break;
                    }
                }

                // init variables
                MatrixTable tbl = MatrixTable.CreateEmptyTable(tblfields);
                string sval;
                int ival = 0;
                double dval = 0.0;
                TableRow trow;

                // read cell values row by row
                for (int i = 0; i < range.Rows.Count - 1; i++)
                {
                    trow = TableRow.CreateDefaultTableRow(tblfields);

                    for (int j = 0; j < range.Columns.Count; j++)
                    {
                        CellValue = (object)range.Cells[2 + i, j + 1].Value2;
                        fieldname = fieldnames[j];

                        ftype = MetaData.GetFieldType(md, fieldname);

                        switch (ftype)
                        {
                            case FieldType.TextAttribute:

                                if (CellValue == null)
                                    sval = TextReplaceNull;
                                else
                                    sval = CellValue.ToString();

                                TableRow.SetTextAttributeValue(trow, fieldname, sval);
                                break;

                            case FieldType.IntegerAttribute:
                            case FieldType.DateAttribute:

                                if (CellValue == null)
                                    ival = NumReplaceNull;
                                else
                                {
                                    // throw an error if element value can't be converted to integer
                                    try
                                    {
                                        ival = Convert.ToInt32(CellValue.ToString());
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("Cell value " + CellValue + " cannot be converted to integer! " +
                                            "Row: " + (i + 1) + " Column: " + fieldname + "\n" + ex.Message);
                                    }
                                }
                                TableRow.SetNumAttributeValue(trow, fieldname, ival);
                                break;

                            case FieldType.KeyFigure:

                                if (CellValue == null)
                                    dval = KeyFigReplaceNull;
                                else
                                { 
                                    // throw an error if element value can't be converted to double
                                    try
                                    {
                                        dval = Convert.ToDouble(CellValue.ToString());
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("Cell value " + CellValue + " cannot be converted to double! " +
                                            "Row: " + i + " Column: " + fieldname + "\n" + ex.Message);
                                    }
                                }
                                TableRow.SetKeyFigureValue(trow, fieldname, dval);
                                break;

                            default:  // should not happen
                                break;
                        }
                    }  // for j
                    tbl = MatrixTable.AddRowToTable(tbl, trow);

                }  // for i
                // return data table
                return tbl;
            }
            catch (Exception ex)
            {
                throw new Exception("ReadTableFromExcelRange: " + ex.Message);
            }
            finally
            {
                try
                {
                    releaseObject(wsheet);
                    releaseObject(xrange);
                }
                catch { }
            }
        }

        /// <summary>
        /// Read a ListObject (excel table) into MatrixTable.
        /// Excel table must have a header row with field names.
        /// </summary>
        /// <param name="xTableName">Name of Excel Table (ListObject)</param>
        /// <param name="md">Meta data object with field definitions</param>
        /// <param name="WorkbookPath">Full name of excel workbook</param>
        /// <param name="TextReplaceNull">Replace null strings in DataTable with this value</param>
        /// <param name="NumReplaceNull">Replace null integers in DataTable with this value</param>
        /// <param name="KeyFigReplaceNull">Replace null numbers in DataTable with this value</param>
        /// <returns>Table of type MatrixTable</returns>
        public static MatrixTable ReadMatrixTableFromExcel(string xTableName, MetaData md, string WorkbookPath = null,
            string TextReplaceNull = "", int NumReplaceNull = 0, double KeyFigReplaceNull = 0)
        {
            // PARAMETER CHECKS
            if (xTableName == null || xTableName == "")
                throw new Exception("ExcelFunc_NO.ReadMatrixTableFromExcel: Null-valued or empty string xTableName!");
            if (md == null)
                throw new Exception("ExcelFunc_NO.ReadMatrixTableFromExcel: Null-valued meta data object.");

            try
            {
                // Get application insance
                Excel.Application xlapp = GetExcelApplicationInstance();

                // Get workbook object
                Excel.Workbook wbook;

                if (WorkbookPath == null)
                    wbook = xlapp.ActiveWorkbook;
                else
                    wbook = GetWorkbookByFullName(xlapp, WorkbookPath);

                // get ListObject (excel table) handle
                Excel.ListObject xTbl = GetListObject(wbook, xTableName);

                if (xTbl == null)
                    throw new Exception("A ListObject (excel table) with the given name '"
                        + xTableName + "' could not be found!\n");

                // read MatrixTable from the range of excel table
                return ReadMatrixTableFromRange(xTbl.Range, md, TextReplaceNull, NumReplaceNull, KeyFigReplaceNull);
            }
            catch (Exception ex)
            {
                throw new Exception("ExcelFunc_NO.ReadMatrixTableFromExcel: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Get field names and types of an excel table
        /// </summary>
        /// <param name="range">Table range</param>
        /// <param name="md">Metadata object</param>
        /// <returns>Table Fields</returns>
        public static TableFields ReadTableFieldsFromRange(Excel.Range range, MetaData md)
        {
            // PARAMETER CHECKS
            if (range == null)
                throw new Exception("ReadTableFieldsFromRange: Null-valued excel range!");
            if (md == null)
                throw new Exception("ReadTableFieldsFromRange: Null-valued meta data object.");

            // parameter checks OK ...
            Excel.Worksheet wsheet = null;
            // Excel.Range xrange = null;

            try
            {
                // get worksheet from range
                wsheet = (Excel.Worksheet)range.Worksheet;

                // init variables
                var tblfields = TableFields.CreateEmptyTableFields(md);
                string fieldname;
                var fieldnames = new List<string>();
                FieldType ftype;

                // read column names from header line into MatrixTable
                for (int i = 0; i < range.Columns.Count; i++)
                {
                    // all field names are stored in small letters in MetaData
                    fieldname = ((string)range.Cells[1, i + 1].Value2).ToLower();
                    fieldnames.Add(fieldname);

                    ftype = MetaData.GetFieldType(md, fieldname);

                    switch (ftype)
                    {
                        case FieldType.TextAttribute:
                        case FieldType.IntegerAttribute:
                        case FieldType.DateAttribute:
                        case FieldType.KeyFigure:

                            // add field to table structure
                            TableFields.AddNewField(tblfields, fieldname);
                            break;

                        case FieldType.Undefined:            // undefined field type

                            // get field type from the first cell under column name
                            FieldType ft = GetFieldTypeOfExcelCell(range.Cells[2, i + 1]);

                            if (ft != FieldType.Undefined)
                                TableFields.AddNewField(tblfields, fieldname, ft);
                            else
                                throw new Exception("Undefined field type! Field '" + fieldname + "' is not defined in MetaData "
                                  + "and its field type could not be determined from the first cell of the column.");

                            break;
                    }
                }
                return tblfields;
            }
            catch (Exception ex)
            {
                throw new Exception("ReadTableFieldsFromRange: " + ex.Message);
            }
            finally
            {
                try
                {
                    releaseObject(wsheet);
                }
                catch { }
            }
        }

        /// <summary>
        /// Get field names and types of an excel table
        /// </summary>
        /// <param name="xTableName">Name of excel table (ListObject)</param>
        /// <param name="md">Metadata object</param>
        /// <param name="WorkbookPath">Full workbook name</param>
        /// <returns>Table Fields</returns>
        public static TableFields ReadTableFieldsFromExcel(string xTableName, MetaData md, string WorkbookPath = null)
        {
            // PARAMETER CHECKS
            if (xTableName == null || xTableName == "")
                throw new Exception("ExcelFunc_NO.ReadTableFieldsFromExcel: Null-valued or empty string xTableName!");
            if (md == null)
                throw new Exception("ExcelFunc_NO.ReadTableFieldsFromExcel: Null-valued meta data object.");

            try
            {
                // Get application insance
                Excel.Application xlapp = GetExcelApplicationInstance();

                // Get workbook object
                Excel.Workbook wbook;

                if (WorkbookPath == null)
                    wbook = xlapp.ActiveWorkbook;
                else
                    wbook = GetWorkbookByFullName(xlapp, WorkbookPath);

                // get ListObject (excel table) handle
                Excel.ListObject xTbl = GetListObject(wbook, xTableName);

                if (xTbl == null)
                    throw new Exception("A ListObject (excel table) with the given name '"
                        + xTableName + "' could not be found!\n");

                // read MatrixTable from the range of excel table
                return ReadTableFieldsFromRange(xTbl.Range, md);
            }
            catch (Exception ex)
            {
                throw new Exception("ExcelFunc_NO.ReadTableFieldsFromExcel: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Write DataTable to excel sheet. Header row contains field names as defined in DataTable.
        /// </summary>
        /// <param name="wsheet">Excel Worksheet object</param>
        /// <param name="tbl">Table data of type DataTable containing field names</param>
        /// <param name="TopLeftCell">Upper-left corner of table in excel sheet</param>
        /// <remarks>
        /// For faster execution, range values are assigned in bulks with arrays.
        /// See: How to speed up dumping a DataTable into an Excel worksheet?
        /// http://stackoverflow.com/questions/2692979/how-to-speed-up-dumping-a-datatable-into-an-excel-worksheet
        /// </remarks>
        public static void WriteTableToExcel(Excel.Worksheet wsheet, DataTable tbl, string TopLeftCell = "A1")
        {
            // PARAMETER CHECKS
            if (wsheet == null )
                throw new Exception("WriteTableToExcelSheet: Null-valued Worksheet object!");
            if (tbl == null || tbl.Columns.Count == 0)
                throw new Exception("WriteTableToExcelSheet: Null or empty DataTable!");
            if (TopLeftCell == null || TopLeftCell == "")
                throw new Exception("WriteTableToExcelSheet: Null or empty string TopLeftCell!");

            Excel.Range range = null;
            Excel.Range UpperLeftCell = null;

            try
            {
                // get upper left cell in range
                UpperLeftCell = (Excel.Range)wsheet.Range(TopLeftCell).Cells[1, 1];

                // clear range for table
                range = wsheet.Range(UpperLeftCell.Address);
                range = range.Resize(tbl.Rows.Count + 1, tbl.Columns.Count);
                range.Cells.ClearContents();

                // write column names as table header (first line)
                string[] fieldNames = new string[tbl.Columns.Count];

                for (int i = 0; i < tbl.Columns.Count; i++)
                {
                    UpperLeftCell.Cells[1, i + 1].Value = tbl.Columns[i].ColumnName;
                    fieldNames[i] = tbl.Columns[i].ColumnName;
                }

                // for each column, create an array and set the array 
                // to the excel range for that column.
                for (int i = 0; i < fieldNames.Length; i++)
                {
                    string[,] clnDataString = new string[tbl.Rows.Count, 1];
                    int[,] clnDataInt = new int[tbl.Rows.Count, 1];
                    double[,] clnDataDouble = new double[tbl.Rows.Count, 1];

                    // row and column offsets (shifts w.r.t. upper left corner)
                    range = wsheet.Range(UpperLeftCell.Address).Offset(1,i);

                    // define range size
                    range = range.Resize(tbl.Rows.Count, 1);

                    string dataTypeName = tbl.Columns[fieldNames[i]].DataType.Name;

                    for (int j = 0; j < tbl.Rows.Count; j++)
                    {
                        switch (dataTypeName.ToLower())
                        {
                            case "int32":
                                clnDataInt[j, 0] = Convert.ToInt32(tbl.Rows[j][fieldNames[i]]);
                                break;

                            case "double":
                                clnDataDouble[j, 0] = Convert.ToDouble(tbl.Rows[j][fieldNames[i]]);
                                break;

                            case "datetime":
                                clnDataInt[j, 0] = Convert.ToInt32(DateFunctions.DateToNumber((DateTime) tbl.Rows[j][fieldNames[i]]));

                                // clnDataString[j, 0] = ((DateTime)tbl.Rows[j][fieldNames[i]]).ToShortDateString();

                                //if (fieldNames[i].ToLower().Contains("time"))
                                //    clnDataString[j, 0] = Convert.ToDateTime(tbl.Rows[j][fieldNames[i]]).ToShortTimeString();
                                //else if (fieldNames[i].ToLower().Contains("date"))
                                //    clnDataString[j, 0] = Convert.ToDateTime(tbl.Rows[j][fieldNames[i]]).ToShortDateString();
                                //else
                                //    clnDataString[j, 0] = Convert.ToDateTime(tbl.Rows[j][fieldNames[i]]).ToString();

                                break;

                            case "string":
                                clnDataString[j, 0] = tbl.Rows[j][fieldNames[i]].ToString();
                                break;

                            default:
                                throw new Exception("Invalid data type! Any data type other than "
                                    + "string, integer, DateTime and double is not compatible with "
                                    + " finaquant's table of type MatrixType");
                        }

                    }
                    // set values in the sheet wholesale.

                    // Excel Interop cell formatting of Dates
                    // http://stackoverflow.com/questions/15603098/excel-interop-cell-formatting-of-dates
                    // Excel Number Format: What is “[$-409]”?
                    // http://stackoverflow.com/questions/894805/excel-number-format-what-is-409
                    // Creating international number formats
                    // http://office.microsoft.com/en-us/excel-help/creating-international-number-formats-HA001034635.aspx
                    // microsoft.interop.excel Formatting cells
                    // http://stackoverflow.com/questions/7401996/microsoft-interop-excel-formatting-cells
                    
                    if (dataTypeName == "Int32")
                    {
                        range.Value = clnDataInt;
                        range.NumberFormat = "General";
                    }
                    else if (dataTypeName == "DateTime")
                    {
                        // range.set_Value(value: clnDataInt);
                        range.Value = clnDataInt;
                        range.NumberFormat = "dd/mm/yyyy";  //"dd-mmm-yy";
                    }

                    else if (dataTypeName == "Double")
                    {
                        // range.set_Value(value: clnDataDouble);
                        range.Value = clnDataDouble;
                        range.NumberFormat = "General";
                    }
                    else
                    {
                        // range.set_Value(value: clnDataString);
                        range.Value = clnDataString;
                        range.NumberFormat = "General";
                    }
                }

                // make the header range bold
                range = wsheet.Range(UpperLeftCell.Address).Resize(1, tbl.Columns.Count);
                range.Font.Bold = true;

                // autofit for better view
                wsheet.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                throw new Exception("WriteTableToExcelSheet: " + ex.Message);
            }
            finally
            {
                releaseObject(range);
                releaseObject(UpperLeftCell);
            }
        }

        /// <summary>
        /// Write MatrixTable to excel sheet. Header row contains field names as defined in MatrixTable.
        /// </summary>
        /// <param name="wsheet">Excel Worksheet object</param>
        /// <param name="tbl">Table data of type MatrixTable</param>
        /// <param name="TopLeftCell">Upper-left corner of table in excel sheet</param>
        /// <remarks>
        /// For faster execution, range values are assigned in bulks with arrays.
        /// See: How to speed up dumping a DataTable into an Excel worksheet?
        /// http://stackoverflow.com/questions/2692979/how-to-speed-up-dumping-a-datatable-into-an-excel-worksheet
        /// </remarks>
        public static void WriteTableToExcel(Excel.Worksheet wsheet, MatrixTable tbl, string TopLeftCell = "A1")
        {
            // PARAMETER CHECKS
            if (wsheet == null)
                throw new Exception("WriteTableToExcelSheet: Null-valued Worksheet object!");
            if (tbl == null || tbl.IsEmpty)
                throw new Exception("WriteTableToExcelSheet: Null or empty MatrixTable!");
            if (TopLeftCell == null || TopLeftCell == "")
                throw new Exception("WriteTableToExcelSheet: Null or empty string TopLeftCell!");

            // init parameters
            Excel.Range range = null;
            Excel.Range UpperLeftCell = null;
            string[] fieldnames;
            string fieldname;
            MetaData md = tbl.metaData;

            try
            {
                // get upper left cell in range
                UpperLeftCell = (Excel.Range)wsheet.Range(TopLeftCell).Cells[1, 1];

                // clear range for table, +1 row for header
                range = wsheet.Range(UpperLeftCell.Address);
                range = range.Resize(tbl.RowCount + 1, tbl.ColumnCount);
                range.Cells.ClearContents();

                // Write field values to range in this order: 
                // Text fields (string), numeric fields (int), Key figures (double)
                int ColumnOffset = 0;

                // Text attributes
                //***********************************************************
                int TextAtttribCount = 0;

                if (tbl.TextAttributeFields != null && tbl.TextAttributeFields.nLength > 0)
                {
                    fieldnames = tbl.TextAttributeFields.toArray;
                    TextAtttribCount = tbl.TextAttributeFields.nLength;

                    // write field names to header
                    for (int i = 0; i < fieldnames.Length; i++)
                    {
                        UpperLeftCell.Cells[1, i + 1].Value = fieldnames[i];
                    }

                    // get value matrix for fields (2-dim array)
                    string[,] FieldValues = tbl.TextAttribValues.toArray;

                    // row and column offsets (shifts w.r.t. upper left corner, +1 row-offset for header )
                    range = wsheet.Range(UpperLeftCell.Address).Offset(1, 0 + ColumnOffset);

                    // define range size for value matrix
                    range = range.Resize(tbl.RowCount, fieldnames.Length);

                    // set range values in bulk
                    // range.set_Value(value: FieldValues);
                    range.Value = FieldValues;

                    // set number format
                    range.NumberFormat = GetExcelNumberFormat(FieldType.TextAttribute);

                } // text attributes


                // Numeric attributes
                //***********************************************************
                int NumAtttribCount = 0;

                if (tbl.NumAttributeFields != null && tbl.NumAttributeFields.nLength > 0)
                {
                    fieldnames = tbl.NumAttributeFields.toArray;
                    NumAtttribCount = tbl.NumAttributeFields.nLength;

                    // write field names to header
                    for (int i = 0; i < fieldnames.Length; i++)
                    {
                        UpperLeftCell.Cells[1, i + 1 + TextAtttribCount].Value = fieldnames[i];
                    }

                    // get value matrix for fields (2-dim array)
                    int[,] FieldValues = tbl.NumAttribValues.toArray;

                    // row and column offsets (shifts w.r.t. upper left corner, +1 row-offset for header )
                    range = wsheet.Range(UpperLeftCell.Address).Offset(1, 0 + TextAtttribCount);

                    // define range size for value matrix
                    range = range.Resize(tbl.RowCount, fieldnames.Length);

                    // set range values in bulk
                    // range.set_Value(value: FieldValues);
                    range.Value = FieldValues;

                    // set number format for dates
                    range.NumberFormat = GetExcelNumberFormat(FieldType.IntegerAttribute);    // text for integer attributes

                    for (int i = 0; i < fieldnames.Length; i++)
                    {
                        if (MetaData.GetFieldType(md, fieldnames[i]) == FieldType.DateAttribute)
                        {
                            range = wsheet.Range(UpperLeftCell.Address).Offset(1, 0 + TextAtttribCount + i);
                            range = range.Resize(tbl.RowCount, 1);
                            range.NumberFormat = GetExcelNumberFormat(FieldType.DateAttribute); // "dd/mm/yyyy"
                        }
                    }

                } // numeric attributes

                // Key Figures
                //***********************************************************

                if (tbl.KeyFigureFields != null && tbl.KeyFigureFields.nLength > 0)
                {
                    fieldnames = tbl.KeyFigureFields.toArray;

                    // write field names to header
                    for (int i = 0; i < fieldnames.Length; i++)
                    {
                        UpperLeftCell.Cells[1, i + 1 + TextAtttribCount + NumAtttribCount].Value = fieldnames[i];
                    }

                    // get value matrix for fields (2-dim array)
                    double[,] FieldValues = tbl.KeyFigValues.toArray;

                    // row and column offsets (shifts w.r.t. upper left corner, +1 row-offset for header )
                    range = wsheet.Range(UpperLeftCell.Address).Offset(1, 0 + TextAtttribCount + NumAtttribCount);

                    // define range size for value matrix
                    range = range.Resize(tbl.RowCount, fieldnames.Length);

                    // set range values in bulk
                    // range.set_Value(value: FieldValues);
                    range.Value = FieldValues;

                    // set number format for dates
                    range.NumberFormat = GetExcelNumberFormat(FieldType.KeyFigure);

                } // key figures

                // make the header range bold
                range = wsheet.Range(UpperLeftCell.Address).Resize(1, tbl.ColumnCount);
                range.Font.Bold = true;
            }
            catch (Exception ex)
            {
                throw new Exception("WriteTableToExcelSheet: " + ex.Message);
            }
            finally
            {
                releaseObject(range);
                releaseObject(UpperLeftCell);
            }
        }

        /// <summary>
        /// Write MatrixTable into an Excel Table (ListObject).
        /// </summary>
        /// <param name="tbl">Input table of type MatrixTable</param>
        /// <param name="wsheet">Excel Worksheet object</param>
        /// <param name="xTableName">Name of Excel Table</param>
        /// <param name="TopLeftCell">Upper-left corner of Excel Table in worksheet</param>
        /// <param name="ClearSheetContent">If true, clear whole sheet </param>
        /// <returns>Excel table (ListObject)</returns>
        public static Excel.ListObject WriteTableToExcel(Excel.Worksheet wsheet, MatrixTable tbl,
            string xTableName, string TopLeftCell = "A1", 
            bool ClearSheetContent = false)
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
            if (wsheet == null)
                throw new Exception("ExcelFunc_NO.WriteToExcelTable: Null-valued Worksheet object!");

            if (xTableName == null || xTableName == "")
                throw new Exception("ExcelFunc_NO.WriteToExcelTable: Null or empty string TableName");
            
            // parameter checks OK..
            Excel.Range range = null;

            try
            {
                if (ClearSheetContent)
                    wsheet.Cells.ClearContents();  // clear sheet content

                // get upper left corner of range defined by CellStr
                range = (Excel.Range)wsheet.Range(TopLeftCell).Cells[1, 1];   //

                // Write table to range
                ExcelFunc_NO.WriteTableToExcel(wsheet, tbl, range.Address);

                // derive range for table, +1 row for table header
                range = range.Resize(tbl.RowCount + 1, tbl.ColumnCount);

                // add ListObject to sheet

                // ListObjects.AddEx Method 
                // http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.listobjects.addex%28v=office.14%29.aspx

                Excel.ListObject xtbl = ExcelFunc_NO.AddListObject(range, xTableName);

                // return excel table (ListObject)
                return (Excel.ListObject)xtbl;
            }
            catch (Exception ex)
            {
                throw new Exception("ExcelFunc_NO.WriteToExcelTable: " + ex.Message + "\n");
            }
            finally
            {
                try
                {
                    ExcelFunc_NO.releaseObject(range);
                }
                catch { }
            }
        }

        /// <summary>
        ///  Write MatrixTable into an Excel Table (ListObject).
        /// </summary>
        /// <param name="wbook">Excel Workbook</param>
        /// <param name="tbl">Input table of type MatrixTable</param>
        /// <param name="SheetName">Name of Excel Worksheet</param>
        /// <param name="xTableName">Name of Excel Table</param>
        /// <param name="TopLeftCell">Upper-left corner of Excel Table in worksheet</param>
        /// <param name="ClearSheetContent">If true, clear whole sheet</param>
        /// <returns>Excel table (ListObject)</returns>
        /// <remarks>
        /// If it doesn't exist already, a new worksheet with the given name will be created.
        /// </remarks>
        public static Excel.ListObject WriteTableToExcel(Excel.Workbook wbook, MatrixTable tbl,
            string SheetName, string xTableName, string TopLeftCell = "A1",
            bool ClearSheetContent = false)
        {
            // PARAMETER CHECKS
            if (wbook == null)
                throw new Exception("WriteToExcelTable: Null-valued Workbook object!");

            if (SheetName == null || SheetName == "")
                throw new Exception("WriteToExcelTable: Null-valued Workbook object!");

            if (xTableName == null || xTableName == "")
                throw new Exception("WriteToExcelTable: Null or empty string TableName");

            try
            {
                // Find worksheet , create one if not exists
                Excel.Worksheet WSheet = ExcelFunc_NO.GetWorksheet(wbook, SheetName, AddSheetIfNotFound: true);

                // call func
                return WriteTableToExcel(WSheet, tbl, xTableName, TopLeftCell, ClearSheetContent);
            }
            catch (Exception ex)
            {
                throw new Exception("WriteToExcelTable: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Read values of a key figure matrix from an excel range
        /// </summary>
        /// <param name="range">Excel range object</param>
        /// <param name="KeyFigReplaceNull">Replace null or empty values with this value</param>
        /// <returns>Key figure matrix with values from range</returns>
        public static KeyMatrix ReadKeyMatrixFromExcel(Excel.Range range, double KeyFigReplaceNull = 0)
        {
            // PARAMETER CHECKS
            if (range == null)
                throw new Exception("ReadKeyMatrixFromExcel: Null-valued excel range!");

            try
            {
                // get column and row sizes
                int rows = range.Rows.Count;
                int cols = range.Columns.Count;

                // create all-zeros KeyMatrix
                var M = KeyMatrix.CreateConstantMatrix(rows, cols);

                // read cell values row-by-row
                object CellValue;
                double dval;

                for (int i = 1; i <= rows; i++)
                {
                    for (int j = 1; j <= cols; j++)
                    {
                        CellValue = (object)range.Cells[i, j].Value2;

                        if (CellValue == null || CellValue == "")
                            dval = KeyFigReplaceNull;
                        else
                        {
                            // throw an error if element value can't be converted to double
                            try
                            {
                                dval = Convert.ToDouble(CellValue.ToString());
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("Cell value '" + CellValue + "' cannot be converted to double! " +
                                    "Row: " + i + " Column: " + j + " in Range\n" + ex.Message);
                            }
                        }

                        // assign value to matrix
                        M[i - 1, j - 1] = dval;
                    }
                }
                return M;
            }
            catch (Exception ex)
            {
                throw new Exception("ReadKeyMatrixFromExcel: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Read values of a key figure matrix from a named range
        /// </summary>
        /// <param name="RangeName">Name of range</param>
        /// <param name="WorkbookPath">Full file path of Excel workbook</param>
        /// <param name="KeyFigReplaceNull">Replace null or empty values with this value</param>
        /// <returns>Key figure matrix with values from range</returns>
        public static KeyMatrix ReadKeyMatrixFromExcel(string RangeName, string WorkbookPath = null,
            double KeyFigReplaceNull = 0)
        {
            // PARAMETER CHECKS
            if (RangeName == null || RangeName == "")
                throw new Exception("ReadKeyMatrixFromExcel: Null-valued or empty string RangeName!");

            try
            {
                // Get application insance
                Excel.Application xlapp = GetExcelApplicationInstance();

                // Get workbook object
                Excel.Workbook wbook;

                if (WorkbookPath == null || WorkbookPath == "")
                    wbook = xlapp.ActiveWorkbook;
                else
                    wbook = GetWorkbookByFullName(xlapp, WorkbookPath);

                // get range from name
                Excel.Range rng = wbook.Names[RangeName, null, null].RefersToRange;   // ?

                if (rng == null)
                    throw new Exception("Range could not be derived from rage name!");

                // call method
                return ReadKeyMatrixFromExcel(rng, KeyFigReplaceNull);
            }
            catch (Exception ex)
            {
                throw new Exception("ReadKeyMatrixFromExcel: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Write Key Figure Matrix (KeyMatrix) to an Excel range
        /// </summary>
        /// <param name="wsheet">Excel worksheet</param>
        /// <param name="M">Key Figure Matrix (KeyMatrix)</param>
        /// <param name="TopLeftCell">Upper-left corner of matrix in excel sheet</param>
        /// <param name="RangeName">Name to be assigned to the range of matrix</param>
        /// <returns>Range of matrix</returns>
        public static Excel.Range WriteKeyMatrixToExcel(Excel.Worksheet wsheet, KeyMatrix M,
            string TopLeftCell = "A1", string RangeName = null)
        {
            // PARAMETER CHECKS
            if (wsheet == null)
                throw new Exception("WriteKeyMatrixToExcel: Null-valued Worksheet object!");
            if (M == null || M.IsEmpty)
                throw new Exception("WriteKeyMatrixToExcel: Null or empty KeyMatrix!");
            if (TopLeftCell == null || TopLeftCell == "")
                throw new Exception("WriteKeyMatrixToExcel: Null or empty string TopLeftCell!");
            if (RangeName != null && RangeName == "")
                throw new Exception("WriteKeyMatrixToExcel: Empty string RangeName!");

            // init parameters
            Excel.Range range = null;
            Excel.Range UpperLeftCell = null;
            int rows = M.nRows;
            int cols = M.nCols;

            try
            {
                // get upper left cell in range
                UpperLeftCell = (Excel.Range)wsheet.Range(TopLeftCell).Cells[1, 1];

                // clear range for matrix, row for header
                range = wsheet.Range(UpperLeftCell.Address);
                range = range.Resize(rows, cols);
                range.Cells.ClearContents();

                // write matrix values to range, in bulk (array assignment)
                double[,] dvalues = M.toArray;
                range.Value = dvalues;

                // assign name to range
                if (RangeName != null) range.Name = RangeName;

                // return range
                return range;
            }
            catch (Exception ex)
            {
                throw new Exception("WriteKeyMatrixToExcel: " + ex.Message + "\n");
            }
            finally
            {
                releaseObject(UpperLeftCell);
            }
        }

        /// <summary>
        /// Write Key Figure Matrix (KeyMatrix) to an Excel range
        /// </summary>
        /// <param name="wbook">Excel workbook</param>
        /// <param name="M">Key Figure Matrix (KeyMatrix)</param>
        /// <param name="SheetName">Name of worksheet</param>
        /// <param name="TopLeftCell">Upper-left corner of matrix in excel sheet</param>
        /// <param name="RangeName">Name to be assigned to the range of matrix</param>
        /// <returns>Excel range</returns>
        public static Excel.Range WriteKeyMatrixToExcel(Excel.Workbook wbook, KeyMatrix M,
            string SheetName, string TopLeftCell = "A1", string RangeName = null)
        {
            // PARAMETER CHECKS
            if (wbook == null)
                throw new Exception("WriteKeyMatrixToExcel: Null-valued Workbook object!");

            if (SheetName == null || SheetName == "")
                throw new Exception("WriteKeyMatrixToExcel: Null-valued Workbook object!");

            try
            {
                // Find worksheet , create one if not exists
                Excel.Worksheet WSheet = ExcelFunc_NO.GetWorksheet(wbook, SheetName, AddSheetIfNotFound: true);

                // call func
                return WriteKeyMatrixToExcel(WSheet, M, TopLeftCell, RangeName);
            }
            catch (Exception ex)
            {
                throw new Exception("WriteKeyMatrixToExcel: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Checks if an excel workbook is opened (with the given application instance)
        /// </summary>
        /// <param name="xlapp">Excel Application instance</param>
        /// <param name="WBfilename">File name of excel workbook (not file path)</param>
        /// <returns>True if workbook is open</returns>
        public static bool IsExcelWBOpen1(Excel.Application xlapp, string WBfilename)
        {
            bool isOpened = true;
            try
            {
                var wb = xlapp.Workbooks[WBfilename];
            }
            catch (Exception)
            {
                isOpened = false;
            }
            return isOpened;
        }

        /// <summary>
        /// Checks if an excel workbook is opened (with the given application instance)
        /// </summary>
        /// <param name="xlapp">Excel Application instance</param>
        /// <param name="WBfilepath">Complete file path of excel workbook</param>
        /// <returns>True if workbook is open</returns>
        public static bool IsExcelWBOpen2(Excel.Application xlapp, string WBfilepath)
        {
            if (xlapp == null)
                throw new Exception("ExcelFunc.IsExcelWBOpen2: Null-valued application instance");

            var wbs = (Excel.Workbooks) xlapp.Workbooks;

            try
            {
                foreach (var wbook in wbs)
                {
                    if (WBfilepath == ((Excel.Workbook)wbook).FullName)
                        return true;
                }
            }
            catch { return false; }

            return false;
        }

        /// <summary>
        /// Open an excel workbook in editable mode. 
        /// </summary>
        /// <param name="xlapp">Current Excel Application</param>
        /// <param name="WBfilepath">Full file name (path)</param>
        /// <param name="ReadOnly">If true, open workbook in read-only mode; otherwise in editable mode.</param>
        /// <remarks>
        /// Does nothing if workbook was already opened.
        /// Throws error file does not exist.
        /// </remarks>
        public static Excel.Workbook OpenExcelWorkbook(Excel.Application xlapp, string WBfilepath, bool ReadOnly = false)
        {
            // check inputs
            if (xlapp == null)
                throw new Exception("ExcelFunc.OpenExcelWorkbook: Null-valued application instance!\n");
            if (WBfilepath == null || WBfilepath == "")
                throw new Exception("ExcelFunc.OpenExcelWorkbook: Null or empty string WBfilepath!\n");
            try
            {
                // check if file was already opened
                if (IsExcelWBOpen2(xlapp, WBfilepath))
                    return xlapp.Workbooks[Path.GetFileName(WBfilepath)];

                // check if file exists
                if (! File.Exists(WBfilepath))
                    throw new Exception("ExcelFunc.OpenExcelWorkbook: A file with the given path could not be found!\n");

                // open excel file in editable modus
                return xlapp.Workbooks.Open(Path.GetFileName(WBfilepath), Type.Missing, ReadOnly,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            }
            catch (Exception ex)
            {
                throw new Exception("ExcelFunc_NO.OpenExcelWorkbook: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Find ListObject by name in workbook. Return null if not found.
        /// </summary>
        /// <param name="wbook">Workbook object</param>
        /// <param name="TableName">Name of excel table</param>
        /// <returns>ListObject object</returns>
        public static Excel.ListObject GetListObject(Excel.Workbook wbook, string TableName)
        {
            if (wbook == null) return null;
            Excel.ListObject xlist = null;
            try
            {
                if (wbook.Worksheets.Count == 0) return xlist;

                foreach (var wsheet in wbook.Worksheets)
                {
                    if (((Excel.Worksheet)wsheet).ListObjects == null) continue; // don't know if this is necessary

                    for (int i = 1; i <= ((Excel.Worksheet)wsheet).ListObjects.Count; i++)
                    {
                        if (((Excel.Worksheet)wsheet).ListObjects[i].Name == TableName)
                        {
                            xlist = ((Excel.Worksheet)wsheet).ListObjects[i];
                            return xlist;
                        }
                    }
                }
            }
            catch
            {
                return xlist;
            }
            return xlist;
        }

        /// <summary>
        /// Return Worksheet object if found, otherwise null.
        /// </summary>
        /// <param name="wbook">Excel Workbook</param>
        /// <param name="SheetName">Name of worksheet</param>
        /// <param name="AddSheetIfNotFound">If true, add a new worksheet if not found</param>
        /// <returns>Worksheet object</returns>
        public static Excel.Worksheet GetWorksheet(Excel.Workbook wbook, string SheetName,
            bool AddSheetIfNotFound = false)
        {
            Excel.Worksheet wsheet = null;
            if (wbook == null) return null;

            foreach (Excel.Worksheet sheet in wbook.Worksheets)
            {
                if (sheet.Name.Equals(SheetName))
                {
                    wsheet = (Excel.Worksheet) wbook.Worksheets[SheetName];
                    return wsheet;
                }
            }

            if (AddSheetIfNotFound)
            {
                wsheet = (Excel.Worksheet)wbook.Worksheets.Add();
                wsheet.Name = SheetName;
                return wsheet;
            }
            return wsheet;
        }

        /// <summary>
        /// Define a range with a header row including column names as ListObject (excel table) 
        /// </summary>
        /// <param name="rng">Excel range</param>
        /// <param name="xTableName">Name for ListObject (excel table)</param>
        /// <returns>ListObject handle</returns>
        public static Excel.ListObject AddListObject(Excel.Range rng, string xTableName)
        {
            // get worksheet of range
            Excel.Worksheet ws = rng.Worksheet;

            // add ListObject
            Excel.ListObject tbl = (Excel.ListObject)ws.ListObjects.Add(
                sourceType: XlListObjectSourceType.xlSrcRange,
                source: rng,
                linkSource: Type.Missing,
                xlListObjectHasHeaders: XlYesNoGuess.xlYes); // ?

            // set name of excel table
            tbl.Name = xTableName;
            return tbl;
        }

        /// <summary>
        /// Return an object instance of excel application.
        /// Start excel application if it has not been started already.
        /// </summary>
        /// <returns>Excel Application object</returns>
        public static Excel.Application GetExcelApplicationInstance()
        {
            Excel.Application instance = null;

            try
            {
                // instance = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                // Excel DNA
                instance = new Excel.Application(null, ExcelDnaUtil.Application);

                // Local Tests in Console Application mode
                //instance = NetOffice.ExcelApi.Application.GetActiveInstance();
            }
            catch (Exception ex)
            {
                instance = new Excel.Application();
            }

            return instance;
        }

        /// <summary>
        /// Get the list of all ListObjects (excel tables) in Workbook
        /// </summary>
        /// <param name="wbook">Excel Workbook</param>
        /// <param name="TableNames">Array with table names</param>
        /// <returns>Array of all ListObjects. Return null if there is none.</returns>
        public static Excel.ListObject[] GetAllExcelTables(Excel.Workbook wbook, out string[] TableNames)
        {
            var TableList = new List<Excel.ListObject>();
            var TableNameList = new List<string>();

            // search for all ListObjects in all sheets
            foreach (var wsheet in wbook.Worksheets)
            {
                foreach (var tbl in ((Excel.Worksheet)wsheet).ListObjects)
                {
                    TableList.Add(tbl);
                    TableNameList.Add(tbl.Name);
                }
            }
            // return
            TableNames = TableNameList.ToArray();
            return TableList.ToArray();
        }

        /// <summary>
        /// Get names of all worksheets in an excel workbook
        /// </summary>
        /// <param name="wbook">Excel workbook</param>
        /// <returns>Array of worksheet names</returns>
        public static string[] GetAllSheetNames(Excel.Workbook wbook)
        {
            var SheetList = new List<string>();

            if (wbook == null || wbook.Worksheets == null || wbook.Worksheets.Count == 0) 
                return null;

            // search for all ListObjects in all sheets
            foreach (var wsheet in wbook.Worksheets)
            {
                SheetList.Add(((Excel.Worksheet)wsheet).Name);
            }
            return SheetList.ToArray();
        }

        /// <summary>
        /// Get all range names in a workbook
        /// </summary>
        /// <param name="wbook">Excel workbook</param>
        /// <returns>Array of range names</returns>
        public static string[] GetAllRangeNames(Excel.Workbook wbook)
        {
            if (wbook == null || wbook.Names == null || wbook.Names.Count == 0)
                return null;

            string[] NameList = new string[wbook.Names.Count];

            int i = 0;
            foreach (var item in wbook.Names)
            {
                NameList[i++] = item.Name;
            }
            return NameList;
        }

        /// <summary>
        /// Get field type of a cell in worksheet.
        /// Possible types: (text, integer, date, key figure, undefined)
        /// </summary>
        /// <param name="xcell">Worksheet cell</param>
        /// <returns>Field type</returns>
        /// <remarks>
        /// Rules:
        /// ValueType = String --> text attribute
        /// ValueType = DateTime AND Value2Type = Double --> date attribute
        /// ValueType = Double AND NumberFormat = @ (text) --> integer attribute
        /// ValueType = Double (and not integer or DateTime) --> key figure
        /// </remarks>
        public static FieldType GetFieldTypeOfExcelCell(Excel.Range xcell)
        {
            // get upper-left cell
            Excel.Range TopLeftCell = xcell.Cells[1, 1];

            // get values & number format of cell
            FieldType ftype;
            object val = TopLeftCell.Value;
            object val2 = TopLeftCell.Value2;
            string NumberFormat = TopLeftCell.NumberFormat.ToString();

            if (val2 == null || val2 == "")
                return FieldType.Undefined;

            // TEST
            System.Diagnostics.Debug.WriteLine("GetFieldTypeOfExcelCell: Cell Value: " + val);
            System.Diagnostics.Debug.WriteLine("GetFieldTypeOfExcelCell: Cell Value2: " + val2);

            if (val2.GetType() == typeof(string))
            {
                ftype = FieldType.TextAttribute;
            }
            else if (val2.GetType() == typeof(double) && val.GetType() == typeof(DateTime))
            {
                ftype = FieldType.DateAttribute;
            }
            else if (val2.GetType() == typeof(double) && NumberFormat == "@") // number format == text
            {
                ftype = FieldType.IntegerAttribute;
            }
            else if (val2.GetType() == typeof(double))
            {
                ftype = FieldType.KeyFigure;
            }
            else
            {
                ftype = FieldType.Undefined;
            }
            return ftype;
        }

        /// <summary>
        /// Return proper excel number format for the given field type
        /// </summary>
        /// <param name="ftype">Field type</param>
        /// <returns>Number format for excel</returns>
        public static string GetExcelNumberFormat(FieldType ftype)
        {
            string NumberFormat = "General";

            switch (ftype)
            {
                case FieldType.DateAttribute:
                    NumberFormat = "dd/mm/yyyy";
                    break;

                case FieldType.IntegerAttribute:
                    NumberFormat = "@";     // text format
                    break;

                case FieldType.KeyFigure:
                    NumberFormat = "General";   // default format
                    break;

                case FieldType.TextAttribute:
                    NumberFormat = "General";   // default format
                    break;

                case FieldType.Undefined:
                    throw new Exception("ExcelFunc_NO.GetExcelNumberFormat: Undefined field type!\n"); 
            }
            return NumberFormat;
        }

        /// <summary>
        /// General method for calling all non-static methods of given finaquant class 
        /// (Finaquant Protos or Finaquant Calcs).
        /// </summary>
        /// <param name="ClassName">Name of finaquant class, like 'MatrixTable'</param>
        /// <param name="Instance">Object instance of class which is owner of the method</param>
        /// <param name="MethodName">Method's name</param>
        /// <param name="ParameterTypeNames">Array of parameter type names, like 'System.Double' or 'MatrixTable' (set to null if not required)</param>
        /// <param name="parameters">Parameters of method in correct order (set to null if there is no parameter)</param>
        /// <returns>Outputs of method; null if method's return type is void</returns>
        public static object CallAnyMethodByName(string ClassName, object Instance, string MethodName, string[] ParameterTypeNames,
            params object[] parameters)
        {
            try
            {
                Type ClassType = GetTypeFromTypeName(ClassName);

                var TypeList = new List<Type>();
                Type[] ParTypes;

                if (ParameterTypeNames != null)
                {
                    for (int i = 0; i < ParameterTypeNames.Length; i++)
                    {
                        TypeList.Add(GetTypeFromTypeName(ParameterTypeNames[i]));
                    }
                    ParTypes = TypeList.ToArray();
                }
                else
                {
                    ParTypes = null;
                }

                // call primary method
                return CallAnyMethodByName(ClassType, Instance, MethodName, ParTypes, parameters);
            }
            catch (Exception ex)
            {
                throw new Exception("ExcelFunc_NO.CallNonStaticMethod: " + ex.Message);
            }
        }

        /// <summary>
        /// General method for calling all methods of given finaquant class 
        /// (Finaquant Protos or Finaquant Calcs).
        /// </summary>
        /// <param name="ClassType">Type of finaquant class, like typeof(MatrixTable)</param>
        /// <param name="Instance">Object instance of class which is owner of the method</param>
        /// <param name="MethodName">Method's name</param>
        /// <param name="ParameterTypes">Array of parameter types (set to null if not required)</param>
        /// <param name="parameters">Parameters of method in correct order (set to null if there is no parameter)</param>
        /// <returns>Outputs of method; null if method's return type is void</returns>
        /// <remarks>
        /// See: MethodBase.Invoke Method (Object, Object[])
        /// http://msdn.microsoft.com/en-us/library/a89hcwhh%28v=vs.110%29.aspx
        /// Set ParameterTypes to null if parameter types (signature) are not required to resolve
        /// method name ambiguities.
        /// The parameter 'Instance' is ignored if the method is a static one.
        /// </remarks>
        public static object CallAnyMethodByName(Type ClassType, object Instance, string MethodName, Type[] ParameterTypes,
            params object[] parameters)
        {
            // Get type handle
            Type tp = Type.GetType(ClassType.AssemblyQualifiedName);

            // Get method handle
            MethodInfo mtd;

            if (ParameterTypes == null)
                mtd = tp.GetMethod(MethodName);
            else
                mtd = tp.GetMethod(MethodName, ParameterTypes);

            object result = null;

            if (mtd.ReturnType == typeof(void))
            {
                mtd.Invoke(Instance, parameters); // call method
            }
            else
            {
                result = mtd.Invoke(Instance, parameters); // call method
            }

            return result;
        }

        /// <summary>
        /// Get type from type (or class) name; for System and Finaquant types only
        /// </summary>
        /// <param name="TypeName">Name of type, like 'System.Double' or 'MatrixTable'</param>
        /// <returns>Type; null if not found</returns>
        public static Type GetTypeFromTypeName(string TypeName)
        {
            if (TypeName == null || TypeName == "")
            {
                return null;
            }

            // capture type names like System.String, System.Double. System.Int32
            Type ftype = Type.GetType(TypeName);

            if (ftype != null)
                return ftype;

            else // check if finaquant type
            {
                // get finaquant namespace (can be calcs or protos)
                // see: http://stackoverflow.com/questions/179102/getting-a-system-type-from-types-partial-name

                var assembly = typeof(MatrixTable).Assembly;
                string nspace = typeof(MatrixTable).Namespace;
                return assembly.GetType(nspace + "." + TypeName);
            }
        }

        /// <summary>
        /// Get Workbook object by workbook's full name (file path)
        /// </summary>
        /// <param name="xlapp">Excel application</param>
        /// <param name="WbookFullName">Workbook's full file path</param>
        /// <returns>Workbook object</returns>
        public static Excel.Workbook GetWorkbookByFullName(Excel.Application xlapp, string WbookFullName)
        {
           Excel.Workbook wbook = null;

           if (ExcelFunc_NO.IsExcelWBOpen2(xlapp, WbookFullName))
           {
               wbook = xlapp.Workbooks[Path.GetFileName(WbookFullName)];
           }
           else
           {
               throw new Exception("ExcelFunc_NO.GetWorkbookByFullName: An opened workbook with the given full name could not be found!");
           }
           return wbook;
        }

        /// <summary>
        /// Input box with range type for passing parameter values to a function
        /// </summary>
        /// <param name="xlapp">Excel Application</param>
        /// <param name="prompt">Prompt</param>
        /// <param name="title">Title</param>
        /// <param name="ParameterCount">Parameter count</param>
        /// <param name="ParameterValues">Returned parameter values</param>
        /// <param name="CheckEqualty">If true, check equalty of selected-cell-count with ParameterCount;
        /// otherwise, check if selected-cell-count >= ParameterCount</param>
        public static bool InputBoxRange(Excel.Application xlapp, string prompt, string title, 
            int ParameterCount, out List<object> ParameterValues, bool CheckEqualty = true)
        {
            try
            {
                object response = xlapp.InputBox(prompt, title, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, 8);

                Excel.Range rng = null;

                if (!(response is Excel.Range))
                {
                    ParameterValues = null;
                    return false;
                }
                else
                {
                    rng = (Excel.Range)response;
                }

                if (rng == null)
                {
                    ParameterValues = null;
                    return false;
                }

                if (rng.Columns.Count != 1 )
                {
                    throw new Exception("Please select a single-column (vertical) range for parameters.");
                }

                if (CheckEqualty && rng.Cells.Count != ParameterCount)
                {
                    throw new Exception(ParameterCount.ToString() + " cells need to be selected as input.");
                }

                if (! CheckEqualty && rng.Cells.Count < ParameterCount)
                {
                    throw new Exception("At least " + ParameterCount + " cells need to be selected as input.");
                }

                // read parameter values from selected range
                ParameterValues = new List<object>();
                object val;

                for (int i = 0; i < ParameterCount; i++)
                {
                    val = rng[i + 1].Value2;

                    // remove leading and trailing white-space characters
                    if (val.GetType() == typeof(string))
                        val = val.ToString().Trim();

                    ParameterValues.Add(val);
                }
                return true;
            }
            catch(Exception ex)
            {
                throw new Exception("ExcelFunc_NO.InputBoxRange: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Return string address of shifted cell
        /// </summary>
        /// <param name="wsheet">worksheet</param>
        /// <param name="CellStr">String address of a cell like "A1"</param>
        /// <param name="nRows">Vertical row offset</param>
        /// <param name="nColumns">Horizontal column offset</param>
        /// <returns>String address of shifted cell</returns>
        public static string ShiftCell(Excel.Worksheet wsheet, string CellStr, int nRows, int nColumns)
        {
            return wsheet.Range(CellStr).Offset(nRows, nColumns).Address;
        }

        /// <summary>
        /// Helper function for releasing unused resources (COM objects)
        /// </summary>
        /// <param name="obj">Excel and COM objects</param>
        public static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                System.Diagnostics.Debug.WriteLine("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        /// <summary>
        /// Create Test Tables in Excel:
        /// 
        /// </summary>
        public static void CreateTestTables(Excel.Workbook wbook)
        {
            // define metadata
            MetaData md = MetaData.CreateEmptyMetaData();
            MetaData.AddNewField(md, "country", FieldType.TextAttribute);
            MetaData.AddNewField(md, "category", FieldType.TextAttribute);
            MetaData.AddNewField(md, "brand", FieldType.TextAttribute);
            MetaData.AddNewField(md, "product", FieldType.TextAttribute);
            MetaData.AddNewField(md, "modelyear", FieldType.IntegerAttribute);
            MetaData.AddNewField(md, "typeid", FieldType.IntegerAttribute);
            MetaData.AddNewField(md, "date", FieldType.DateAttribute);
            MetaData.AddNewField(md, "costs", FieldType.KeyFigure);
            MetaData.AddNewField(md, "price", FieldType.KeyFigure);
            MetaData.AddNewField(md, "margin", FieldType.KeyFigure);
            MetaData.AddNewField(md, "sales", FieldType.KeyFigure);

            // define table structure
            var CostTableFields = TableFields.CreateEmptyTableFields(md);
            // TableFields.AddNewField(CostTableFields, "country");
            TableFields.AddNewField(CostTableFields, "category");
            TableFields.AddNewField(CostTableFields, "product");
            TableFields.AddNewField(CostTableFields, "brand");
            TableFields.AddNewField(CostTableFields, "modelyear");
            TableFields.AddNewField(CostTableFields, "date");
            TableFields.AddNewField(CostTableFields, "costs");

            // create test table by combinating field values
            // ... random values for key figures

            // attribute values
            // text attributes
            // TextVector CountryVal = TextVector.CreateVectorWithElements("Peru", "Paraguay", "Argentina", "Brasil");
            TextVector CategoryVal = TextVector.CreateVectorWithElements("Economy", "Luxury", "Sports", "Family");
            TextVector BrandVal = TextVector.CreateVectorWithElements("Toyota", "Honda", "BMW", "Audi");
            TextVector ProductVal = TextVector.CreateVectorWithElements("Car", "Bus", "Motor");
            // numeric attributes
            NumVector ModelVal = NumVector.CreateVectorWithElements(2008, 2009, 2010);
            NumVector DateVal = NumVector.CreateSequenceVector(
                StartValue: DateFunctions.DayToNumber(1, 1, 2010),
                Interval: 10, nLength: 5);

            // initiate field value dictionaries
            var TextAttribValues = new Dictionary<string, TextVector>();
            var NumAttribValues = new Dictionary<string, NumVector>();
            // assign a value vector to each field
            // TextAttribValues["country"] = CountryVal;
            TextAttribValues["category"] = CategoryVal;
            TextAttribValues["product"] = ProductVal;
            TextAttribValues["brand"] = BrandVal;
            NumAttribValues["modelyear"] = ModelVal;
            NumAttribValues["date"] = DateVal;

            // default range for all key figures
            KeyValueRange DefaultRangeForAllKeyFigures = KeyValueRange.CreateRange(5000, 10000);

            // range for selected key figures
            var RangeForSelectedKeywords = new Dictionary<string, KeyValueRange>();
            RangeForSelectedKeywords["margin"] = KeyValueRange.CreateRange(0.10, 0.60);

            //***************************************************************************
            // Create Cost Table
            //***************************************************************************

            MatrixTable CostTable = MatrixTable.CombinateFieldValues_B(CostTableFields,
                TextAttribValues, NumAttribValues, DefaultRangeForAllKeyFigures, RangeForSelectedKeywords);

            // round all key figures to 2 digits after decimal point
            CostTable = MatrixTable.TransformKeyFigures(CostTable, x => Math.Round(x, 2),
                InputKeyFig: null, OutputKeyFig: null);

            // view table
            // MatrixTable.View_MatrixTable(CostTable, "Cost table");

            //***************************************************************************
            // Create Margin Tables
            //***************************************************************************

            // define table structure for MarginTable1
            var MarginTable1Fields = TableFields.CreateEmptyTableFields(md);
            TableFields.AddNewField(MarginTable1Fields, "category");
            TableFields.AddNewField(MarginTable1Fields, "margin");

            // create MarginTable1 with elements
            MatrixTable MarginTable1 = MatrixTable.CreateTableWithElements_A(MarginTable1Fields,
                "Economy", 0.20,
                "Luxury", 0.50,
                "Sports", 0.35,
                "Family", 0.30
                );

            // define table structure for MarginTable2
            var MarginTable2Fields = TableFields.CreateEmptyTableFields(md);
            TableFields.AddNewField(MarginTable2Fields, "category");
            TableFields.AddNewField(MarginTable2Fields, "brand");
            TableFields.AddNewField(MarginTable2Fields, "margin");

            // create MarginTable2 with elements
            MatrixTable MarginTable2 = MatrixTable.CreateTableWithElements_A(MarginTable2Fields,
                "Economy", "ALL", 0.20,
                "Luxury", "Toyota", 0.30,
                "Luxury", "Honda", 0.35,
                "Luxury", "ALL", 0.40,
                "Sports", "ALL", 0.50,
                "Family", "ALL", 0.33
                );

            // define table structure for MarginTable3
            var MarginTable3Fields = TableFields.CreateEmptyTableFields(md);
            TableFields.AddNewField(MarginTable3Fields, "modelyear");
            TableFields.AddNewField(MarginTable3Fields, "margin");

            // create MarginTable3 with elements
            MatrixTable MarginTable3 = MatrixTable.CreateTableWithElements_A(MarginTable3Fields,
                2008, 0.2,
                2009, 0.3,
                0, 0.4);

            //***************************************************************************
            // Create Condition Table
            //***************************************************************************

            // define table structure for CondTable
            var CondTableFields = TableFields.CreateEmptyTableFields(md);
            TableFields.AddNewField(CondTableFields, "category");
            TableFields.AddNewField(CondTableFields, "product");
            TableFields.AddNewField(CondTableFields, "brand");
            TableFields.AddNewField(CondTableFields, "modelyear");

            // create CondTable with elements
            MatrixTable CondTable = MatrixTable.CreateTableWithElements_A(CondTableFields,
                "Economy", "ALL", "Toyota", 0,
                "ALL", "ALL", "Honda", 2010,
                "ALL", "Bus", "BMW", 0);

            //***************************************************************************
            // Create Distribution Key Tables
            //***************************************************************************

            // define table structure for DistrKeyTbl1
            var TeamTblFields = TableFields.CreateEmptyTableFields(md);
            TableFields.AddNewField(TeamTblFields, "department", FieldType.TextAttribute);
            TableFields.AddNewField(TeamTblFields, "team", FieldType.TextAttribute);

            MatrixTable TeamTbl = MatrixTable.CreateTableWithElements_A(TeamTblFields,
                "Management", "Management Team A",
                "Management", "Management Team B",
                "Management", "Markteting Team X",
                "Management", "Marketing Team Y",
                "Production", "Production Team 1",
                "Production", "Production Team 2");

            // get aggregated cost table
            MatrixTable Tbl = CostTable.SelectColumns(TextVector.CreateVectorWithElements(
                "category", "brand"));

            NumVector nv1, nv2;
            Tbl = MatrixTable.UniqueAttributeRows(Tbl, out nv1, out nv2);

            // cartesian multiplication of tables
            MatrixTable DistrKeyTbl1 = MatrixTable.CombinateTableRows(Tbl, TeamTbl);

            // insert key figure
            md.AddFieldIfNew("distr_ratio", FieldType.KeyFigure);
            DistrKeyTbl1 = MatrixTable.InsertNewColumn(DistrKeyTbl1, "distr_ratio", 0.0);
            DistrKeyTbl1 = MatrixTable.AssignRandomValues(DistrKeyTbl1, "distr_ratio", 0.1, 0.5);
            DistrKeyTbl1 = MatrixTable.Round(DistrKeyTbl1, 2);

            //***************************************************************************
            // Create Table with Field Values (for combination example)
            //***************************************************************************

            // define table structure for ValueTbl
            var ValueTblFields = TableFields.CreateEmptyTableFields(md);
            TableFields.AddNewField(ValueTblFields, "department", FieldType.TextAttribute);
            TableFields.AddNewField(ValueTblFields, "team", FieldType.TextAttribute);
            TableFields.AddNewField(ValueTblFields, "year", FieldType.IntegerAttribute);
            TableFields.AddNewField(ValueTblFields, "date", FieldType.DateAttribute);

            MatrixTable ValueTbl = MatrixTable.CreateTableWithElements_A(ValueTblFields,
                "Management", "Team A", 2008, DateFunctions.DayToNumber(1, 1, 2010),
                "", "Team B", 2010, DateFunctions.DayToNumber(1, 1, 2012),
                "", "Team X", 0, 0,
                "Management", "Team Y", 0, 0, 
                "Production", "Team 1", 2008, 0,
                "", "Team 2", 0, 0
                );

            // insert tables into excel sheets
            WriteTableToExcel(wbook, CostTable, "CostTable", "Cost", "A1", true);
            
            WriteTableToExcel(wbook, MarginTable1, "MarginTables", "Margin1", "A1", true);
            WriteTableToExcel(wbook, MarginTable2, "MarginTables", "Margin2", "A7", false);
            WriteTableToExcel(wbook, MarginTable3, "MarginTables", "Margin3", "A16", false);

            WriteTableToExcel(wbook, CondTable, "CondTable", "Cond", "A1", true);
            WriteTableToExcel(wbook, DistrKeyTbl1, "DistrKeyTable", "DistrKey", "A1", true);
            WriteTableToExcel(wbook, ValueTbl, "FieldValuesTable", "FieldValues", "A1", true);

            // create test tables for commission calculation
            MatrixTable SalesTable, ComScaleTable, ScalePoolTable, ProdPoolTable;

            Create_Sales_And_Commission_Tables(md,
                out SalesTable, out ComScaleTable, out ScalePoolTable, out ProdPoolTable);

            WriteTableToExcel(wbook, SalesTable, "Sales", "SalesTable", "A1", true);
            WriteTableToExcel(wbook, ComScaleTable, "ComScale", "ComScaleTable", "A1", true);
            WriteTableToExcel(wbook, ScalePoolTable, "ScalePool", "ScalePoolTable", "A1", true);
            WriteTableToExcel(wbook, ProdPoolTable, "ProductPool", "ProdPoolTable", "A1", true);

        }


        // Helper method for creating sales and commission (schedule) tables
        private static void Create_Sales_And_Commission_Tables(MetaData md,
            out MatrixTable SalesTable, out MatrixTable ComScaleTable, out MatrixTable ScalePoolTable,
            out MatrixTable PoolTable)
        {
            // define metadata
            // md = MetaData.CreateEmptyMetaData();
            md.AddFieldIfNew("category", FieldType.TextAttribute);
            md.AddFieldIfNew("product", FieldType.TextAttribute);
            md.AddFieldIfNew("dealer", FieldType.TextAttribute);
            md.AddFieldIfNew("sales", FieldType.KeyFigure);
            md.AddFieldIfNew("date", FieldType.DateAttribute);

            md.AddFieldIfNew("scale_id", FieldType.IntegerAttribute);   // commission scale no
            md.AddFieldIfNew("pool_id", FieldType.IntegerAttribute);

            md.AddFieldIfNew("lower_limit", FieldType.KeyFigure);
            md.AddFieldIfNew("commission_rate", FieldType.KeyFigure);
            md.AddFieldIfNew("scale_logic", FieldType.TextAttribute);   // level or ?
            md.AddFieldIfNew("sales_per_period", FieldType.KeyFigure);

            // define Product-Category table
            var PrdCategoryFields = TableFields.CreateEmptyTableFields(md);
            TableFields.AddNewField(PrdCategoryFields, "category");
            TableFields.AddNewField(PrdCategoryFields, "product");

            // create CostTable with elements
            MatrixTable PrdCategoryTable = MatrixTable.CreateTableWithElements_A(PrdCategoryFields,
                "Sports", "BMW speed 200",
                "Sports", "BMW speed 300",
                "Sports", "Audi race X",
                "Sports", "Audi race Z",
                "Sports", "Honda adonis 7",
                "Sports", "Honda adonis 9",
                "Luxury", "BMW comfort",
                "Luxury", "Audi limu",
                "Luxury", "Honda expense",
                "Economy", "BMW budget",
                "Economy", "Audi cheap",
                "Economy", "Honda save"
                );

            // SalesTable

            // define table structure
            var SalesTableFields = TableFields.CreateEmptyTableFields(md);
            // TableFields.AddNewField(SalesTableFields, "category");
            TableFields.AddNewField(SalesTableFields, "product");
            TableFields.AddNewField(SalesTableFields, "dealer");
            TableFields.AddNewField(SalesTableFields, "date");
            TableFields.AddNewField(SalesTableFields, "sales");

            // create test table by combinating field values
            // ... random values for key figures

            // attribute values
            // text attributes
            TextVector ProductVal = TextVector.Unique(PrdCategoryTable.GetColumnTextAttribute("product"));
            TextVector DealerVal = TextVector.CreateVectorWithElements("Merlin", "Belinda", "Timon", "Proda", "Maximus", "Tinkerbell");

            // numeric attributes
            int StartVal = DateFunctions.DayToNumber(20, 10, 2011);
            int EndVal = DateFunctions.DayToNumber(30, 3, 2013);
            int interval = 10;   // days
            int Len = (EndVal - StartVal) / interval;
            NumVector DateVal = NumVector.CreateSequenceVector(StartValue: StartVal, Interval: interval, nLength: Len);

            // initiate field value dictionaries
            var TextAttribValues = new Dictionary<string, TextVector>();
            var NumAttribValues = new Dictionary<string, NumVector>();
            // assign a value vector to each field
            TextAttribValues["product"] = ProductVal;
            TextAttribValues["dealer"] = DealerVal;
            NumAttribValues["date"] = DateVal;

            // default range for all key figures
            KeyValueRange DefaultRangeForAllKeyFigures = KeyValueRange.CreateRange(5000, 10000);
            // range for selected key figures
            var RangeForSelectedKeywords = new Dictionary<string, KeyValueRange>();
            RangeForSelectedKeywords["sales"] = KeyValueRange.CreateRange(5.0, 200.0);

            // create sales table
            SalesTable = MatrixTable.CombinateFieldValues_B(SalesTableFields,
                TextAttribValues, NumAttribValues, DefaultRangeForAllKeyFigures, RangeForSelectedKeywords);

            // sales multiplier
            var SalesMultFields = TableFields.CreateEmptyTableFields(md);
            TableFields.AddNewField(SalesMultFields, "dealer");
            TableFields.AddNewField(SalesMultFields, "sales");

            var SalesMult = MatrixTable.CreateTableWithElements_A(SalesMultFields,
                "Merlin", 0.3,
                "Belinda", 0.8,
                "Timon", 1.5,
                "Proda", 1.2,
                "Maximus", 1.3,
                "Tinkerbell", 0.5
                );

            SalesTable = SalesTable * SalesMult;

            // round all key figures to 2 digits after decimal point
            SalesTable = MatrixTable.TransformKeyFigures(SalesTable, x => Math.Round(x, 2),
                InputKeyFig: null, OutputKeyFig: null);

            // add category to table
            SalesTable = MatrixTable.CombineTables(SalesTable, PrdCategoryTable);

            // view SalesTable
            // MatrixTable.View_MatrixTable(SalesTable, "Sales Table");

            // ScaleTable
            var ScaleTableFields = TableFields.CreateEmptyTableFields(md);
            TableFields.AddNewField(ScaleTableFields, "scale_id");
            TableFields.AddNewField(ScaleTableFields, "lower_limit");
            TableFields.AddNewField(ScaleTableFields, "commission_rate");

            ComScaleTable = MatrixTable.CreateTableWithElements_A(ScaleTableFields,
                    1, 0.0, 0.10,
                    1, 1000.0, 0.20,
                    2, 0.0, 0.05,
                    2, 500.0, 0.10,
                    2, 2000.0, 0.20,
                    3, 0.0, 0.02,
                    3, 800.0, 0.03,
                    3, 1600.0, 0.04
                );

            // view ComScaleTable
            // MatrixTable.View_MatrixTable(ComScaleTable, "Commission Scales");

            // DealerScale table
            var ScalePoolFields = TableFields.CreateEmptyTableFields(md);

            TableFields.AddNewField(ScalePoolFields, "scale_logic");
            TableFields.AddNewField(ScalePoolFields, "pool_id");
            TableFields.AddNewField(ScalePoolFields, "scale_id");

            ScalePoolTable = MatrixTable.CreateTableWithElements_A(ScalePoolFields,
                    "level", 11, 1,
                    "level", 22, 2,
                    "class", 33, 3,
                    "class", 44, 2
                );

            // view ScalePoolTable
            // MatrixTable.View_MatrixTable(ScalePoolTable, "Commission Scale ID assigned to each Product Pool");

            // PoolTable
            var PoolTableFields = TableFields.CreateEmptyTableFields(md);
            TableFields.AddNewField(PoolTableFields, "dealer");
            TableFields.AddNewField(PoolTableFields, "category");
            TableFields.AddNewField(PoolTableFields, "product");
            TableFields.AddNewField(PoolTableFields, "pool_id");

            PoolTable = MatrixTable.CreateTableWithElements_A(PoolTableFields,
                "Merlin", "ALL", "ALL", 11,
                "Belinda", "Sports", "ALL", 22,
                "Belinda", "Luxury", "ALL", 33,
                "Belinda", "Economy", "ALL", 44,
                "Timon", "Sports", "ALL", 11,
                "Timon", "ALL", "ALL", 33,
                "Proda", "Sports", "ALL", 11,
                "Proda", "Luxury", "ALL", 22,
                "Proda", "Economy", "BMW budget", 11,
                "Proda", "Economy", "ALL", 44,
                "Maximus", "Luxury", "ALL", 11,
                "Maximus", "ALL", "ALL", 33,
                "Tinkerbell", "Luxury", "ALL", 22,
                "Tinkerbell", "ALL", "ALL", 44
                );

            // view PoolTable
            // MatrixTable.View_MatrixTable(PoolTable, "Definition of Product Pools");
        }

    }
}

