// Finaquant Analytics - http://finaquant.com/
// Copyright © Finaquant Analytics GmbH
// Email: support@finaquant.com

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
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

namespace FinaquantInExcel
{

    /// <summary>
    /// Class for central meta data: fields and hierarchies for all tables.
    /// </summary>
    /// <remarks>
    /// MetaDataX is a class which envelopes finaquant's MetaData class, and 
    /// makes only some of MetaData's properties and methods available for excel
    /// users.
    /// </remarks>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("Finaquant_MetaDataX")]
    public class MetaDataX
    {
        private MetaData _md = null;

        #region "MetaDataX constructors and creators"

        /// <summary>
        /// Default constructor which creates an empty meta data object
        /// without any fields. Fields can be added subsequently.
        /// </summary>
        public MetaDataX()
        {
            this._md = MetaData.CreateEmptyMetaData();
        }

        /// <summary>
        /// Constructor with MetaData input
        /// </summary>
        /// <param name="md">Meta data object</param>
        internal MetaDataX(MetaData md)
        {
            this._md = md;
        }

        /// <summary>
        /// Create an empty meta data object without any fields.
        /// Fields can be added subsequently.
        /// </summary>
        /// <returns>Meta data object</returns>
        internal static MetaDataX CreateEmptyMetaDataX()
        {
            return new MetaDataX();
        }

        /// <summary>
        /// Read field names and types from an excel table.
        /// Excel table must have 2 columns for field names and types:
        /// text, integer, date, keyfig (case insensitive)
        /// </summary>
        /// <param name="xtbl">Excel table</param>
        /// <remarks>
        /// All field names are stored in small letters in meta data object.
        /// </remarks>
        public void ReadFieldsFromExcelTable(Excel.ListObject xtbl)
        {
            // check input table
            if (xtbl == null)
                throw new Exception("MetaDataX.ReadFieldsFromExcelTable: Null valued input table!\n");

            if (xtbl.ListColumns.Count != 2)
                throw new Exception("MetaDataX.ReadFieldsFromExcelTable: Excel table must have two columns: Field name and type.\n");

            if (xtbl.HeaderRowRange == null)
                throw new Exception("MetaDataX.ReadFieldsFromExcelTable: Excel table must have a header row with column names.\n");

            if (xtbl.ListRows.Count == 0)
                throw new Exception("MetaDataX.ReadFieldsFromExcelTable: Excel table has no rows!\n");

            // input checks done, continue..

            // init variables
            string fieldname, fieldtype;
            object CellValue;
            string StrPat = @"^\s*([\w\-_]+)\s*$";
            var mdata = MetaData.CreateEmptyMetaData();

            //try
            //{
                // get table size
                int nrows = xtbl.ListRows.Count;
                int ncols = xtbl.ListColumns.Count;

                // add fields to meta data
                for (int i = 1+1; i <= nrows+1; i++)
                {
                    // read field name
                    // CellValue = ((Excel.Range)xtbl.Range.get_Item(i, 1)).Value2;
                    CellValue = ((Excel.Range)xtbl.Range[i, 1]).Value2;

                    fieldname = CellValue.ToString().ToLower();

                    // check field name
                    Match m = Regex.Match(fieldname, StrPat);
                    if (!m.Success)
                        throw new Exception("Field name '" + fieldname + "' does not match the required string pattern!\n");

                    // read field type
                    // CellValue = ((Excel.Range)xtbl.Range.get_Item(i, 2)).Value2;
                    CellValue = ((Excel.Range)xtbl.Range[i, 2]).Value2;

                    fieldtype = CellValue.ToString().ToLower();

                    // check field type
                    m = Regex.Match(fieldname, StrPat);
                    if (!m.Success)
                        throw new Exception("Field type '" + fieldtype + "' does not match the required string pattern!\n");

                    // add field to meta data
                    if (Regex.IsMatch(fieldtype, "text"))
                    {
                        try
                        {
                            mdata.AddNewField(fieldname, FieldType.TextAttribute);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("Field '" + fieldname +
                                "' could not be added to meta data!\n" + ex.Message + "\n");
                        }
                    }
                    else if (Regex.IsMatch(fieldtype, "integer"))
                    {
                        try
                        {
                            mdata.AddNewField(fieldname, FieldType.IntegerAttribute);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("Field '" + fieldname +
                                "' could not be added to meta data!\n" + ex.Message + "\n");
                        }
                    }
                    else if (Regex.IsMatch(fieldtype, "date"))
                    {
                        try
                        {
                            mdata.AddNewField(fieldname, FieldType.DateAttribute);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("Field '" + fieldname +
                                "' could not be added to meta data!\n" + ex.Message + "\n");
                        }
                    }
                    else if (Regex.IsMatch(fieldtype, "keyfig"))
                    {
                        try
                        {
                            mdata.AddNewField(fieldname, FieldType.KeyFigure);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("Field '" + fieldname +
                                "' could not be added to meta data!\n" + ex.Message + "\n");
                        }
                    }
                    else
                    {
                        throw new Exception("Improper field type value '" + fieldtype + "' in excel table!\n");
                    }

                }   // for i

                // return meta data object
                this._md = mdata;
            //}
            //catch (Exception ex)
            //{
            //    throw new Exception("MetaDataX.ReadFieldsFromExcelTable: " + ex.Message);
            //}
        }

        /// <summary>
        /// Read field names and types from an excel table.
        /// Excel table must have 2 columns for field names and types:
        /// text, integer, date, keyfig (case insensitive)
        /// </summary>
        /// <param name="wbook">Excel Workbook object</param>
        /// <param name="xMetaTblName">Name of ListObject (excel table)</param>
        public void ReadFieldsFromExcelTable2(Excel.Workbook wbook, string xMetaTblName)
        {
            if (xMetaTblName == null && xMetaTblName == "")
                throw new Exception("MetaDataX.ReadFieldsFromExcelTable2: Null or empty string xMetaTblName!\n");

            // Find excel table (ListObject) by its name and return its ListObject instance
            Excel.ListObject xTblMeta = ExcelFunc_NO.GetListObject(wbook, xMetaTblName);

            // Read field definitions from ListObject
            ReadFieldsFromExcelTable(xTblMeta);
        }

        public void ReadFieldsFromExcelTable9(ref object wbook, ref string xMetaTblName)
        {
            // TEST
            MessageBox.Show("wbook type: " + wbook.GetType());
        }

        /// <summary>
        /// Read field names and types from an excel table.
        /// Excel table must have 2 columns for field names and types:
        /// text, integer, date, keyfig (case insensitive)
        /// </summary>
        /// <param name="WorkbookPath">Full file path of excel Workbook</param>
        /// <param name="xMetaTblName">Name of ListObject (excel table)</param>
        public void ReadFieldsFromExcelTable3(string WorkbookPath, string xMetaTblName)
        {
            if (WorkbookPath == null)
                throw new Exception("MetaDataX.ReadFieldsFromExcelTable3: Null-valued WorkbookPath!\n");

            // check if workbook is open
            Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();
            string filename = Path.GetFileName(WorkbookPath);
            Excel.Workbook wbook;

            if (ExcelFunc_NO.IsExcelWBOpen2(xlapp, WorkbookPath))
            {
                // get workbook object
                wbook = xlapp.Workbooks[filename];
            }
            else
            {
                throw new Exception("MetaDataX.ReadFieldsFromExcelTable3: A workbook with the given path is not opened!\n");
            }

            //TEST
            MessageBox.Show("Workbook is found, now read field names and types from excel");

            ReadFieldsFromExcelTable2((Excel.Workbook) wbook, xMetaTblName);
        }

        /// <summary>
        /// Read meta data (fields and hierarchies) from XML file
        /// </summary>
        /// <param name="Filename">A file name like "InventoryTable" without file extension</param>
        /// <param name="FileDir">A valid directory like @"C:\Users\John\Documents\</param>
        public void ImportFromXMLfile(string Filename, string FileDir = "")
        {
            this._md = MetaData.ImportFromXMLfile(Filename, FileDir);
        }

        #endregion "MetaDataX constructors and creators"

        #region "MetaDataX properties"

        /// <summary>
        /// Get underlying meta data object
        /// </summary>
        internal MetaData metaData
        {
            get { return this._md; }
        }

        /// <summary>
        /// Get underlying meta data instance (object)
        /// </summary>
        public object MetaDataObj
        {
            get { return this._md; }
        }

        /// <summary>
        /// Return true if no field is registered yet
        /// </summary>
        public bool IsEmpty
        {
            get { return this._md.IsEmpty; }
        }

        /// <summary>
        /// Return number of fields registered in meta data
        /// </summary>
        public int FieldCount
        {
            get { return this._md.FieldCount; }
        }

        /// <summary>
        /// Return number of hierarchy tables in meta data
        /// </summary>
        public int HierarchyCount
        {
            get { return this._md.HierarchyCount; }
        }

        /// <summary>
        /// Return array of field names for each hierarchy
        /// </summary>
        public string[][] HierarchyFields
        {
            get 
            {
                try
                {
                    TextVector[] fieldvecs = this._md.HierarchyFields;

                    var fieldarr = new string[fieldvecs.Count()][];

                    for (int i = 0; i < fieldvecs.Count(); i++)
                    {
                        fieldarr[i] = fieldvecs[i].toArray;
                    }
                    return fieldarr;
                }
                catch (Exception ex)
                {
                    throw new Exception("MatrixTableX.HierarchyFields: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Return all text attribute fields
        /// </summary>
        public string[] TextAttributes
        {
            get { return this._md.TextAttributes.toArray; }
        }

        /// <summary>
        /// Return all date attribute fields
        /// </summary>
        public string[] DateAttributes
        {
            get { return this._md.DateAttributes.toArray; }
        }

        /// <summary>
        /// Return all numeric attribute fields
        /// </summary>
        /// <remarks>
        /// Note: Date and Integer attributes are both numeric attributes.
        /// </remarks>
        public string[] IntegerAttributes
        {
            get { return this._md.IntegerAttributes.toArray; }
        }

        /// <summary>
        /// Return all key figure fields
        /// </summary>
        public string[] KeyFigures
        {
            get { return this._md.KeyFigures.toArray; }
        }

        /// <summary>
        /// Return all fields defined in MetaData
        /// </summary>
        public string[] AllFields
        {
            get { return this._md.AllFields.toArray; }
        }

        /// <summary>
        /// Return a 2-column table with field names and types
        /// </summary>
        public MatrixTableX FieldTable
        {
            get
            {
                // define table structure
                var TblFields = new TableFields(this._md);
                TblFields.AddNewField("field_names", FieldType.TextAttribute);
                TblFields.AddNewField("field_types", FieldType.TextAttribute);

                // type to string dictionary
                var typedic = new Dictionary<FieldType, string>();
                typedic.Add(FieldType.DateAttribute, "date");
                typedic.Add(FieldType.IntegerAttribute, "integer");
                typedic.Add(FieldType.TextAttribute, "text");
                typedic.Add(FieldType.KeyFigure, "keyfig");

                // create text matrix for values
                TextMatrix FieldsTypes = TextMatrix.CreateConstantMatrix(this.FieldCount, 2);
                TextVector AllFields = this._md.AllFields;

                for (int i = 0; i < AllFields.nLength; i++)
                {
                    FieldsTypes[i, 0] = AllFields[i];
                    FieldsTypes[i, 1] = typedic[this._md.GetFieldType(AllFields[i])];
                }

                var ftbl = MatrixTable.CreateTableWithMatrices_A(TblFields,
                    FieldsTypes, null, null);

                return new MatrixTableX(ftbl);
            }
        }

        #endregion "MetaDataX properties"

        #region "MetaDataX methods"

        /// <summary>
        /// Add new field definition to meta data.
        /// </summary>
        /// <param name="fieldname">Field name, case insensitive</param>
        /// <param name="fieldtype">Field type (text, integer, date, keyfig), case insensitive</param>
        public void AddNewField(string fieldname, string fieldtype)
        {
            // check inputs
            if (fieldname == null || fieldname == "")
                throw new Exception("MetaDataX.AddNewField: Null or empty FieldName!");

            if (fieldtype == null || fieldtype == "")
                throw new Exception("MetaDataX.AddNewField: Null or empty FieldType!");

            FieldType ft;

            try
            {
                // check field type
                string name = fieldname.ToLower();
                string type = fieldtype.ToLower();

                if (type == "text") ft = FieldType.TextAttribute;
                else if (type == "integer") ft = FieldType.IntegerAttribute;
                else if (type == "date") ft = FieldType.DateAttribute;
                else if (type == "keyfig") ft = FieldType.KeyFigure;
                else // undefined field type
                    throw new Exception("MetaDataX.AddNewField: Undefined field type '" + "fieldtype'!");

                // call method to add the field
                this._md.AddNewField(fieldname, ft);
            }
            catch (Exception ex)
            {
                throw new Exception("MetaDataX.AddNewField: " + ex.Message);
            }
        }

        /// <summary>
        /// Return field name with the given position index
        /// </summary>
        public string GetFieldName(int FieldIndex)
        {
            return MetaData.GetFieldName(this._md, FieldIndex - 1); 
        }

        /// <summary>
        /// Return true if field is found in meta data
        /// </summary>
        public bool IfFieldExists(string FieldName)
        {
            return MetaData.IfFieldExists(this._md, FieldName);
        }

        /// <summary>
        /// Check if there is a field with the given type in meta data
        /// </summary>
        /// <param name="fieldname">Field name</param>
        /// <param name="fieldtype">Field type (text, integer, date, keyfig)</param>
        /// <param name="warning">informative warning message</param>
        /// <returns>True if field name with the given type is found in meta data</returns>
        public bool CheckFieldAndType(string fieldname, string fieldtype, out string warning)
        {
            // check inputs
            if (fieldname == null || fieldname == "")
                throw new Exception("MetaDataX.CheckFieldAndType: Null or empty FieldName!");

            if (fieldtype == null || fieldtype == "")
                throw new Exception("MetaDataX.CheckFieldAndType: Null or empty FieldType!");

            FieldType ft;

            // check field type
            string name = fieldname.ToLower();
            string type = fieldtype.ToLower();

            if (type == "text") ft = FieldType.TextAttribute;
            else if (type == "integer") ft = FieldType.IntegerAttribute;
            else if (type == "date") ft = FieldType.DateAttribute;
            else if (type == "keyfig") ft = FieldType.KeyFigure;
            else // undefined field type
                throw new Exception("MetaDataX.CheckFieldAndType: Undefined field type '" + "fieldtype'!");

            return MetaData.CheckFieldAndType(this._md,fieldname, ft, out warning);
        }

        /// <summary>
        /// Convert meta data into printable string
        /// </summary>
        public override string ToString()
        {
            return this._md.ToString();
        }

        /// <summary>
        /// Write (export) meta data into an XML file
        /// </summary>
        /// <param name="Filename">A file name like "InventoryTable" without file extension</param>
        /// <param name="FileDir">A valid directory like @"C:\Users\John\Documents\</param>
        /// <param name="IfAddTimeStamp">If true, add timestamp to file name </param>
        /// <remarks>
        /// - first convert meta data to DataSet, than write DataSet into an XML file
        /// - error if meta data is empty; i.e. no field definitions
        /// </remarks>
        public void ExportToXMLfile(string Filename, string FileDir = "", bool IfAddTimeStamp = true)
        {
            MetaData.ExportToXMLfile(this._md, Filename, FileDir, IfAddTimeStamp);
        }

        /// <summary>TODO
        /// Add new hierarchy table with text and numeric attributes to meta data.
        /// </summary>
        /// <param name="OrderedHierarchyFields">Ordered hierarchy attributes from highest to lowest</param>
        /// <param name="HierarchyTable">Hierarchy table with text and numeric attributes</param>
        /// <remarks>
        public void AddNewHierarchy(string[] OrderedHierarchyFields, Excel.ListObject HierarchyTable)
        {

        }

        #endregion "MetaDataX methods"

    }   // class MetaDataX

}   // namespace
