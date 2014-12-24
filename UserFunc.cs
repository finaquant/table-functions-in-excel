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

// Recommended steps for adding a user-defined method/macro to Excel
/* 1) Write your user-defined method here as a member of class UserFunc
 * 2) Add new non-static methods to class ExcelTable that call your method here
 * 3) Add new static methods to class ExcelTableDNA with attributes to make
 * your method available in excel
 */

namespace FinaquantInExcel
{
    /// <summary>
    /// Class reserved for user-defined table functions.
    /// </summary>
    public class UserFunc
    {
        /// <summary>
        /// Get price table with costs, margins and prices according to given cost and margin tables.
        /// </summary>
        /// <param name="CostTable">Cost table with key figure named "costs"</param>
        /// <param name="MarginTable">Margin table with key figure named "margin"</param>
        /// <returns>Price table</returns>
        public static MatrixTable GetPriceTable(MatrixTable CostTable, MatrixTable MarginTable)
        {
            // check inputs
            if (CostTable == null || CostTable.IsEmpty || CostTable.RowCount == 0)
                throw new Exception("GetPriceTable: Null or empty input table CostTable\n");

            if (MarginTable == null || MarginTable.IsEmpty || MarginTable.RowCount == 0)
                throw new Exception("GetPriceTable: Null or empty input table MarginTable\n");

            // check if CostTable has a key figure named "costs"
            if (! TextVector.IfValueFoundInSet("costs", CostTable.KeyFigureFields))
                throw new Exception("GetPriceTable: CostTable must have a key figure named costs\n");

            // check if MarginTable has a key figure named "margin"
            if (!TextVector.IfValueFoundInSet("margin", MarginTable.KeyFigureFields))
                throw new Exception("GetPriceTable: MarginTable must have a key figure named margin\n");

            // input checks OK, continue..
            try
            {
                // CostTable: Exclude all key figures other than "costs" (only "costs" is required)
                MatrixTable CostTbl = MatrixTable.ExcludeColumns(CostTable,
                    TextVector.SetDifference(CostTable.KeyFigureFields,
                    new TextVector(new string[] { "costs" })));

                // MarginTable: Exclude all key figures other than "margin" (only "margin" is required)
                MatrixTable MarginTbl = MatrixTable.ExcludeColumns(MarginTable,
                    TextVector.SetDifference(MarginTable.KeyFigureFields,
                    new TextVector(new string[] { "margin" })));

                // Get Price Table
                MatrixTable PriceTbl = MatrixTable.MultiplySelectedKeyFigures(
                    CostTbl, MarginTbl + 1, "costs", "margin", "price", JokerMatchesAllvalues: true);

                // Combine MarginTbl to see margins in resultant price table
                PriceTbl = MatrixTable.CombineTables(PriceTbl, MarginTbl);

                return PriceTbl;
            }
            catch (Exception ex)
            {
                throw new Exception("GetPriceTable: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Calculate sales commissions per product pool (group) and dealer with tiered commission rates.
        /// See related article: http://finaquant.com/commission-calculation-with-finaquant-calcs/3607
        /// </summary>
        /// <param name="Period">Calculation and payment period, 'month' or 'quarter'</param>
        /// <param name="SalesTable">Input table with with sales transaction data for each dealer</param>
        /// <param name="ComScaleTable">Input table with tiered rates for each commission scale</param>
        /// <param name="ScalePoolTable">Input table that maps commission scales to product pools, with additional scale logic (class or level)</param>
        /// <param name="PoolTable">Input table that maps categories and products to a product pool for each dealer</param>
        /// <param name="CommissionsPerPool">Output table: Sales Commissions per Product Pool</param>
        /// <param name="CommissionsPerDealer">Output table: Sales Commissions per Dealer</param>
        public static void CalculateSalesCommissions(string Period, MatrixTable SalesTable,
            MatrixTable ComScaleTable, MatrixTable ScalePoolTable, MatrixTable PoolTable,
            out MatrixTable CommissionsPerPool, out MatrixTable CommissionsPerDealer)
        {
            string MethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;

            #region "input checks"

            // period
            Period = Period.ToLower();
            if (Period != "month" && Period != "quarter")
                throw new Exception(MethodName + ": Invalid Period value which must be either 'month' or 'quarter'.");
            
            // sales table
            if (SalesTable == null || SalesTable.IsEmpty || SalesTable.RowCount == 0)
                throw new Exception(MethodName + ": Null-valued or empty input sales table!");

            // check if sales table contains fields 'sales', 'date', 'dealer', 'product'
            TextVector SalesTblFields = SalesTable.ColumnNames;

            if (! TextVector.IfV2containsV1(new TextVector(new string[] {"sales", "date", "dealer", "product"}),
                        SalesTblFields))
            {
                throw new Exception(MethodName + ": Input Sales Table must contain fields 'sales', 'date', 'dealer', 'product'.");
            }

            // pool table
            if (PoolTable == null || PoolTable.IsEmpty || PoolTable.RowCount == 0)
                throw new Exception(MethodName + ": Null-valued or empty input product pool table!");

            // check if pool table contains field 'pool_id'
            TextVector PoolTblFields = PoolTable.ColumnNames;

            if (!TextVector.IfV2containsV1(new TextVector(new string[] { "pool_id", "product" }), 
                PoolTblFields))
            {
                throw new Exception(MethodName + ": Input Product Pool Table must contain fields 'pool_id' and 'product'.");
            }

            // check if pool table fields (excluding pool_id) are a subset of sales table fields
            if (! TextVector.IfV2containsV1(TextVector.SetDifference(PoolTblFields, 
                new TextVector(new string[] { "pool_id"})), SalesTblFields))
            {
                throw new Exception(MethodName + ": All fields of Product Pool Table (excluding pool_id) must be contained in Sales Table!");
            }

            // commission scale table
            if (ComScaleTable == null || ComScaleTable.IsEmpty || ComScaleTable.RowCount == 0)
                throw new Exception(MethodName + ": Null-valued or empty input commission scale table!");

            // check fields of commission scale table
            TextVector ComScaleTblFields = ComScaleTable.ColumnNames;
            
            if (! TextVector.IsEqual(ComScaleTblFields, 
                new TextVector(new string[] { "scale_id", "lower_limit", "commission_rate"}), true))
            {
                throw new Exception(MethodName + ": Commission Scale Table must contain fields 'scale_id', 'lower_limit', 'commission_rate'.");
            }

            // pool to scale table
            if (ScalePoolTable == null || ScalePoolTable.IsEmpty || ScalePoolTable.RowCount == 0)
                throw new Exception(MethodName + ": Null-valued or empty input pool-to-scale table!");
            
            // check fields of commission scale table
            TextVector ScalePoolTblFields = ScalePoolTable.ColumnNames;

            if (!TextVector.IsEqual(ScalePoolTblFields,
                new TextVector(new string[] { "scale_logic", "pool_id", "scale_id" }), true))
            {
                throw new Exception(MethodName + ": Pool-To-Scale Table must contain fields 'scale_logic', 'pool_id', 'scale_id'.");
            }

            #endregion "input checks"

            // input checks OK, continue..
            try
            {
                MetaData md = SalesTable.metaData;

                // define date-related  attributes
                md.AddFieldIfNew("year", FieldType.IntegerAttribute);
                md.AddFieldIfNew("quarter", FieldType.IntegerAttribute);
                md.AddFieldIfNew("month", FieldType.IntegerAttribute);

                // define required key figures
                md.AddFieldIfNew("pooled_sales", FieldType.KeyFigure);
                md.AddFieldIfNew("interval_rate", FieldType.KeyFigure);
                md.AddFieldIfNew("effective_rate", FieldType.KeyFigure);
                md.AddFieldIfNew("commission", FieldType.KeyFigure);

                // add year to table
                MatrixTable SalesTableWithYear = MatrixTable.InsertDateRelatedAttribute(SalesTable, "date", "year", DateRelation.Year);

                // add period (month or quarter) to table
                MatrixTable SalesTableWithPeriod;
                TextVector RefAttrib;
                Period = Period.ToLower();

                if (Period == "month")
                {
                    SalesTableWithPeriod = MatrixTable.InsertDateRelatedAttribute(SalesTableWithYear, "date", "month", DateRelation.Month);
                    RefAttrib = TextVector.CreateVectorWithElements("dealer", "month", "year", "pool_id");
                }
                else if (Period == "quarter")
                {
                    SalesTableWithPeriod = MatrixTable.InsertDateRelatedAttribute(SalesTableWithYear, "date", "quarter", DateRelation.Quarter);
                    RefAttrib = TextVector.CreateVectorWithElements("dealer", "quarter", "year", "pool_id");
                }
                else
                    throw new Exception("Undefined period: " + Period);

                // add attribute pool_id to table
                MatrixTable SalesTableWithPool = MatrixTable.CombineTables(SalesTableWithPeriod, PoolTable,
                    JokerMatchesAllvalues: true);

                // aggregate sales table w.r.t. pool
                TextVector SubTableFields = TextVector.CreateVectorWithElements("dealer", "year", Period, "pool_id", "sales");
                SalesTableWithPool = SalesTableWithPool.PartitionColumn(SubTableFields);

                MatrixTable AggregatedSalesTableWithPool = MatrixTable.AggregateAllKeyFigures(SalesTableWithPool, null);

                // add attributes scale_id & scale_logic to AggregatedSalesTableWithPool
                MatrixTable AggregatedSalesWithScale = MatrixTable.CombineTables(AggregatedSalesTableWithPool, ScalePoolTable,
                    JokerMatchesAllvalues: true);

                // get scale table for each scale_id
                SubTableFields = TextVector.CreateVectorWithElements("lower_limit", "commission_rate");
                MatrixTable UniqueScaleIds;
                MatrixTable[] ScaleTables;
                MatrixTable.GetSubTables(ComScaleTable, SubTableFields, out UniqueScaleIds, out ScaleTables);

                // Table array indexed with a dictionary (assoc array indexed with scale_id)
                var ScaleTableDic = new Dictionary<int, MatrixTable>();

                for (int i = 0; i < UniqueScaleIds.RowCount; i++)
                {
                    ScaleTableDic[(int)UniqueScaleIds.GetFieldValue("scale_id", i)] = ScaleTables[i];
                }

                // insert new key figures to table
                AggregatedSalesWithScale = MatrixTable.InsertNewColumn(AggregatedSalesWithScale, "interval_rate", 0.0);
                AggregatedSalesWithScale = MatrixTable.InsertNewColumn(AggregatedSalesWithScale, "effective_rate", 0.0);
                AggregatedSalesWithScale = MatrixTable.InsertNewColumn(AggregatedSalesWithScale, "commission", 0.0);

                // row-by-row processing of table for calculating commission rates
                var TextAttribDic = new Dictionary<string, string>();
                var NumAttribDic = new Dictionary<string, int>();
                var KeyFigDic = new Dictionary<string, double>();

                CommissionsPerPool = MatrixTable.TransformRowsDic(AggregatedSalesWithScale,
                     CalCommissionForEachRow, ScaleTableDic);

                // calculate total commission per dealer for each period (quarter or month)
                SubTableFields = TextVector.CreateVectorWithElements("dealer", "year", Period, "sales", "commission");

                CommissionsPerDealer = CommissionsPerPool.PartitionColumn(SubTableFields);
                CommissionsPerDealer = MatrixTable.AggregateAllKeyFigures(CommissionsPerDealer, null);

                // round sales & commission amounts to two digits after decimal point
                CommissionsPerDealer = MatrixTable.Round(CommissionsPerDealer, 2);
            }
            catch (Exception ex)
            {
                throw new Exception(MethodName + ": " + ex.Message + "\n");
            }
        }

        // Helper function for row-by-row processing: Commission Calculation
        private static void CalCommissionForEachRow(ref Dictionary<string, string> TextAttribDic,
        ref Dictionary<string, int> NumAttribDic, ref Dictionary<string, double> KeyFigDic,
        params Object[] OtherParameters)
        {
            try
            {
                // get parameters
                var ScaleTableDic = (Dictionary<int, MatrixTable>)OtherParameters[0];

                // get field values from table row
                int scale_id = NumAttribDic["scale_id"];
                double sales = KeyFigDic["sales"];
                string scale_logic = TextAttribDic["scale_logic"];

                // calculate commission rates
                double interval_rate, effective_rate, commission;

                CalculateCommission(ScaleTableDic, scale_id, sales, scale_logic,
                out interval_rate, out effective_rate, out commission);

                // assign values to fields
                KeyFigDic["interval_rate"] = interval_rate;
                KeyFigDic["effective_rate"] = effective_rate;
                KeyFigDic["commission"] = commission;
            }
            catch (Exception ex)
            {
                throw new Exception("ERROR in CalCommissionForEachRow: " + ex.Message);
            }
        }

        // Helper function for calculating commission rates
        private static void CalculateCommission(Dictionary<int, MatrixTable> ScaleTableDic, int scale_id, double sales, string scale_logic,
            out double interval_rate, out double effective_rate, out double commission)
        {
            // subtable with scale_id only
            MatrixTable ScaleSubTable = ScaleTableDic[scale_id];

            // get scale matrix with lower_limit & rate
            KeyMatrix ScaleMatrix = ScaleSubTable.KeyFigValues;

            // get scale logic
            bool IfLevelLogic;

            if (scale_logic.ToLower() == "level")
                IfLevelLogic = true;
            else if (scale_logic.ToLower() == "class")
                IfLevelLogic = false;
            else
                throw new Exception("ERROR in CalculateCommission: Unknown scale logic " + scale_logic);

            // calculate effective rate & commission
            AppUtility.CalcEffectiveRate(ScaleMatrix, sales, out interval_rate, out effective_rate, out commission,
                ApplyLevelLogic: IfLevelLogic);
        }

    }
}
