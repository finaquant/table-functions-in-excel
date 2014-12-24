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

namespace FinaquantInExcel
{
    /// <summary>
    /// Class with application oriented matrix functions
    /// for excel users and VBA programmers. 
    /// Input and Output parameters are generally named ranges for matrices.
    /// All public and non-static methods in this class are available in Excel VBA.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("Finaquant_ExcelMatrix")]
    public class ExcelMatrix
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public ExcelMatrix() { }

        /// <summary>
        /// Read matrix from a named Excel range
        /// </summary>
        /// <param name="RangeName">Name of ran^ge</param>
        /// <param name="WorkbookPath">Full path of Excel Workbook</param>
        /// <param name="KeyFigReplaceNull">Replace null or empty values in Excel with this value</param>
        /// <returns>Matrix</returns>
        public MatrixTableX ExcelToMatrix(string RangeName, string WorkbookPath = null, 
            double KeyFigReplaceNull = 0)
        {
            KeyMatrix M = ExcelFunc_NO.ReadKeyMatrixFromExcel(RangeName, null, KeyFigReplaceNull);
            return new MatrixTableX(M);
        }

        /// <summary>
        /// Write Matrix to an Excel range
        /// </summary>
        /// <param name="wbook">Excel workbook</param>
        /// <param name="M">Key Figure Matrix (KeyMatrix)</param>
        /// <param name="SheetName">Name of worksheet</param>
        /// <param name="TopLeftCell">Upper-left corner of matrix in excel sheet</param>
        /// <param name="RangeName">Name to be assigned to the range of matrix</param>
        public void MatrixToExcel(Excel.Workbook wbook, KeyMatrixX M, string SheetName, string TopLeftCell = "A1",
            string RangeName = null)
        {
            ExcelFunc_NO.WriteKeyMatrixToExcel(wbook, M.keymatrix, SheetName, TopLeftCell, RangeName);
        }

        /// <summary>
        /// Create a NxM matrix with random valued elements
        /// </summary>
        /// <param name="nRows">N</param>
        /// <param name="nCols">M</param>
        /// <param name="RangeName">Name of matrix range in Excel</param>
        /// <param name="WorkbookFullName">Full workbook path; if null, active workbook is assumed</param>
        /// <param name="TargetSheetName">Name of the sheet in which the matrix is inserted</param>
        /// <param name="TopLeftCell">Upper left cell of matrix in target sheet</param>
        public void CreateRandomMatrix(int nRows, int nCols, string RangeName, 
            string WorkbookFullName = null, string TargetSheetName = "ResultMatrix", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Read excel ranges into matrices

            // Step 2: Generate resultant (output) matrices
            KeyMatrix M = KeyMatrix.CreateRandomMatrix(nRows, nCols, 100);

            // Step 3: Write resultant matrices to excel
            ExcelFunc_NO.WriteKeyMatrixToExcel(wbook, M, TargetSheetName, TopLeftCell, RangeName);
            
            ((Excel.Worksheet)wbook.Worksheets[TargetSheetName]).Activate();
        }

        /// <summary>
        /// Create a NxM matrix with random valued elements
        /// </summary>
        public void CreateRandomMatrix_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for CreateRandomMatrix";

                string prompt = "Please select a single-column range with 5 input values for following parameters:\n\n"
                    + "1) Row count\n"
                    + "2) Column count\n"
                    + "3) Name of matrix range\n"
                    + "4) Sheet name\n"
                    + "5) Upper-Left cell address";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 5, out ParameterValues))
                {
                    // assign parameter values
                    int nRows = Convert.ToInt32(ParameterValues[0].ToString());
                    int nCols = Convert.ToInt32(ParameterValues[1].ToString());
                    string RangeName = (string)ParameterValues[2];
                    string SheetName = (string)ParameterValues[3];
                    string UpperLeftCell = (string)ParameterValues[4];

                    CreateRandomMatrix(nRows, nCols, RangeName, null, SheetName, UpperLeftCell);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("CreateRandomMatrix: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Create a NxM matrix with random valued elements
        /// </summary>
        public void CreateRandomMatrix_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;

                // form parameters
                string FormTitle = "Parameter Form for CreateRandomMatrix";
                string FuncTitle = "Create Random Matrix";
                string FuncDescr = "Create a NxM matrix with random valued elements (with values between 0 and 1), where N is row count, and M is column count.";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_CreateRandomMatrix(FormTitle, FuncTitle, FuncDescr, "Matrix1", "Matrix1", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                int RowCount = int.Parse(ParameterForm1.txtBox1_st);
                int ColumnCount = int.Parse(ParameterForm1.txtBox2_st);
                string RangeName = ParameterForm1.OutMatrix1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                // call matrix function
                CreateRandomMatrix(RowCount, ColumnCount, RangeName, null, SheetName, TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("CreateRandomMatrix: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Multiply Matrices (matrix multiplication in linear algebra)
        /// </summary>
        /// <param name="InRangeName1">Range name of 1. input matrix</param>
        /// <param name="InRangeName2">Range name of 1. input matrix</param>
        /// <param name="OutRangeName">Range name of output matrix</param>
        /// <param name="WorkbookFullName">Full workbook path; if null, active workbook is assumed</param>
        /// <param name="TargetSheetName">Name of the sheet in which the matrix is inserted</param>
        /// <param name="TopLeftCell">Upper left cell of matrix in target sheet</param>
        public void MultiplyMatrices(string InRangeName1, string InRangeName2, string OutRangeName,
            string WorkbookFullName = null, string TargetSheetName = "ResultMatrix", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Read excel ranges into input matrices
            KeyMatrix M1 = ExcelFunc_NO.ReadKeyMatrixFromExcel(InRangeName1, WorkbookFullName);
            KeyMatrix M2 = ExcelFunc_NO.ReadKeyMatrixFromExcel(InRangeName2, WorkbookFullName);

            // Step 2: Generate resultant (output) matrices
            KeyMatrix R = KeyMatrix.MultiplyMatrices(M1, M2);

            // Step 3: Write resultant matrices to excel
            ExcelFunc_NO.WriteKeyMatrixToExcel(wbook, R, TargetSheetName, TopLeftCell, OutRangeName);

            ((Excel.Worksheet)wbook.Worksheets[TargetSheetName]).Activate();
        }

        /// <summary>
        /// Multiply Matrices; matrix multiplication in linear algebra
        /// </summary>
        public void MultiplyMatrices_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for MultiplyMatrices";

                string prompt = "Please select a single-column range with 5 input values for following parameters:\n\n"
                    + "1) Range name 1. input matrix\n"
                    + "2) Range name 2. input matrix\n"
                    + "3) Range name output matrix\n"
                    + "4) Sheet name\n"
                    + "5) Upper-Left cell address";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 5, out ParameterValues))
                {
                    // assign parameter values
                    string InRangeName1 = (string)ParameterValues[0];
                    string InRangeName2 = (string)ParameterValues[1];
                    string RangeNameOut = (string)ParameterValues[2];
                    string SheetName = (string)ParameterValues[3];
                    string UpperLeftCell = (string)ParameterValues[4];

                    MultiplyMatrices(InRangeName1, InRangeName2, RangeNameOut, null, SheetName, UpperLeftCell);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("MultiplyMatrices: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Multiply Matrices; matrix multiplication in linear algebra
        /// </summary>
        public void MultiplyMatrices_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;

                // inbox parameters
                string title = "Parameter Form for MultiplyMatrices";

                // form parameters
                string FormTitle = "Parameter Form for MultiplyMatrices";
                string FuncTitle = "Multiply Matrices";
                string FuncDescr = "Matrix Multiplication in Linear Algebra: C = A x B where A is a NxM, and B is a MxP matrix. Resultant C is a NxP matrix.";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_BinaryMatrixOp(FormTitle, FuncTitle, FuncDescr, "MatrixC", "MatrixC", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string Matrix1 = ParameterForm1.Matrix1_st;
                string Matrix2 = ParameterForm1.Matrix2_st; 
                string OutMatrix = ParameterForm1.OutMatrix1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                MultiplyMatrices(Matrix1, Matrix2, OutMatrix, null, SheetName, TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("MultiplyMatrices: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Add Matrices; element-wise addition of two equal-sized matrices      
        /// </summary>
        /// <param name="InRangeName1">Range name of 1. input matrix</param>
        /// <param name="InRangeName2">Range name of 1. input matrix</param>
        /// <param name="OutRangeName">Range name of output matrix</param>
        /// <param name="WorkbookFullName">Full workbook path; if null, active workbook is assumed</param>
        /// <param name="TargetSheetName">Name of the sheet in which the matrix is inserted</param>
        /// <param name="TopLeftCell">Upper left cell of matrix in target sheet</param>
        public void AddMatrices(string InRangeName1, string InRangeName2, string OutRangeName,
            string WorkbookFullName = null, string TargetSheetName = "ResultMatrix", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Read excel ranges into input matrices
            KeyMatrix M1 = ExcelFunc_NO.ReadKeyMatrixFromExcel(InRangeName1, WorkbookFullName);
            KeyMatrix M2 = ExcelFunc_NO.ReadKeyMatrixFromExcel(InRangeName2, WorkbookFullName);

            // Step 2: Generate resultant (output) matrices
            KeyMatrix R = KeyMatrix.AddMatrices(M1, M2);

            // Step 3: Write resultant matrices to excel
            ExcelFunc_NO.WriteKeyMatrixToExcel(wbook, R, TargetSheetName, TopLeftCell, OutRangeName);

            ((Excel.Worksheet)wbook.Worksheets[TargetSheetName]).Activate();
        }

        /// <summary>
        /// Add Matrices; element-wise addition of two equal-sized matrices  
        /// </summary>
        public void AddMatrices_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for AddMatrices";

                string prompt = "Please select a single-column range with 5 input values for following parameters:\n\n"
                    + "1) Range name 1. input matrix\n"
                    + "2) Range name 2. input matrix\n"
                    + "3) Range name output matrix\n"
                    + "4) Sheet name\n"
                    + "5) Upper-Left cell address";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 5, out ParameterValues))
                {
                    // assign parameter values
                    string InRangeName1 = (string)ParameterValues[0];
                    string InRangeName2 = (string)ParameterValues[1];
                    string RangeNameOut = (string)ParameterValues[2];
                    string SheetName = (string)ParameterValues[3];
                    string UpperLeftCell = (string)ParameterValues[4];

                    AddMatrices(InRangeName1, InRangeName2, RangeNameOut, null, SheetName, UpperLeftCell);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("AddMatrices: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Add Matrices; element-wise addition of two equal-sized matrices  
        /// </summary>
        public void AddMatrices_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;

                // form parameters
                string FormTitle = "Parameter Form for AddMatrices";
                string FuncTitle = "Add Matrices";
                string FuncDescr = "Matrix Addition; Element-wise addition of two equal-sized matrices.";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_BinaryMatrixOp(FormTitle, FuncTitle, FuncDescr, "MatrixC", "MatrixC", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string Matrix1 = ParameterForm1.Matrix1_st;
                string Matrix2 = ParameterForm1.Matrix2_st;
                string OutMatrix = ParameterForm1.OutMatrix1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                AddMatrices(Matrix1, Matrix2, OutMatrix, null, SheetName, TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("AddMatrices: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Inverse Matrix
        /// </summary>
        /// <param name="InRangeName">Range name of input matrix</param>
        /// <param name="OutRangeName">Range name of output matrix</param>
        /// <param name="WorkbookFullName">Full workbook path; if null, active workbook is assumed</param>
        /// <param name="TargetSheetName">Name of the sheet in which the matrix is inserted</param>
        /// <param name="TopLeftCell">Upper left cell of matrix in target sheet</param>
        public void InverseMatrix(string InRangeName, string OutRangeName,
            string WorkbookFullName = null, string TargetSheetName = "ResultMatrix", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Read excel ranges into input matrices
            KeyMatrix M = ExcelFunc_NO.ReadKeyMatrixFromExcel(InRangeName, WorkbookFullName);

            // Step 2: Generate resultant (output) matrices
            KeyMatrix R = M.Inverse();

            // Step 3: Write resultant matrices to excel
            ExcelFunc_NO.WriteKeyMatrixToExcel(wbook, R, TargetSheetName, TopLeftCell, OutRangeName);

            ((Excel.Worksheet)wbook.Worksheets[TargetSheetName]).Activate();
        }

        /// <summary>
        /// Inverse Matrix
        /// </summary>
        public void InverseMatrix_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for InverseMatrix";

                string prompt = "Please select a single-column range with 4 input values for following parameters:\n\n"
                    + "1) Range name input matrix\n"
                    + "2) Range name output matrix\n"
                    + "3) Sheet name\n"
                    + "4) Upper-Left cell address";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 4, out ParameterValues))
                {
                    // assign parameter values
                    string InRangeName = (string)ParameterValues[0];
                    string RangeNameOut = (string)ParameterValues[1];
                    string SheetName = (string)ParameterValues[2];
                    string UpperLeftCell = (string)ParameterValues[3];

                    InverseMatrix(InRangeName, RangeNameOut, null, SheetName, UpperLeftCell);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("InverseMatrix: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Inverse Matrix
        /// </summary>
        public void InverseMatrix_macro2()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;

                // form parameters
                string FormTitle = "Parameter Form for InverseMatrix";
                string FuncTitle = "Inverse Matrix";
                string FuncDescr = "Get inverse of square (NxN) input matrix, such that A x inv(A) = I";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_SingleMatrixOp(FormTitle, FuncTitle, FuncDescr, "InvMatrix", "InvMatrix", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string Matrix1 = ParameterForm1.Matrix1_st;
                string OutMatrix = ParameterForm1.OutMatrix1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                InverseMatrix(Matrix1, OutMatrix, null, SheetName, TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("AddMatrices: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Transpose Matrix
        /// </summary>
        /// <param name="InRangeName">Range name of input matrix</param>
        /// <param name="OutRangeName">Range name of output matrix</param>
        /// <param name="WorkbookFullName">Full workbook path; if null, active workbook is assumed</param>
        /// <param name="TargetSheetName">Name of the sheet in which the matrix is inserted</param>
        /// <param name="TopLeftCell">Upper left cell of matrix in target sheet</param>
        public void TransposeMatrix(string InRangeName, string OutRangeName,
            string WorkbookFullName = null, string TargetSheetName = "ResultMatrix", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Read excel ranges into input matrices
            KeyMatrix M = ExcelFunc_NO.ReadKeyMatrixFromExcel(InRangeName, WorkbookFullName);

            // Step 2: Generate resultant (output) matrices
            KeyMatrix R = M.Transpose();

            // Step 3: Write resultant matrices to excel
            ExcelFunc_NO.WriteKeyMatrixToExcel(wbook, R, TargetSheetName, TopLeftCell, OutRangeName);

            ((Excel.Worksheet)wbook.Worksheets[TargetSheetName]).Activate();
        }

        /// <summary>
        /// Transpose Matrix
        /// </summary>
        public void TransposeMatrix_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for TransposeMatrix";

                string prompt = "Please select a single-column range with 4 input values for following parameters:\n\n"
                    + "1) Range name input matrix\n"
                    + "2) Range name output matrix\n"
                    + "3) Sheet name\n"
                    + "4) Upper-Left cell address";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 4, out ParameterValues))
                {
                    // assign parameter values
                    string InRangeName = (string)ParameterValues[0];
                    string RangeNameOut = (string)ParameterValues[1];
                    string SheetName = (string)ParameterValues[2];
                    string UpperLeftCell = (string)ParameterValues[3];

                    TransposeMatrix(InRangeName, RangeNameOut, null, SheetName, UpperLeftCell);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("TransposeMatrix: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Transpose Matrix
        /// </summary>
        public void TransposeMatrix_macro2() 
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;

                // form parameters
                string FormTitle = "Parameter Form for TransposeMatrix";
                string FuncTitle = "Transpose Matrix";
                string FuncDescr = "Get transposed matrix";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_SingleMatrixOp(FormTitle, FuncTitle, FuncDescr, "TransMatrix", "TransMatrix", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string Matrix1 = ParameterForm1.Matrix1_st;
                string OutMatrix = ParameterForm1.OutMatrix1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                TransposeMatrix(Matrix1, OutMatrix, null, SheetName, TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("TransposeMatrix: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Add a scalar value to all elements of matrix
        /// </summary>
        /// <param name="InRangeName">Range name of input matrix</param>
        /// <param name="x">Scalar value</param>
        /// <param name="OutRangeName">Range name of output matrix</param>
        /// <param name="WorkbookFullName">Full workbook path; if null, active workbook is assumed</param>
        /// <param name="TargetSheetName">Name of the sheet in which the matrix is inserted</param>
        /// <param name="TopLeftCell">Upper left cell of matrix in target sheet</param>
        public void AddScalarToMatrix(string InRangeName, double x, string OutRangeName,
            string WorkbookFullName = null, string TargetSheetName = "ResultMatrix", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Read excel ranges into input matrices
            KeyMatrix M1 = ExcelFunc_NO.ReadKeyMatrixFromExcel(InRangeName, WorkbookFullName);

            // Step 2: Generate resultant (output) matrices
            KeyMatrix R = KeyMatrix.AddScalarToMatrix(M1, x);

            // Step 3: Write resultant matrices to excel
            ExcelFunc_NO.WriteKeyMatrixToExcel(wbook, R, TargetSheetName, TopLeftCell, OutRangeName);

            ((Excel.Worksheet)wbook.Worksheets[TargetSheetName]).Activate();
        }

        /// <summary>
        /// Add a scalar value to all elements of matrix
        /// </summary>
        public void AddScalarToMatrix_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for AddScalarToMatrix";

                string prompt = "Please select a single-column range with 5 input values for following parameters:\n\n"
                    + "1) Range name input matrix\n"
                    + "2) Scalar value\n"
                    + "3) Range name output matrix\n"
                    + "4) Sheet name\n"
                    + "5) Upper-Left cell address";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 5, out ParameterValues))
                {
                    // assign parameter values
                    string InRangeName = (string)ParameterValues[0];
                    double x = Convert.ToDouble(ParameterValues[1].ToString());
                    string RangeNameOut = (string)ParameterValues[2];
                    string SheetName = (string)ParameterValues[3];
                    string UpperLeftCell = (string)ParameterValues[4];

                    AddScalarToMatrix(InRangeName, x, RangeNameOut, null, SheetName, UpperLeftCell);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("AddScalarToMatrix: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Add a scalar value to all elements of matrix
        /// </summary>
        public void AddScalarToMatrix_macro2()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;

                // form parameters
                string FormTitle = "Parameter Form for AddScalarToMatrix";
                string FuncTitle = "Add a Scalar Value to Matrix";
                string FuncDescr = "Add a number like 5.25 to all elements of matrix";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_MatrixScalarOp(FormTitle, FuncTitle, FuncDescr, "ResultMatrix", "ResultMatrix", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string Matrix1 = ParameterForm1.Matrix1_st;
                double x = double.Parse(ParameterForm1.txtBox1_st);
                string OutMatrix = ParameterForm1.OutMatrix1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                AddScalarToMatrix(Matrix1, x, OutMatrix, null, SheetName, TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("AddScalarToMatrix: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Multiply all elements of matrix with a scalar value
        /// </summary>
        /// <param name="InRangeName">Range name of input matrix</param>
        /// <param name="x">Scalar value</param>
        /// <param name="OutRangeName">Range name of output matrix</param>
        /// <param name="WorkbookFullName">Full workbook path; if null, active workbook is assumed</param>
        /// <param name="TargetSheetName">Name of the sheet in which the matrix is inserted</param>
        /// <param name="TopLeftCell">Upper left cell of matrix in target sheet</param>
        public void MultiplyMatrixWithScalar(string InRangeName, double x, string OutRangeName,
            string WorkbookFullName = null, string TargetSheetName = "ResultMatrix", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Read excel ranges into input matrices
            KeyMatrix M1 = ExcelFunc_NO.ReadKeyMatrixFromExcel(InRangeName, WorkbookFullName);

            // Step 2: Generate resultant (output) matrices
            KeyMatrix R = KeyMatrix.MultiplyMatrixWithScalar(M1, x);

            // Step 3: Write resultant matrices to excel
            ExcelFunc_NO.WriteKeyMatrixToExcel(wbook, R, TargetSheetName, TopLeftCell, OutRangeName);

            ((Excel.Worksheet)wbook.Worksheets[TargetSheetName]).Activate();
        }

        /// <summary>
        /// Multiply all elements of matrix with a scalar value
        /// </summary>
        public void MultiplyMatrixWithScalar_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for MultiplyMatrixWithScalar";

                string prompt = "Please select a single-column range with 5 input values for following parameters:\n\n"
                    + "1) Range name input matrix\n"
                    + "2) Scalar value\n"
                    + "3) Range name output matrix\n"
                    + "4) Sheet name\n"
                    + "5) Upper-Left cell address";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 5, out ParameterValues))
                {
                    // assign parameter values
                    string InRangeName = (string)ParameterValues[0];
                    double x = Convert.ToDouble(ParameterValues[1].ToString());
                    string RangeNameOut = (string)ParameterValues[2];
                    string SheetName = (string)ParameterValues[3];
                    string UpperLeftCell = (string)ParameterValues[4];

                    MultiplyMatrixWithScalar(InRangeName, x, RangeNameOut, null, SheetName, UpperLeftCell);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("MultiplyMatrixWithScalar: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Multiply all elements of matrix with a scalar value
        /// </summary>
        public void MultiplyMatrixWithScalar_macro2()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;

                // form parameters
                string FormTitle = "Parameter Form for MultiplyMatrixWithScalar";
                string FuncTitle = "Multiply Matrix with a Scalar Value";
                string FuncDescr = "Multiply all elements of matrix with a number like 5.25";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_MatrixScalarOp(FormTitle, FuncTitle, FuncDescr, "ResultMatrix", "ResultMatrix", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string Matrix1 = ParameterForm1.Matrix1_st;
                double x = double.Parse(ParameterForm1.txtBox1_st);
                string OutMatrix = ParameterForm1.OutMatrix1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                MultiplyMatrixWithScalar(Matrix1, x, OutMatrix, null, SheetName, TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("MultiplyMatrixWithScalar: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Column-wise sum of matrix elements. Returns a single-column (Nx1) matrix.
        /// </summary>
        /// <param name="InRangeName">Range name of input matrix</param>
        /// <param name="OutRangeName">Range name of output matrix</param>
        /// <param name="WorkbookFullName">Full workbook path; if null, active workbook is assumed</param>
        /// <param name="TargetSheetName">Name of the sheet in which the matrix is inserted</param>
        /// <param name="TopLeftCell">Upper left cell of matrix in target sheet</param>
        public void SumHorizontal(string InRangeName, string OutRangeName,
            string WorkbookFullName = null, string TargetSheetName = "ResultMatrix", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Read excel ranges into input matrices
            KeyMatrix M = ExcelFunc_NO.ReadKeyMatrixFromExcel(InRangeName, WorkbookFullName);

            // Step 2: Generate resultant (output) matrices
            KeyMatrix R = M.SumOfElements(RowColDirection.ColByCol);

            // Step 3: Write resultant matrices to excel
            ExcelFunc_NO.WriteKeyMatrixToExcel(wbook, R, TargetSheetName, TopLeftCell, OutRangeName);

            ((Excel.Worksheet)wbook.Worksheets[TargetSheetName]).Activate();
        }

        /// <summary>
        /// Column-wise sum of matrix elements. Returns a single-column (Nx1) matrix.
        /// </summary>
        public void SumHorizontal_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for SumHorizontal";

                string prompt = "Please select a single-column range with 4 input values for following parameters:\n\n"
                    + "1) Range name input matrix\n"
                    + "2) Range name output matrix\n"
                    + "3) Sheet name\n"
                    + "4) Upper-Left cell address";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 4, out ParameterValues))
                {
                    // assign parameter values
                    string InRangeName = (string)ParameterValues[0];
                    string RangeNameOut = (string)ParameterValues[1];
                    string SheetName = (string)ParameterValues[2];
                    string UpperLeftCell = (string)ParameterValues[3];

                    SumHorizontal(InRangeName, RangeNameOut, null, SheetName, UpperLeftCell);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("SumHorizontal: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Column-wise sum of matrix elements. Returns a single-column (Nx1) matrix.
        /// </summary>
        public void SumHorizontal_macro2()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;

                // form parameters
                string FormTitle = "Parameter Form for SumHorizontal";
                string FuncTitle = "Horizontal Sum of Matrix";
                string FuncDescr = "Column-wise sum of matrix elements. Returns a single-column (Nx1) matrix.";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_SingleMatrixOp(FormTitle, FuncTitle, FuncDescr, "InvMatrix", "InvMatrix", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string Matrix1 = ParameterForm1.Matrix1_st;
                string OutMatrix = ParameterForm1.OutMatrix1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                SumHorizontal(Matrix1, OutMatrix, null, SheetName, TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("SumHorizontal: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Row-wise sum of matrix elements. Returns a single-row (1xN) matrix.
        /// </summary>
        /// <param name="InRangeName">Range name of input matrix</param>
        /// <param name="OutRangeName">Range name of output matrix</param>
        /// <param name="WorkbookFullName">Full workbook path; if null, active workbook is assumed</param>
        /// <param name="TargetSheetName">Name of the sheet in which the matrix is inserted</param>
        /// <param name="TopLeftCell">Upper left cell of matrix in target sheet</param>
        public void SumVertical(string InRangeName, string OutRangeName,
            string WorkbookFullName = null, string TargetSheetName = "ResultMatrix", string TopLeftCell = "A1")
        {
            // Typical flow of calculation:

            // Step 0: Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook;

            if (WorkbookFullName == null || WorkbookFullName == String.Empty)
                wbook = xlApp.ActiveWorkbook;
            else
                wbook = ExcelFunc_NO.GetWorkbookByFullName(xlApp, WorkbookFullName);

            // Step 1: Read excel ranges into input matrices
            KeyMatrix M = ExcelFunc_NO.ReadKeyMatrixFromExcel(InRangeName, WorkbookFullName);

            // Step 2: Generate resultant (output) matrices
            KeyMatrix R = M.SumOfElements(RowColDirection.RowByRow);

            // Step 3: Write resultant matrices to excel
            ExcelFunc_NO.WriteKeyMatrixToExcel(wbook, R, TargetSheetName, TopLeftCell, OutRangeName);

            ((Excel.Worksheet)wbook.Worksheets[TargetSheetName]).Activate();
        }

        /// <summary>
        /// Row-wise sum of matrix elements. Returns a single-row (1xN) matrix.
        /// </summary>
        public void SumVertical_macro()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // inbox parameters
                string title = "Parameter Form for SumVertical";

                string prompt = "Please select a single-column range with 4 input values for following parameters:\n\n"
                    + "1) Range name input matrix\n"
                    + "2) Range name output matrix\n"
                    + "3) Sheet name\n"
                    + "4) Upper-Left cell address";

                List<object> ParameterValues;

                // call input box
                if (ExcelFunc_NO.InputBoxRange(xlapp, prompt, title, 4, out ParameterValues))
                {
                    // assign parameter values
                    string InRangeName = (string)ParameterValues[0];
                    string RangeNameOut = (string)ParameterValues[1];
                    string SheetName = (string)ParameterValues[2];
                    string UpperLeftCell = (string)ParameterValues[3];

                    SumVertical(InRangeName, RangeNameOut, null, SheetName, UpperLeftCell);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("SumVertical: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Row-wise sum of matrix elements. Returns a single-row (1xN) matrix.
        /// </summary>
        public void SumVertical_macro2()
        {
            try
            {
                // get current excel application
                Excel.Application xlapp = ExcelFunc_NO.GetExcelApplicationInstance();

                // get active workbook
                Excel.Workbook wbook = xlapp.ActiveWorkbook;

                // set global form parameters
                ParameterForm1.wbook_st = wbook;

                // form parameters
                string FormTitle = "Parameter Form for SumVertical";
                string FuncTitle = "Vertical Sum of Matrix";
                string FuncDescr = "Row-wise sum of matrix elements. Returns a single-row (1xN) matrix.";

                using (ParameterForm1 myform = new ParameterForm1())
                {
                    myform.PrepForm_SingleMatrixOp(FormTitle, FuncTitle, FuncDescr, "InvMatrix", "InvMatrix", "A1");
                    myform.ShowDialog();
                }

                // return if cancel button is pressed
                if (ParameterForm1.IfCancel) return;

                // assign parameter values
                string Matrix1 = ParameterForm1.Matrix1_st;
                string OutMatrix = ParameterForm1.OutMatrix1_st;
                string SheetName = ParameterForm1.SheetName_st;
                string TopLeftCell = ParameterForm1.TopLeftCell_st;

                SumVertical(Matrix1, OutMatrix, null, SheetName, TopLeftCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show("SumVertical: " + ex.Message + "\n");
            }
        }

    }
}
