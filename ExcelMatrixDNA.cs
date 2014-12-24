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
    public class ExcelMatrixDNA
    {
        /// <summary>
        /// Sum of all matrix elements
        /// </summary>
        /// <param name="RangeName">Range name of matrix in Excel</param>
        /// <returns>Sum of all elements</returns>
        [ExcelFunction(Description = "Sum of all matrix elements", Category = "Matrix Functions")]
        public static double MatrixSum(string RangeName)
        {
            try
            {
                // Get current application and active workbook
                Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
                Excel.Workbook wbook = xlApp.ActiveWorkbook;

                // Read excel range into input matrix
                KeyMatrix M = ExcelFunc_NO.ReadKeyMatrixFromExcel(RangeName);

                // return matrix sum
                return M.Sum();
            }
            catch (Exception ex)
            {
                throw new Exception("Sum: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Determinant of a NxN square matrix
        /// </summary>
        /// <param name="RangeName">Range name of matrix in Excel</param>
        /// <returns>Sum of all elements</returns>
        [ExcelFunction(Description = "Determinant of a NxN square matrix", Category = "Matrix Functions")]
        public static double Determinant(string RangeName)
        {
            try
            {
                // Get current application and active workbook
                Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
                Excel.Workbook wbook = xlApp.ActiveWorkbook;

                // Read excel range into input matrix
                KeyMatrix M = ExcelFunc_NO.ReadKeyMatrixFromExcel(RangeName);

                // return matrix sum
                return M.Determinant();
            }
            catch (Exception ex)
            {
                throw new Exception("Determinant: " + ex.Message + "\n");
            }
        }

        /// <summary>
        /// Multiply Matrices (matrix multiplication in linear algebra)
        /// </summary>
        /// <param name="RangeName1">Range name of 1. input matrix</param>
        /// <param name="RangeName2">Range name of 2. input matrix</param>
        /// <returns>Resultant matrix</returns>
        [ExcelFunction(Description = "Matrix Multiplication, returns a 2-dim array (array formula)", Category = "Matrix Functions")]
        public static double[,] MatrixMultiplication(string RangeName1, string RangeName2)
        {
            // Get current application and active workbook
            Excel.Application xlApp = ExcelFunc_NO.GetExcelApplicationInstance();
            Excel.Workbook wbook = xlApp.ActiveWorkbook;

            // Read excel range into input matrix
            KeyMatrix M1 = ExcelFunc_NO.ReadKeyMatrixFromExcel(RangeName1);
            KeyMatrix M2 = ExcelFunc_NO.ReadKeyMatrixFromExcel(RangeName2);

            // multiply matrices
            return KeyMatrix.MultiplyMatrices(M1, M2).toArray;
        }

        /// <summary>
        /// Create a NxM matrix with random valued elements
        /// </summary>
        [ExcelCommand(MenuName = "Matrix Functions", MenuText = "Create Random Matrix (InputBox)")]
        public static void CreateRandomMatrix_macro()
        {
            var xm = new ExcelMatrix();
            xm.CreateRandomMatrix_macro();
        }

        /// <summary>
        /// Create a NxM matrix with random valued elements
        /// </summary>
        [ExcelCommand(MenuName = "Matrix Functions", MenuText = "Create Random Matrix")]
        public static void CreateRandomMatrix_macro2()
        {
            var xm = new ExcelMatrix();
            xm.CreateRandomMatrix_macro2();
        }

        /// <summary>
        /// Multiply Matrices (matrix multiplication in linear algebra)
        /// </summary>
        //[ExcelCommand(MenuName = "Matrix Functions", MenuText = "Multiply Matrices (InputBox)")]
        public static void MultiplyMatrices_macro()
        {
            var xm = new ExcelMatrix();
            xm.MultiplyMatrices_macro();
        }

        /// <summary>
        /// Multiply Matrices (matrix multiplication in linear algebra)
        /// </summary>
        [ExcelCommand(MenuName = "Matrix Functions", MenuText = "Multiply Matrices")]
        public static void MultiplyMatrices_macro2() 
        {
            var xm = new ExcelMatrix();
            xm.MultiplyMatrices_macro2();
        }

        /// <summary>
        /// Add Matrices; element-wise addition of two equal-sized matrices 
        /// </summary>
        //[ExcelCommand(MenuName = "Matrix Functions", MenuText = "Add Matrices (InputBox)")]
        public static void AddMatrices_macro()
        {
            var xm = new ExcelMatrix();
            xm.AddMatrices_macro();
        }

        /// <summary>
        /// Add Matrices; element-wise addition of two equal-sized matrices 
        /// </summary>
        [ExcelCommand(MenuName = "Matrix Functions", MenuText = "Add Matrices")]
        public static void AddMatrices_macro2()
        {
            var xm = new ExcelMatrix();
            xm.AddMatrices_macro2();
        }

        /// <summary>
        /// Inverse Matrix
        /// </summary>
        //[ExcelCommand(MenuName = "Matrix Functions", MenuText = "Inverse Matrix (InputBox)")]
        public static void InverseMatrix_macro()
        {
            var xm = new ExcelMatrix();
            xm.InverseMatrix_macro();
        }

        /// <summary>
        /// Inverse Matrix
        /// </summary>
        [ExcelCommand(MenuName = "Matrix Functions", MenuText = "Inverse Matrix")]
        public static void InverseMatrix_macro2()
        {
            var xm = new ExcelMatrix();
            xm.InverseMatrix_macro2();
        }

        /// <summary>
        /// Transpose Matrix
        /// </summary>
        //[ExcelCommand(MenuName = "Matrix Functions", MenuText = "Transpose Matrix (InputBox)")]
        public static void TransposeMatrix_macro()
        {
            var xm = new ExcelMatrix();
            xm.TransposeMatrix_macro();
        }

        /// <summary>
        /// Transpose Matrix
        /// </summary>
        [ExcelCommand(MenuName = "Matrix Functions", MenuText = "Transpose Matrix")]
        public static void TransposeMatrix_macro2()
        {
            var xm = new ExcelMatrix();
            xm.TransposeMatrix_macro2();
        }

        /// <summary>
        /// Add a scalar value to all elements of matrix
        /// </summary>
        //[ExcelCommand(MenuName = "Matrix Functions", MenuText = "Add Scalar Value (InputBox)")]
        public static void AddScalarToMatrix_macro()
        {
            var xm = new ExcelMatrix();
            xm.AddScalarToMatrix_macro();
        }

        /// <summary>
        /// Add a scalar value to all elements of matrix
        /// </summary>
        [ExcelCommand(MenuName = "Matrix Functions", MenuText = "Add Scalar Value")]
        public static void AddScalarToMatrix_macro2()
        {
            var xm = new ExcelMatrix();
            xm.AddScalarToMatrix_macro2();
        }

        /// <summary>
        /// Multiply all elements of matrix with a scalar value
        /// </summary>
        //[ExcelCommand(MenuName = "Matrix Functions", MenuText = "Multiply with Scalar Value (InputBox)")]
        public static void MultiplyMatrixWithScalar_macro()
        {
            var xm = new ExcelMatrix();
            xm.MultiplyMatrixWithScalar_macro();
        }

        /// <summary>
        /// Multiply all elements of matrix with a scalar value
        /// </summary>
        [ExcelCommand(MenuName = "Matrix Functions", MenuText = "Multiply with Scalar Value")]
        public static void MultiplyMatrixWithScalar_macro2()
        {
            var xm = new ExcelMatrix();
            xm.MultiplyMatrixWithScalar_macro2();
        }

        /// <summary>
        /// Column-wise sum of matrix elements. Returns a single-column (Nx1) matrix.
        /// </summary>
        //[ExcelCommand(MenuName = "Matrix Functions", MenuText = "Horizontal Matrix Sum (InputBox)")]
        public static void SumHorizontal_macro()
        {
            var xm = new ExcelMatrix();
            xm.SumHorizontal_macro();
        }

        /// <summary>
        /// Column-wise sum of matrix elements. Returns a single-column (Nx1) matrix.
        /// </summary>
        [ExcelCommand(MenuName = "Matrix Functions", MenuText = "Horizontal Matrix Sum")]
        public static void SumHorizontal_macro2()
        {
            var xm = new ExcelMatrix();
            xm.SumHorizontal_macro2();
        }

        /// <summary>
        /// Row-wise sum of matrix elements. Returns a single-row (1xN) matrix.
        /// </summary>
        //[ExcelCommand(MenuName = "Matrix Functions", MenuText = "Vertical Matrix Sum (InputBox)")]
        public static void SumVertical_macro()
        {
            var xm = new ExcelMatrix();
            xm.SumVertical_macro();
        }

        /// <summary>
        /// Row-wise sum of matrix elements. Returns a single-row (1xN) matrix.
        /// </summary>
        [ExcelCommand(MenuName = "Matrix Functions", MenuText = "Vertical Matrix Sum")]
        public static void SumVertical_macro2()
        {
            var xm = new ExcelMatrix();
            xm.SumVertical_macro2();
        }

    }
}
