using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Microsoft.CSharp;
using System.CodeDom.Compiler;
using System.Reflection;

namespace FinaquantInExcel
{
    /// <summary>
    /// Class with various helper methods
    /// </summary>
    public class HelperFunc
    {
        private static MethodInfo _methodinfo = null;
        private static object _assemblyinstance = null;

        /// <summary>
        /// Dynamic compiler and executer for applying a user-defined function on each row of a table.
        /// </summary>
        /// <param name="UserCode">C# code entered by user</param>
        /// <remarks>
        /// See: How to compile and execute a user defined formula dynamically (runtime) in c#?
        /// http://stackoverflow.com/questions/26505145/how-to-compile-and-execute-a-user-defined-formula-dynamically-runtime-in-c/26509760#26509760
        /// </remarks>
        internal static void UserDefinedTransformFunction(string UserCode)
        {
            CSharpCodeProvider c = new CSharpCodeProvider();
            ICodeCompiler icc = c.CreateCompiler();
            CompilerParameters cp = new CompilerParameters();

            cp.ReferencedAssemblies.Add("system.dll");
            cp.CompilerOptions = "/t:library";
            cp.GenerateInMemory = true;

            StringBuilder sb = new StringBuilder("");
            sb.Append("using System; \n");
            sb.Append("using System.Collections.Generic; \n");

            sb.Append("namespace CSCodeEvaler{ \n");
            sb.Append("public class CSCodeEvaler{ \n");

            // start function envelope
            sb.Append("public void UserFunc(ref Dictionary<string, string> TA, ref Dictionary<string, int> NA, ref Dictionary<string, double> KF){ \n");

            // enveloped user code
            sb.Append(UserCode + "\n");

            // close function envelope
            sb.Append("} \n");

            sb.Append("} \n");
            sb.Append("}\n");

            CompilerResults cr = icc.CompileAssemblyFromSource(cp, sb.ToString());
            if (cr.Errors.Count > 0)
            {
                throw new Exception("ERROR in user code at line " + (cr.Errors[0].Line - 5) + ": " + cr.Errors[0].ErrorText + "\n");
            }

            System.Reflection.Assembly a = cr.CompiledAssembly;
            object o = a.CreateInstance("CSCodeEvaler.CSCodeEvaler");

            Type t = o.GetType();
            MethodInfo mi = t.GetMethod("UserFunc");
            // mi.Invoke(o, Parameters);

            // assign static parameters of class
            HelperFunc._methodinfo = mi;
            HelperFunc._assemblyinstance = o;
        }

        // check validity of user code
        internal static bool CheckUserDefinedTransformFunction(string UserCode, out string ErrorStr)
        {
            CSharpCodeProvider c = new CSharpCodeProvider();
            ICodeCompiler icc = c.CreateCompiler();
            CompilerParameters cp = new CompilerParameters();

            cp.ReferencedAssemblies.Add("system.dll");
            cp.CompilerOptions = "/t:library";
            cp.GenerateInMemory = true;

            StringBuilder sb = new StringBuilder("");
            sb.Append("using System; \n");
            sb.Append("using System.Collections.Generic; \n");

            sb.Append("namespace CSCodeEvaler{ \n");
            sb.Append("public class CSCodeEvaler{ \n");

            // start function envelope
            sb.Append("public void UserFunc(ref Dictionary<string, string> TA, ref Dictionary<string, int> NA, ref Dictionary<string, double> KF){ \n");

            // enveloped user code
            sb.Append(UserCode + "\n");

            // close function envelope
            sb.Append("} \n");

            sb.Append("} \n");
            sb.Append("}\n");

            CompilerResults cr = icc.CompileAssemblyFromSource(cp, sb.ToString());
            if (cr.Errors.Count > 0)
            {
                ErrorStr = "ERROR in user code at line " + (cr.Errors[0].Line - 5) + ": " + cr.Errors[0].ErrorText + "\n";
                return false;
            }
            else
            {
                ErrorStr = "OK";
                return true;
            }
        }

        /// <summary>
        /// User-defined function for transforming rows (row-by-row processing)
        /// </summary>
        /// <param name="TA">Dictionary (associative array) for text attributes</param>
        /// <param name="NA">Dictionary (associative array) for numeric attributes</param>
        /// <param name="KF">Dictionary (associative array) for key figures</param>
        public static void ApplyUserDefinedTransformFuncOnTableRow(
            ref Dictionary<string, string> TA,
            ref Dictionary<string, int> NA,
            ref Dictionary<string, double> KF)
        {
            object[] Parameters = new object[] { TA, NA, KF };
            HelperFunc._methodinfo.Invoke(HelperFunc._assemblyinstance, Parameters);
        }

        /// <summary>
        /// Dynamic compiler and executer for applying a user-defined filter function on each row of a table.
        /// </summary>
        /// <param name="UserCode">C# code entered by user</param>
        internal static void UserDefinedFilterFunction(string UserCode)  
        {
            CSharpCodeProvider c = new CSharpCodeProvider();
            ICodeCompiler icc = c.CreateCompiler();
            CompilerParameters cp = new CompilerParameters();

            cp.ReferencedAssemblies.Add("system.dll");
            cp.CompilerOptions = "/t:library";
            cp.GenerateInMemory = true;

            StringBuilder sb = new StringBuilder("");
            sb.Append("using System; \n");
            sb.Append("using System.Collections.Generic; \n");

            sb.Append("namespace CSCodeEvaler{ \n");
            sb.Append("public class CSCodeEvaler{ \n");

            // start function envelope
            sb.Append("public bool UserFunc(ref Dictionary<string, string> TA, ref Dictionary<string, int> NA, ref Dictionary<string, double> KF){ \n");

            // enveloped user code
            sb.Append(UserCode + "\n");

            // close function envelope
            sb.Append("} \n");

            sb.Append("} \n");
            sb.Append("}\n");

            CompilerResults cr = icc.CompileAssemblyFromSource(cp, sb.ToString());
            if (cr.Errors.Count > 0)
            {
                throw new Exception("ERROR in user code at line " + (cr.Errors[0].Line - 5) + ": " + cr.Errors[0].ErrorText + "\n");
            }

            System.Reflection.Assembly a = cr.CompiledAssembly;
            object o = a.CreateInstance("CSCodeEvaler.CSCodeEvaler");

            Type t = o.GetType();
            MethodInfo mi = t.GetMethod("UserFunc");
            // mi.Invoke(o, Parameters);

            // assign static parameters of class
            HelperFunc._methodinfo = mi;
            HelperFunc._assemblyinstance = o;
        }

        // check validity of user code
        internal static bool CheckUserDefinedFilterFunction(string UserCode, out string ErrorStr) 
        {
            CSharpCodeProvider c = new CSharpCodeProvider();
            ICodeCompiler icc = c.CreateCompiler();
            CompilerParameters cp = new CompilerParameters();

            cp.ReferencedAssemblies.Add("system.dll");
            cp.CompilerOptions = "/t:library";
            cp.GenerateInMemory = true;

            StringBuilder sb = new StringBuilder("");
            sb.Append("using System; \n");
            sb.Append("using System.Collections.Generic; \n");

            sb.Append("namespace CSCodeEvaler{ \n");
            sb.Append("public class CSCodeEvaler{ \n");

            // start function envelope
            sb.Append("public bool UserFunc(ref Dictionary<string, string> TA, ref Dictionary<string, int> NA, ref Dictionary<string, double> KF){ \n");

            // enveloped user code
            sb.Append(UserCode + "\n");

            // close function envelope
            sb.Append("} \n");

            sb.Append("} \n");
            sb.Append("}\n");

            CompilerResults cr = icc.CompileAssemblyFromSource(cp, sb.ToString());
            if (cr.Errors.Count > 0)
            {
                ErrorStr = "ERROR in user code at line " + (cr.Errors[0].Line - 5) + ": " + cr.Errors[0].ErrorText + "\n";
                return false;
            }
            else
            {
                ErrorStr = "OK";
                return true;
            }
        }

        /// <summary>
        /// User-defined function for filtering rows (row-by-row processing)
        /// </summary>
        /// <param name="TA">Dictionary (associative array) for text attributes</param>
        /// <param name="NA">Dictionary (associative array) for numeric attributes</param>
        /// <param name="KF">Dictionary (associative array) for key figures</param>
        public static bool ApplyUserDefinedFilterFuncOnTableRow( 
            ref Dictionary<string, string> TA,
            ref Dictionary<string, int> NA,
            ref Dictionary<string, double> KF)
        {
            object[] Parameters = new object[] { TA, NA, KF };
            return (bool) HelperFunc._methodinfo.Invoke(HelperFunc._assemblyinstance, Parameters);
        }

    }
}
