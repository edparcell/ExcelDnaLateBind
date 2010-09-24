using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;

namespace CobaltQuantware.ExcelDnaLateBind
{
    public class Utils
    {
        public static Assembly Excel12Assembly()
        {
            return Assembly.LoadFrom(@"C:\Program Files (x86)\Microsoft Visual Studio 10.0\Visual Studio Tools for Office\PIA\Office12\Microsoft.Office.Interop.Excel.dll");
        }

        public static Type GetTypeForComObject(object obj, Assembly assembly)
        {
            IntPtr iunkwn = Marshal.GetIUnknownForObject(obj);
            Type[] assemblyTypes = assembly.GetTypes();
            foreach (Type type in assemblyTypes)
            {
                Guid iid = type.GUID;
                if (!type.IsInterface || iid == Guid.Empty) continue;
                
                IntPtr iptr = IntPtr.Zero;
                Marshal.QueryInterface(iunkwn, ref iid, out iptr);

                if (iptr != IntPtr.Zero) return type;
            }
            return null;
        }
    }
}
