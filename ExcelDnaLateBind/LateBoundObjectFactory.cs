using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace CobaltQuantware.ExcelDnaLateBind
{
    internal class LateBoundObjectFactory
    {
        public static T Create<T>(object obj)
        {
            Type type = typeof(T);
            ConstructorInfo constructor = type.GetConstructor(new System.Type[] { typeof(object) });
            T newObj = (T) constructor.Invoke(new object[] {obj});
            return newObj;
        }
    }
}
