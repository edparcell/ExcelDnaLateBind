using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace CobaltQuantware.ExcelDnaLateBind
{
    public class LateBoundObject
    {
        protected Type ObjType { get; private set; }
        protected object Obj { get; private set; }

        public LateBoundObject(object Obj)
        {
            this.ObjType = Obj.GetType();
            this.Obj = Obj;
        }

        public object GetProperty(string name)
        {
            return GetProperty(name, null);
        }

        public object GetProperty(string name, object[] args)
        {
            return ObjType.InvokeMember(
                name,
                BindingFlags.GetProperty | BindingFlags.Public,
                null,
                Obj,
                args
                );
        }

        public void SetProperty(string name, object value)
        {
            SetProperty(name, new object[] { value });
        }

        public void SetProperty(string name, object[] args)
        {
            ObjType.InvokeMember(
                name,
                BindingFlags.SetProperty | BindingFlags.Public,
                null,
                Obj,
                args
                );
        }

        public object InvokeMethod(string name)
        {
            return InvokeMethod(name, null);
        }

        public object InvokeMethod(string name, object[] args)
        {
            return ObjType.InvokeMember(
                name,
                BindingFlags.InvokeMethod | BindingFlags.Public,
                null,
                Obj,
                args
                );
        }
    }

}
