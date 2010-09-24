using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace CobaltQuantware.ExcelDnaLateBind
{
    public class LateBoundObject
    {
        private const int RETRY_COUNT = 10;
        private const int RETRY_DELAY = 25;
        private const string VBA_E_IGNORE = "0x800AC472";

        public Type ObjType { get; private set; }
        public object Obj { get; private set; }

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
            TargetInvocationException lastException = null;
            for (int retry = 0; retry < RETRY_COUNT; ++retry)
            {
                try
                {
                    return ObjType.InvokeMember(
                        name,
                        BindingFlags.GetProperty | BindingFlags.Public,
                        null,
                        Obj,
                        args
                        );
                }
                catch (TargetInvocationException e)
                {
                    lastException = e;
                    if (e.Message.Contains(VBA_E_IGNORE))
                        System.Threading.Thread.Sleep(RETRY_DELAY);
                    else
                        throw;
                }                
            }
            if (lastException != null) throw lastException;
            return null;
        }

        public void SetProperty(string name, object value)
        {
            SetProperty(name, new object[] { value });
        }

        public void SetProperty(string name, object[] args)
        {
            TargetInvocationException lastException = null;
            for (int retry = 0; retry < RETRY_COUNT; ++retry)
            {
                try
                {
                    ObjType.InvokeMember(
                        name,
                        BindingFlags.SetProperty | BindingFlags.Public,
                        null,
                        Obj,
                        args
                        );
                }
                catch (TargetInvocationException e)
                {
                    lastException = e;
                    if (e.Message.Contains(VBA_E_IGNORE))
                        System.Threading.Thread.Sleep(RETRY_DELAY);
                    else
                        throw;

                }
            }
            if (lastException != null) throw lastException;
        }

        public object InvokeMethod(string name)
        {
            return InvokeMethod(name, null);
        }

        public object InvokeMethod(string name, object[] args)
        {
            TargetInvocationException lastException = null;
            for (int retry = 0; retry < RETRY_COUNT; ++retry)
            {
                try
                {
                    return ObjType.InvokeMember(
                        name,
                        BindingFlags.InvokeMethod | BindingFlags.Public,
                        null,
                        Obj,
                        args
                        );
                }
                catch (TargetInvocationException e)
                {
                    lastException = e;
                    if (e.Message.Contains(VBA_E_IGNORE))
                        System.Threading.Thread.Sleep(RETRY_DELAY);
                    else
                        throw;

                }
            }
            if (lastException != null) throw lastException;
            return null;
        }
    }

}
