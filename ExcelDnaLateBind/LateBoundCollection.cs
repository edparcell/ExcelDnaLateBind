using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CobaltQuantware.ExcelDnaLateBind
{
    public class LateBoundCollection<T> : LateBoundObject
    {
        public LateBoundCollection(object Collection) : base(Collection) { }

        public int Count { get { return (int)GetProperty("Count"); } }

        public T this[string name] 
        { 
            get
            {
                return LateBoundObjectFactory.Create<T>(GetProperty("Item", new object[] { name }));
            }
        }

    }
}
