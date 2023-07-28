using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class EventedDictationary<TKey, TValue> : Dictionary<TKey, TValue>
    {
        public class AddEventArgs : EventArgs
        {
            private TKey _key;
            private TValue _value;

            public AddEventArgs(TKey key, TValue value)
            {
                _key= key;
                _value= value;
            }
            public TKey Key
            {
                get
                {
                    return _key;
                }
            }

            public TValue Value
            {
                get
                {
                    return _value;
                }
            }
        }
        public delegate void AddEventHandler(AddEventArgs pAddEventArgs);
        public event AddEventHandler AddEvent;
       
        public void Add(TKey pKey, TValue pValue)
        {
            if (this.ContainsKey(pKey)) return;
            if (AddEvent != null)
                AddEvent(new AddEventArgs(pKey, pValue));
            base.Add(pKey, pValue);
        }
    }
}
