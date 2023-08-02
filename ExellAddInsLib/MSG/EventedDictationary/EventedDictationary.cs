using System;
using System.Collections.Generic;

namespace ExellAddInsLib.MSG
{
    public class EventedDictationary<TKey, TValue> : Dictionary<TKey, TValue>,ICloneable
    {
        public class AddEventArgs : EventArgs
        {
            private TKey _key;
            private TValue _value;

            public AddEventArgs(TKey key, TValue value)
            {
                _key = key;
                _value = value;
            }
            public TKey Key
            {
                get
                {
                    return _key;
                }
                set { _key = value; }
            }

            public TValue Value
            {
                get
                {
                    return _value;
                }
                set { _value = value; }
            }
        }
        public delegate void AddEventHandler(IExcelBindableBase sender, AddEventArgs pAddEventArgs);
        public event AddEventHandler AddEvent;
        private IExcelBindableBase _owner;

        [DontClone]
        public IExcelBindableBase Owner
        {
            get { return _owner; }
            set { _owner = value; }
        }
        new public void Add(TKey pKey, TValue pValue)
        {
            if (this.ContainsKey(pKey)) return;

            if (AddEvent != null)
                AddEvent(Owner, new AddEventArgs(pKey, pValue));

            base.Add(pKey, pValue);
        }

        public object Clone()
        {
            return Activator.CreateInstance(this.GetType());
        }


        //public void Add(IExcelBindableBase sender, TKey pKey, TValue pValue)
        //{
        //    if (this.ContainsKey(pKey)) return;

        //    if (AddEvent != null)
        //        AddEvent(sender, new AddEventArgs(pKey, pValue));

        //    base.Add(pKey, pValue);
        //}

    }
}
