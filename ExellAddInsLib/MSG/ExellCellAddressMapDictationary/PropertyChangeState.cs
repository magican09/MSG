using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public  class PropertyChangeState
    {
        public IExcelBindableBase Sender { get; set; }
        public string  PropertyName { get; set; }
        public bool PropertyIsValid { get; set; }
        public PropertyChangeState(IExcelBindableBase sender, string propertyName,bool propertyIsValid = true)
        {
            Sender = sender;
            PropertyName = propertyName;
            PropertyIsValid = propertyIsValid;
        }
    }
}
