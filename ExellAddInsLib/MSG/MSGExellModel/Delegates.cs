using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public delegate void ExcelPropertyChangedEventHandler(object sender, ExcelPropertyChangedEventArgs e);
    
    public class ExcelPropertyChangedEventArgs : EventArgs
    {
        private readonly string propertyName;

   
        public virtual string PropertyName
        {  
            get
            {
                return propertyName;
            }
        }

        public ExcelPropertyChangedEventArgs(string propertyName)
        {
            this.propertyName = propertyName;
        }
    }
}
