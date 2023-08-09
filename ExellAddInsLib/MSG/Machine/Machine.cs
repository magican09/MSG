using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public class Machine : ExcelBindableBase
    {
        public Machine()
        {

        }
        public Machine(string number, string name)
        {
            Number = number;
            Name = name;
        }
    }
}
