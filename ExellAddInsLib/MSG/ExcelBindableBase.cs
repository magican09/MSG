using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public class ExcelBindableBase :BindableBase, INotifyPropertyChanged, IExcelBindableBase
    {
        
        public Dictionary<string, Tuple<int, int>> CellAddressesMap { get; set; } = new Dictionary<string, Tuple<int, int>>();
    }
}
