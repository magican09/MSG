using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace ExellAddInsLib.MSG
{
    public class ExcelBindableBase : BindableBase, INotifyPropertyChanged, IExcelBindableBase
    {

        public Dictionary<string, Tuple<int, int>> CellAddressesMap { get; set; } = new Dictionary<string, Tuple<int, int>>();
    }
}
