using System;
using System.Collections.Generic;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class ExcelBindableBase : BindableBase, INotifyPropertyChanged, IExcelBindableBase
    {

        public Dictionary<string, Tuple<int, int, Excel.Worksheet>> CellAddressesMap { get; set; } = new Dictionary<string, Tuple<int, int, Excel.Worksheet>>();
    }
}
