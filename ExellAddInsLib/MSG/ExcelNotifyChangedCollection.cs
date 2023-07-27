using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class ExcelNotifyChangedCollection<T> : ObservableCollection<T>, IExcelBindableBase
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void SetProperty<T>(ref T member, T new_val, [CallerMemberName] string property_name = "")
        {

            member = new_val;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property_name));

        }
        public Dictionary<string, Tuple<int, int, Excel.Worksheet>> CellAddressesMap { get; set; } = new Dictionary<string, Tuple<int, int, Excel.Worksheet>>();

    }
}
