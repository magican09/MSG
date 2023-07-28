using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public interface IExcelBindableBase
    {
        event PropertyChangedEventHandler PropertyChanged;
        void SetProperty<T>(ref T member, T new_val, [CallerMemberName] string property_name = "");
        EventedDictationary<string, Tuple<int, int, Excel.Worksheet>> CellAddressesMap { get; set; }
        //  Excel.Worksheet RegisterSheet { get; set; }
    }
}