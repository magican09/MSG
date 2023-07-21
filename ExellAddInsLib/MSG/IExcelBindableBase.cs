using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExellAddInsLib.MSG
{
    public interface IExcelBindableBase
    {
        event PropertyChangedEventHandler PropertyChanged;
        void SetProperty<T>(ref T member, T new_val, [CallerMemberName] string property_name = "");
        Dictionary<string, Tuple<int, int>> CellAddressesMap { get; set; }
    }
}