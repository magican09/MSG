using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExellAddInsLib.MSG
{
    public interface IExcelBindableBase : ICloneable
    {
        event PropertyChangedEventHandler PropertyChanged;
        void PropertyChange(object sender, string property_name);
        void SetProperty<T>(ref T member, T new_val, [CallerMemberName] string property_name = "");
        ExellCellAddressMapDictationary CellAddressesMap { get; set; }
        Guid Id { get; }
        ObservableCollection<IExcelBindableBase> Owners { get; set; }
    }
}