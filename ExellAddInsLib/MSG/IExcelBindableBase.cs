﻿using Microsoft.Office.Interop.Excel;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExellAddInsLib.MSG
{
    public interface IExcelBindableBase : ICloneable, INameable, IPropertyChnagedIsSubscribed
    {

        event PropertyChangedEventHandler PropertyChanged;
        void PropertyChange(object sender, string property_name);
        void SetProperty<T>(ref T member, T new_val, [CallerMemberName] string property_name = "");
        ExellCellAddressMapDictationary CellAddressesMap { get; set; }
        Guid Id { get; }
        string Number { get; set; }
        string NumberSuffix { get; }
        IExcelBindableBase Owner { get; set; }

        bool IsValid { get; set; }

        //   ObservableCollection<IExcelBindableBase> Owners { get; set; }
        Excel.Range GetRange(Excel.Worksheet worksheet);
        Excel.Range GetRange(Excel.Worksheet worksheet, int right_border = 100000000, int low_borde = 1000000000, int left_border = 0, int up_border = 0);
        void SetInvalidateCellsColor(XlRgbColor color);
        void ChangeTopRow(int row);
        int GetBottomRow();
        int GetTopRow();
        int GetRowsCount();
        void SetNumberItem(int possition, string number, bool first_itaration = true);
        string GetSelfNamber();



    }
}