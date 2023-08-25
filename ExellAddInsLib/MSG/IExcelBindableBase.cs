using Microsoft.Office.Interop.Excel;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExellAddInsLib.MSG
{
    public interface IExcelBindableBase : ICloneable, INameable, IPropertyChnagedIsSubscribed
    {
        Excel.Worksheet Worksheet { get; set; }
        event PropertyChangedEventHandler PropertyChanged;
          event BeforePropertyChangeEventHandler BeforePropertyChange;
        void PropertyChange(object sender, string property_name);
   //     void SetProperty<T>(ref T member, T new_val, [CallerMemberName] string property_name = "");
        ExellCellAddressMapDictationary CellAddressesMap { get; set; }
        Guid Id { get; }
        string Number { get; set; }
        string NumberPrefix { get; }
        string NumberSuffix { get; }
        IExcelBindableBase Owner { get; set; }

        bool IsValid { get; set; }
        bool IsChanged { get; set; }
        //   ObservableCollection<IExcelBindableBase> Owners { get; set; }
        Excel.Range GetRange();
        Excel.Range GetRange(int right_border = 100000000, int low_borde = 1000000000, int left_border = 0, int up_border = 0);
        void SetInvalidateCellsColor(XlRgbColor color);
        void ChangeTopRow(int row);
        int GetBottomRow();
        int GetTopRow();
        int GetRowsCount();
        void SetNumberItem(int possition, string number, bool first_itaration = true);
        string GetSelfNamber();

        void UpdateExcelRepresetation();
        int AdjustExcelRepresentionTree(int top_row);
        void SetStyleFormats(int col);
        //void SetBordersLine(Excel.Range range);
        //void SetBordersLine(Excel.Range range, bool right = true, bool left = true, bool top = true, bool bottom = true);
        //void SetBordersLine(Excel.Range range,
        // Excel.XlLineStyle right = Excel.XlLineStyle.xlDouble,
        // Excel.XlLineStyle left = Excel.XlLineStyle.xlDouble,
        // Excel.XlLineStyle top = Excel.XlLineStyle.xlDouble,
        // Excel.XlLineStyle bottom = Excel.XlLineStyle.xlDouble);

    }
}