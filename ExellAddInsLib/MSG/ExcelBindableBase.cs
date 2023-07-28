using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class ExcelBindableBase : BindableBase, INotifyPropertyChanged, IExcelBindableBase
    {

        public EventedDictationary<string, Tuple<int, int, Excel.Worksheet>> CellAddressesMap { get; set; } = new EventedDictationary<string, Tuple<int, int, Excel.Worksheet>>();
        public ExcelBindableBase()
        {
            CellAddressesMap.AddEvent += OnCellAdressAdd;
        }

        private void OnCellAdressAdd(EventedDictationary<string, Tuple<int, int, Worksheet>>.AddEventArgs pAddEventArgs)
        {
           if(pAddEventArgs!=null)
            {
                Excel.Worksheet worksheet = pAddEventArgs.Value.Item3;
                worksheet.Cells[pAddEventArgs.Value.Item1, pAddEventArgs.Value.Item2].Interior.Color
                                = XlRgbColor.rgbAliceBlue;
            }
          
        }
    }
}
