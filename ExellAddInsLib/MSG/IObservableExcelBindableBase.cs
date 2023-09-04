using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExellAddInsLib.MSG
{
    public  interface IObservableExcelBindableBase: IExcelBindableBase,IObservable<PropertyChangeState>
    {
        //ExcelPropAddress this[string i] { get; }
         void SetPropertyValidStatus(string prop_name, bool isValid);
         Excel.Range GetCell(string prop_name);
        ExcelPropAddress GetPropAddress(string prop_name);
        List<IDisposable> Subscribers { get;  }
    }
}
