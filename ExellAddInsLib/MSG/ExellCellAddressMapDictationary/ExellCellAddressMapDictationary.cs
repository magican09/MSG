using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public delegate void SetWorksheet_Heandler(Excel.Worksheet worksheet);
    public class ExellCellAddressMapDictationary : EventedDictationary<string, ExellPropAddress>
    {
        public event SetWorksheet_Heandler OnSetWorksheet;

        public void SetWorksheet(Excel.Worksheet worksheet)
        {

            foreach (KeyValuePair<string, ExellPropAddress> kvp in this)
            {
                kvp.Value.Worksheet = worksheet;
            }
            OnSetWorksheet?.Invoke(worksheet);
        }

        public ExellCellAddressMapDictationary() : base()
        {

        }
        public void SetCellNumberFormat()
        {

            foreach (var kvp in this)
                kvp.Value.SetCellNumberFormat();
        }

    }
}
