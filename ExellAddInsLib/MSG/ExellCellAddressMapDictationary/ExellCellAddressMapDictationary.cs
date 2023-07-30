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

        public ExellCellAddressMapDictationary()
        {

        }
        private IExcelBindableBase  _owner;

        public IExcelBindableBase  Owner
        {
            get { return _owner; }
            set { _owner = value; }
        }

        //public void UpdateWorksheets()
        //{
        //    foreach(KeyValuePair<string, ExellPropAddress> kvp in this)


        //}
    }
}
