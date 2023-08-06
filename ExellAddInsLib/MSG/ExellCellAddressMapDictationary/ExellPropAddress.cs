using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.Data.Common;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class ExellPropAddress
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public Excel.Worksheet Worksheet { get; set; }
        public string ProprertyName { get; set; }
        private bool _isValid = true;
        public bool IsValid
        {
            get { return _isValid; }
            set { _isValid = value; }
        }

        private IExcelBindableBase _owner;

        public IExcelBindableBase Owner
        {
            get { return _owner; }
            set { _owner = value; }
        }

        //private object _cellColor;

        //public object CellColor
        //{
        //    get { return _cellColor; }
        //    set {
        //        _cellColor = value;
        //        if (_cellColor is XlRgbColor rgb_color)
        //            this.Cell.Interior.Color = rgb_color;
        //        else if (_cellColor is XlColorIndex ind_color)
        //            this.Cell.Interior.ColorIndex = ind_color;
        //    }
        //}

        public Excel.Range Cell
        {
            get { return this.Worksheet.Cells[Row, Column]; }
        }
        public ExellPropAddress()
        {

        }
        public ExellPropAddress(int row, int column, Excel.Worksheet worksheet, string prop_name = "")//,)
        {
            Row = row;
            Column = column;
            Worksheet = worksheet;
            ProprertyName = prop_name;
            this.Cell.Interior.Color = XlRgbColor.rgbGreenYellow;

        }
        public ExellPropAddress(ExellPropAddress ex_addr)
        {
            Row = ex_addr.Row;
            Column = ex_addr.Column;
            Worksheet = ex_addr.Worksheet;
            ProprertyName = ex_addr.ProprertyName;
           
        }

        //public void SetColor(XlRgbColor color)
        //{
        //    this.Cell.Interior.Color = color;
        //}
        //public void SetColor(XlRgbColor color)
        //{
        //    this.Cell.Interior.Color = color;
        //}
    }
}
