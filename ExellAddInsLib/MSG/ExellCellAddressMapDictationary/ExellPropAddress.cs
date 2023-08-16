using Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;
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
        public ExellPropAddress(int row, int column, Excel.Worksheet worksheet, Type val_type, string prop_name = "")
        {
            Row = row;
            Column = column;
            Worksheet = worksheet;
            ProprertyName = prop_name;
            this.Cell.Interior.Color = XlRgbColor.rgbGreenYellow;

         
            if (val_type == typeof(int) || val_type == typeof(double) || val_type == typeof(decimal))
            {
                Char separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0];
               this.Cell.NumberFormat = $"0.00";
            }
            if (val_type == typeof(DateTime))
            {
                Char separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0];
                this.Cell.NumberFormat = $"dd.mm.yyyy";
            }
            if (val_type == typeof(string))
            {
                this.Cell.NumberFormat = $"@";
            }
           


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
