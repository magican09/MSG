using Microsoft.Office.Interop.Excel;
using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class ExcelPropAddress
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
        private string _cellNumberFormat;
        public string CellNumberFormat
        {
            get { return _cellNumberFormat; }
            set
            {
                _cellNumberFormat = value;
                this.Cell.NumberFormat = _cellNumberFormat;
            }
        }
        private Type _valueType;
        public Type ValueType
        {
            get { return _valueType; }
            set
            {
                _valueType = value;
                this.SetCellNumberFormat();
            }
        }

        private int _rowHashValue;
        public int RowHashValue
        {
            get { return _rowHashValue; }
            set
            {
                _rowHashValue = value;
            }
        }

        private int _columnHashValue;
        public int ColumnHashValue
        {
            get { return _columnHashValue; }
            set
            {
                _columnHashValue = value;
            }
        }
        public Func<object, bool> ValidateValueCallBack { get; set; }
        public Func<object, object> CoerceValueCallback { get; set; }



        public Excel.Range Cell
        {
            get { return this.Worksheet.Cells[Row, Column]; }
        }
        public ExcelPropAddress()
        {

        }
        public void SetCellNumberFormat()
        {
            if (_valueType == null) return;
            if ( _valueType == typeof(double) || _valueType == typeof(decimal))
            {
                Char separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0];
                CellNumberFormat = $"0.00";
            }
            if (_valueType == typeof(int) )
            {
                Char separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0];
                CellNumberFormat = $"0";
            }
            if (_valueType == typeof(DateTime))
            {
                Char separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0];
                CellNumberFormat = $"dd.mm.yyyy";
            }
            if (_valueType == typeof(string))
            {
                CellNumberFormat = $"@";
            }
        }

        public ExcelPropAddress(int row, int column, Excel.Worksheet worksheet, Type val_type,
            string prop_name = "",
            Func<object, bool> validate_value_call_back = null,
               Func<object, object> coerce_value_call_back = null)
        {
            Row = row;
            Column = column;
            Worksheet = worksheet;
            ProprertyName = prop_name;
            this.Cell.Interior.Color = XlRgbColor.rgbGreenYellow;
            ValidateValueCallBack = validate_value_call_back;
            CoerceValueCallback = coerce_value_call_back;
            ValueType = val_type;
            this.SetCellNumberFormat();

        }
        public ExcelPropAddress(ExcelPropAddress ex_addr)
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
