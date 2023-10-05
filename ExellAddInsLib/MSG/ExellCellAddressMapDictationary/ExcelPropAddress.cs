using Microsoft.Office.Interop.Excel;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class ExcelPropAddress : IObserver<PropertyChangeState>
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public Excel.Worksheet Worksheet { get; set; }
        public string ProprertyName { get; set; }
        private bool _isValid = true;
        public bool IsValid
        {
            get { return _isValid; }
            set { 
                
                _isValid = value;
                if (_isValid == false)
                    this.Cell.Interior.Color = XlRgbColor.rgbRed;
            }
        }
        private bool _isReadOnly;
        public bool IsReadOnly
        {
            get { return _isReadOnly; }
            set { _isReadOnly = value; }
        }
        private IObservable<PropertyChangeState> _owner;

        public IObservable<PropertyChangeState> Owner
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
            if (_valueType == typeof(double) || _valueType == typeof(decimal))
            {
                Char separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0];
                CellNumberFormat =  $"0.00";
            }
            if (_valueType == typeof(int))
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
        private string GetNumberFormat(decimal val)
        {
            int int_num_part = (int)val;
            var fractional_num_part = val - int_num_part;
            string str  =fractional_num_part.ToString();
            string out_str = "0.";
            int ii = 2;
            try
            {
                while (str.Length >= 2 && ii< str.Length && str[ii] == '0')
                {
                    out_str = $"{out_str}0";
                    ii++;
                }
            }
            catch
            {

            }

            if (ii == 2)
                return "#0.00";
            return out_str+'0';
        }
        
        public void OnNext(PropertyChangeState value)
        {
            try
            { 
            var sender = value.Sender;
            string[] prop_chain = this.ProprertyName.Split('.');
            if (prop_chain[0] != value.PropertyName) return;

            var prop_names = value.PropertyName.Split(new char[] { '.' });
            Type prop_type = sender.GetType().GetProperty(prop_names[0]).PropertyType;

            foreach (string prop_name in prop_chain)
            {
                var prop_val = sender.GetType().GetProperty(prop_name).GetValue(sender, null);
                if (prop_val is IExcelBindableBase exbb_val)
                    sender = exbb_val;
                else if (prop_val.GetType() == this.ValueType && IsReadOnly == false)
                {
                    if(prop_val is decimal dec_val)
                    {
                        this.Cell.NumberFormat= GetNumberFormat(dec_val);//
                    }
                    this.Cell.Value = prop_val;
                }
            }
            this.IsValid = value.PropertyIsValid;
            }
            catch(Exception e)
            {
                throw new Exception($"{e.Message}\n Строка:{this.Row} Столбец: {this.Column} \n Свойство:{value.PropertyName} Корректность записи:{value.PropertyIsValid}");;  
            }
        }
      
        private void GetPropValue(IExcelBindableBase obj, string prop_name, bool first_itaration = true)
        {


        }

        public void OnError(Exception error)
        {
            throw new NotImplementedException();
        }

        public void OnCompleted()
        {
            //Row = 0;
            //Column = 0;
            //Worksheet = null;
            //ProprertyName = null;
            //ValidateValueCallBack = null;
            //CoerceValueCallback = null;
            //ValueType = null;
        }

        public ExcelPropAddress(int row, int column, Excel.Worksheet worksheet, Type val_type,
            string prop_name = "", bool read_only = false,
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
            IsReadOnly = read_only;
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
