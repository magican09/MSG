using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class WorkReportCard : ExcelNotifyChangedCollection<WorkDay>
    {
        public const int WRC_DATE_ROW = 6;

        public const int WRC_NUMBER_COL = RCWork.RC_LABOURNESS_COL + 1;
        public const int WRC_PC_QUANTITY_COL = WRC_NUMBER_COL + 1;
        public const int WRC_DATE_COL = WRC_PC_QUANTITY_COL + 1;

        private string _number;

        public string Number
        {
            get { return _number; }
            set { SetProperty(ref _number, value); }
        }//Номер работы

        private decimal _quantity;
        [NonGettinInReflection]
        public decimal Quantity
        {
            get
            {
                decimal out_value = 0;
                foreach (WorkDay work_day in this)
                    out_value += work_day.Quantity;
                _quantity = out_value;
                return _quantity;
            }

        }//Выполенный объем работ

        private decimal _previousComplatedQuantity;

        public decimal PreviousComplatedQuantity
        {
            get { return _previousComplatedQuantity; }
            set { SetProperty(ref _previousComplatedQuantity, value); }
        }//Ранее выполненые объемы
        private IWork _owner;

        public IWork Owner
        {
            get { return _owner; }
            set { _owner = value; }
        }

        public override void SetStyleFormats(int col)
        {

            var cr_range = this.GetRange();
            if (cr_range != null)
            {
                cr_range.Interior.ColorIndex = col;
                cr_range.SetBordersLine(XlLineStyle.xlDashDotDot, XlLineStyle.xlContinuous, XlLineStyle.xlContinuous, XlLineStyle.xlContinuous);

                int max_day_number = WRC_PC_QUANTITY_COL + 30 + this.Count;

                if (this?.Owner?.Owner?.Owner?.Owner?.Owner?.Owner is MSGExellModel model)
                    max_day_number = (model.WorksEndDate - model.WorksStartDate).Days;

                Excel.Range days_row_range = this.Worksheet.Range[
                       this.Worksheet.Cells[cr_range.Row, WRC_PC_QUANTITY_COL + 1],
                       this.Worksheet.Cells[cr_range.Row, WRC_PC_QUANTITY_COL + 1+ max_day_number]];
                days_row_range.Interior.ColorIndex = col;
                days_row_range.Borders.LineStyle = Excel.XlLineStyle.xlDashDotDot;

                days_row_range.SetBordersLine(XlLineStyle.xlDashDot, XlLineStyle.xlDashDot,
                                                  XlLineStyle.xlContinuous, XlLineStyle.xlContinuous);

            }

            //if (this.Count>0)
            //{
            //    Excel.Range last_day_range = this.OrderBy(d => d.Date).LastOrDefault().GetRange();
            //    Excel.Range days_row_range = Worksheet.Application.Union(this.GetRange(), last_day_range);
            //    days_row_range.Interior.ColorIndex = col;
            //    days_row_range.Borders.LineStyle = Excel.XlLineStyle.xlDashDotDot;

            //    days_row_range.SetBordersLine(XlLineStyle.xlDashDot, XlLineStyle.xlDashDot,
            //                                      XlLineStyle.xlContinuous, XlLineStyle.xlContinuous);
            //}

        }
    }
}
