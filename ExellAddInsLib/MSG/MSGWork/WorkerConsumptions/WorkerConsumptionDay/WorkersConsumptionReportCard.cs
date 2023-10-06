using System;

namespace ExellAddInsLib.MSG
{
    public class WorkersConsumptionReportCard : ExcelNotifyChangedCollection<WorkerConsumptionDay>
    {
      

        private DateTime _daysFirsDate;

        public DateTime DaysFirsDate
        {
            get { return _daysFirsDate; }
            set { SetProperty(ref _daysFirsDate, value); }
        }

        public override int AdjustExcelRepresentionTree(int row, int col = 0)
        {

            base.ChangeTopRow(row);

            foreach (var w_day in this)
            {
                try
                {
                    int d_col = (w_day.Date - this.DaysFirsDate).Days;
                    w_day.ChangeTopRow(row);
                    w_day.ChangeLeftColumn(WorkerConsumption.W_CONSUMPTIONS_FIRST_DATE_COL + d_col);

                }
                catch
                {

                }
            }
            return row;
        }
    }
}
