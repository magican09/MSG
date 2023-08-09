using System;

namespace ExellAddInsLib.MSG
{
    public class NeedsOfMachineDay : ExcelBindableBase
    {
        private DateTime _date;

        public DateTime Date
        {
            get { return _date; }
            set { SetProperty(ref _date, value); }
        }

        private decimal _quantity;

        public decimal Quantity
        {
            get { return _quantity; }
            set { SetProperty(ref _quantity, value); }
        }
        public NeedsOfMachineDay()
        {

        }
        public NeedsOfMachineDay(NeedsOfMachineDay day)
        {
            Date = day.Date;
            Quantity = day.Quantity;
        }
    }
}
