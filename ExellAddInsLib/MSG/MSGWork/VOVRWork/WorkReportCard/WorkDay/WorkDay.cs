using System;

namespace ExellAddInsLib.MSG
{
    public class WorkDay : ExcelBindableBase
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

        private decimal _laborСosts;

        public decimal LaborСosts
        {
            get { return _laborСosts; }
            set { SetProperty(ref _laborСosts, value); }
        }

    }
}
