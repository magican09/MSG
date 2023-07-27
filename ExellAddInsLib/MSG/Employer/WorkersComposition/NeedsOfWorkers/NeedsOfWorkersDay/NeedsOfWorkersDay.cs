using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public  class NeedsOfWorkersDay:ExcelBindableBase
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
        public NeedsOfWorkersDay()
        {

        }
        public NeedsOfWorkersDay(NeedsOfWorkersDay day)
        {
            Date = day.Date;
            Quantity = day.Quantity;
        }
    }
}
