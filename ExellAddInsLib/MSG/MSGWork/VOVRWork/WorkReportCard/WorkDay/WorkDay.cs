using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public class WorkDay
    {
		private DateTime _date;

		public DateTime Date
		{
			get { return _date; }
			set { _date = value; }
		}

		private decimal _quantity;

		public decimal Quantity
        {
			get { return _quantity; }
			set { _quantity = value; }
		}


	}
}
