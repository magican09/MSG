using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public  class RCWork:KSWork
    {
		private decimal _labournessCoefficient;

		public decimal LabournessCoefficient
        {
			get { return _labournessCoefficient; }
			set { _labournessCoefficient = value; }
		}

	}
}
