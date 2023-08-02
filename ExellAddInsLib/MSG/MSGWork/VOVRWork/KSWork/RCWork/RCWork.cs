using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public  class RCWork: Work
    {
        private string _code;

        public string Code
        {
            get { return _code; }
            set { SetProperty(ref _code, value); }
        }
        private decimal _labournessCoefficient;

		public decimal LabournessCoefficient
        {
			get { return _labournessCoefficient; }
			set { SetProperty(ref _labournessCoefficient, value); }
		}

	}
}
