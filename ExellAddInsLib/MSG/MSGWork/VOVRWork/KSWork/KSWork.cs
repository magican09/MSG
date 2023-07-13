using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public class KSWork:Work
    {
		private string _code;

		public string Code
		{
			get { return _code; }
			set { _code = value; }
		}

	}
}
