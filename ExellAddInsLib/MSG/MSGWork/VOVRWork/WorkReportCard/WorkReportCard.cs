using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public class WorkReportCard:ObservableCollection<WorkDay>
    {
        private string _number;

        public string Number
        {
            get { return _number; }
            set { _number = value; }
        }//Номер работы
    }
}
