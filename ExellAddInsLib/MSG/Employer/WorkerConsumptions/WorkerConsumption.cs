using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG 
{
    public  class WorkerConsumption : Post
    {
        private int _quantity;

        public int Quantity
        {
            get { return _quantity; }
            set { SetProperty(ref _quantity, value); }
        }
        private WorkersConsumptionReportCard _workersConsumptionReportCard = new WorkersConsumptionReportCard();

        public WorkersConsumptionReportCard WorkersConsumptionReportCard
        {
            get { return _workersConsumptionReportCard; }
            set { _workersConsumptionReportCard = value; }
        }
        private IWork _owner;

        public IWork Owner
        {
            get { return _owner; }
            set { SetProperty(ref _owner, value);

            }
        }

        public WorkerConsumption()
        {

        }
        public WorkerConsumption(string number, string name) : base(number, name)
        {
          
        }

    }
}
