using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG 
{
    public  class NeedsOfWorker:Post
    {
        private int _quantity;

        public int Quantity
        {
            get { return _quantity; }
            set { SetProperty(ref _quantity, value); }
        }
        private NeedsOfWorkersReportCard _needsOfWorkersReportCard = new NeedsOfWorkersReportCard();

        public NeedsOfWorkersReportCard NeedsOfWorkersReportCard
        {
            get { return _needsOfWorkersReportCard; }
            set { _needsOfWorkersReportCard = value; }
        }
        private IWork _owner;

        public IWork Owner
        {
            get { return _owner; }
            set {
                SetProperty(ref _owner, value);

            }
        }

        public NeedsOfWorker()
        {

        }
        public NeedsOfWorker(string number, string name) : base(number, name)
        {
          
        }

    }
}
