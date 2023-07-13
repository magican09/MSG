using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG 
{
    public  class VOVRWork:Work
    {
		private ObservableCollection<KSWork>  _kSWorks = new ObservableCollection<KSWork>();

		public ObservableCollection<KSWork> KSWorks
        {
			get { return _kSWorks; }
			set { _kSWorks = value; }
		}
		private WorkReportCard _workReportCard = new WorkReportCard();

		public WorkReportCard WorkReportCard
        {
			get { return _workReportCard; }
			set { _workReportCard = value; }
		}

	}
}
