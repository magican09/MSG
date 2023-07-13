using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public class MSGWork:Work
    {
        

        private WorkSchedule _workSchedules = new WorkSchedule();

        public WorkSchedule WorkSchedules
        {
            get { return _workSchedules; }
            set { _workSchedules = value; }
        }

        private ObservableCollection<VOVRWork> _vOVRWorks = new ObservableCollection<VOVRWork>();

        public ObservableCollection<VOVRWork> VOVRWorks
        {
            get { return _vOVRWorks; }
            set { _vOVRWorks = value; }
        }

    }
}
