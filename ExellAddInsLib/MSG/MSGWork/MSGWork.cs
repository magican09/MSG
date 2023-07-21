using System.Collections.ObjectModel;

namespace ExellAddInsLib.MSG
{
    public class MSGWork : Work
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
        public int? GetShedulesAllDaysNumber()
        {
            if (WorkSchedules.Count > 0)
            {
                // var time_span = new TimeSpan(
                int? days_count = 0;
                foreach (WorkScheduleChunk chunk in this.WorkSchedules)
                {
                    days_count += (chunk.EndTime - chunk.StartTime)?.Days;
                }
                //var time_span = WorkSchedules[WorkSchedules.Count - 1].EndTime - WorkSchedules[0].StartTime;
                return days_count;
            }
            else
                return 0;
        }


    }
}
