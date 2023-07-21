using System.Collections.ObjectModel;

namespace ExellAddInsLib.MSG
{
    public class WorkSchedule : ObservableCollection<WorkScheduleChunk>
    {
        private int _workerNumber;

        public int WorkerNumber
        {
            get { return _workerNumber; }
            set { _workerNumber = value; }
        }
    }
}
