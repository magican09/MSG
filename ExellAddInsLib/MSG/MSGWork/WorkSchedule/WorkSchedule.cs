using System;
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
        private DateTime _startDate;

        public DateTime StartDate
        {
            get {
                _startDate = DateTime.MaxValue;
                foreach (WorkScheduleChunk chunk in this)
                    if (_startDate > chunk.StartTime)
                        _startDate = chunk.StartTime;
                return _startDate; }
            
        }
        private DateTime _endDate;

        public DateTime EndDate
        {
            get
            {
                _endDate = DateTime.MinValue;
                foreach (WorkScheduleChunk chunk in this)
                    if (_endDate < chunk.EndTime)
                        _endDate = chunk.EndTime;
                return _endDate;
            }

        }

    }
}
