using System;

namespace ExellAddInsLib.MSG
{
    public class WorkSchedule : ExcelNotifyChangedCollection<WorkScheduleChunk>
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
            get
            {
                _startDate = DateTime.MaxValue;
                foreach (WorkScheduleChunk chunk in this)
                    if (_startDate > chunk.StartTime)
                        _startDate = chunk.StartTime;
                return _startDate;
            }

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
        public override object Clone()
        {
            //     var new_works_shedules = new WorkSchedule();

            var new_ob = (WorkSchedule)base.Clone();
            //    new_works_shedules. = new ExcelNotifyChangedCollection<WorkScheduleChunk>();

            //   new_ob.WorkerNumber = this.WorkerNumber;
            return new_ob;
        }
    }
}
