using System;

namespace ExellAddInsLib.MSG
{
    public class WorkScheduleChunk : ExcelBindableBase
    {
        private string _number;

        public override string Number
        {
            get { return _number; }
            set { SetProperty(ref _number, value); }
        }//Номер работы

        private DateTime _startTime;
        public DateTime StartTime
        {
            get { return _startTime; }
            set { SetProperty(ref _startTime, value); }
        }//Дата начала
        private DateTime _endTime;
        public DateTime EndTime
        {
            get { return _endTime; }
            set { SetProperty(ref _endTime, value); }
        }//Дата окончания
        public WorkScheduleChunk(DateTime start_time, DateTime ent_time)
        {
            StartTime = start_time;
            EndTime = ent_time;
        }
        private int _workersNumber;

        public int WorkesNumber
        {
            get { return _workersNumber; }
            set { SetProperty(ref _workersNumber, value); }
        }
        public WorkScheduleChunk()
        {

        }
    }
}
