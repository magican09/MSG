using System;
using System.Linq;

namespace ExellAddInsLib.MSG
{
    public class WorkSchedule : AdjustableCollection<WorkScheduleChunk>
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
        public bool IsIntersections(WorkScheduleChunk chunk)
        {
            if (this.Where(sch => sch.StartTime <= chunk.StartTime && sch.EndTime >= chunk.EndTime).Any())
                return true;
            else
                return false;
        }

        public int? GetShedulesAllDaysNumber()
        {
            if (this.Count > 0)
            {
                // var time_span = new TimeSpan(
                bool is_sunday_vocation = true;
                int? days_count = 0;
                foreach (WorkScheduleChunk chunk in this)
                {
                    if (chunk.IsSundayVacationDay == "Да")
                        is_sunday_vocation = true;
                    else
                        is_sunday_vocation = false;

                    int worked_day_number = 0;
                    for (DateTime date = chunk.StartTime; date <= chunk.EndTime; date = date.AddDays(1)) //Находим количество рабочих дней
                        if (is_sunday_vocation == false || date.DayOfWeek != DayOfWeek.Sunday)
                            worked_day_number++;

                    days_count += worked_day_number;// (chunk.EndTime - chunk.StartTime)?.Days;
                }
                //var time_span = WorkSchedules[WorkSchedules.Count - 1].EndTime - WorkSchedules[0].StartTime;
                return days_count;
            }
            else
                return 0;
        }
    }
}
