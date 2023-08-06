using System;

namespace ExellAddInsLib.MSG
{
    public class MSGWork : Work
    {

        private bool _isSundayVocation = true;

        public bool IsSundayVocation
        {
            get { return _isSundayVocation; }
            set { _isSundayVocation = value; }
        }

        private WorkSchedule _workSchedules = new WorkSchedule();
       
       
        public WorkSchedule WorkSchedules
        {
            get { return _workSchedules; }
            set { _workSchedules = value; }
        }

        private ExcelNotifyChangedCollection<VOVRWork> _vOVRWorks = new ExcelNotifyChangedCollection<VOVRWork>();
      
        [NonRegisterInUpCellAddresMap]
        public ExcelNotifyChangedCollection<VOVRWork> VOVRWorks
        {
            get { return _vOVRWorks; }
            set { _vOVRWorks = value; }
        }
        public int? GetShedulesAllDaysNumber(bool is_sunday_vocation)
        {


            if (WorkSchedules.Count > 0)
            {
                // var time_span = new TimeSpan(
                int? days_count = 0;
                foreach (WorkScheduleChunk chunk in this.WorkSchedules)
                {
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
        public MSGWork() : base()
        {

        }
       
      
        public override object Clone()
        {
            MSGWork new_obj = (MSGWork)base.Clone();
            new_obj.WorkSchedules = (WorkSchedule)this.WorkSchedules.Clone();
            new_obj.VOVRWorks = (ExcelNotifyChangedCollection<VOVRWork>)this.VOVRWorks.Clone();
            new_obj.WorkSchedules.Owner = new_obj;
            new_obj.VOVRWorks.Owner = new_obj;
            return new_obj;
        }
    }
}
