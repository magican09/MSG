using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public  class WorkScheduleChunk
    {
        private DateTime _startTime;
        public DateTime StartTime
        {
            get { return _startTime; }
            set {  _startTime =  value; }
        }//Дата начала
        private DateTime? _endTime;
        public DateTime? EndTime
        {
            get { return _endTime; }
            set {  _endTime =  value; }
        }//Дата окончания
        public WorkScheduleChunk(DateTime start_time, DateTime ent_time)
        {
            StartTime = start_time;
            EndTime = ent_time;
        }
    }
}
