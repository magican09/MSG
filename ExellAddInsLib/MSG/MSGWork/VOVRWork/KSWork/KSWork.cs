namespace ExellAddInsLib.MSG
{
    public class KSWork : Work
    {
        private string _code;

        public string Code
        {
            get { return _code; }
            set { SetProperty(ref _code, value); }
        }
        private ExcelNotifyChangedCollection<RCWork> _rCWorks = new ExcelNotifyChangedCollection<RCWork>();

        public ExcelNotifyChangedCollection<RCWork> RCWorks
        {
            get { return _rCWorks; }
            set { SetProperty(ref _rCWorks, value); }
        }
        public KSWork() : base()
        {

        }
        public override void SetSectionNumber(string section_number)
        {
            base.SetSectionNumber(section_number);
            foreach (RCWork work in this.RCWorks)
            {
                work.SetSectionNumber(section_number);
               if (work.ReportCard!=null)
                    work.ReportCard.Number = work.Number;
            }

        }
        new public object Clone()
        {
            KSWork new_obj = (KSWork)base.Clone();
        //    new_obj.UnitOfMeasurement = (UnitOfMeasurement)UnitOfMeasurement.Clone();
          //  new_obj.RCWorks = (ExcelNotifyChangedCollection<RCWork>)this.RCWorks.Clone();

            return new_obj;
        }
    }
}
