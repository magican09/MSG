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

        [NonRegisterInUpCellAddresMap]
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
                if (work.ReportCard != null)
                    work.ReportCard.Number = work.Number;
            }

        }

        public override object Clone()
        {
            KSWork new_work = (KSWork)base.Clone();
            new_work.Code = Code;
            new_work.RCWorks = (ExcelNotifyChangedCollection<RCWork>)this.RCWorks.Clone();
            new_work.RCWorks.Owner = new_work;
            return new_work;
        }
    }
}
