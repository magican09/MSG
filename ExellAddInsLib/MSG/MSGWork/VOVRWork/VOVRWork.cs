namespace ExellAddInsLib.MSG
{
    public class VOVRWork : Work
    {
        private ExcelNotifyChangedCollection<KSWork> _kSWorks = new ExcelNotifyChangedCollection<KSWork>();
       
        [NonRegisterInUpCellAddresMap]
        public ExcelNotifyChangedCollection<KSWork> KSWorks
        {
            get { return _kSWorks; }
            set { SetProperty(ref _kSWorks, value); }
        }

        public VOVRWork() : base()
        {

        }
       
        public override void SetSectionNumber(string section_number)
        {
            base.SetSectionNumber(section_number);
            foreach (KSWork work in this.KSWorks)
            {
                work.SetSectionNumber(section_number);
            }

        }
        public override object Clone()
        {
            VOVRWork new_work = (VOVRWork)base.Clone();
            new_work.KSWorks = (ExcelNotifyChangedCollection<KSWork>)this.KSWorks.Clone();
            new_work.KSWorks.Owner = new_work;
            return new_work;
        }
    }
}
