namespace ExellAddInsLib.MSG
{
    public class VOVRWork : Work
    {
        private ExcelNotifyChangedCollection<KSWork> _kSWorks = new ExcelNotifyChangedCollection<KSWork>();

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
        new public object Clone()
        {
            VOVRWork new_obj = (VOVRWork)base.Clone();
          //  new_obj.UnitOfMeasurement = (UnitOfMeasurement)UnitOfMeasurement.Clone();
          //  new_obj.KSWorks = (ExcelNotifyChangedCollection<KSWork>)this.KSWorks.Clone();
            return new_obj;
        }
    }
}
