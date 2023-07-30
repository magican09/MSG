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
        new public object Clone()
        {
            VOVRWork new_obj = (VOVRWork)this.Clone();
            new_obj.UnitOfMeasurement = (UnitOfMeasurement)UnitOfMeasurement.Clone();
            new_obj.KSWorks = (ExcelNotifyChangedCollection<KSWork>)this.KSWorks.Clone();
            return new_obj;
        }
    }
}
