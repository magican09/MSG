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
    }
}
