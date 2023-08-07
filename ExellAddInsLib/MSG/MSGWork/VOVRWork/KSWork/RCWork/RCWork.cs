namespace ExellAddInsLib.MSG
{
    public class RCWork : Work
    {
        private string _code;

        public string Code
        {
            get { return _code; }
            set { SetProperty(ref _code, value); }
        }
        private decimal _labournessCoefficient;

        public decimal LabournessCoefficient
        {
            get { return _labournessCoefficient; }
            set { SetProperty(ref _labournessCoefficient, value); }
        }

        public override object Clone()
        {
            RCWork new_work = (RCWork)base.Clone();
            new_work.Code = Code;
            new_work.LabournessCoefficient = LabournessCoefficient;
            return new_work;
        }
    }
}
