namespace ExellAddInsLib.MSG
{
    public class WorkerConsumption : Post
    {
        public const int W_CONSUMPTIONS_FIRST_ROW_INDEX = 4;
        public const int W_CONSUMPTIONS_NUMBER_COL = 1;
        public const int W_CONSUMPTIONS_NAME_COL = 2;
        public const int W_CONSUMPTIONS_DATE_RAW = 3;
        public const int W_CONSUMPTIONS_FIRST_DATE_COL = 3;

        private decimal _quantity;

        public decimal Quantity
        {
            get { return _quantity; }
            set { SetProperty(ref _quantity, value); }
        }
        private WorkersConsumptionReportCard _workersConsumptionReportCard = new WorkersConsumptionReportCard();

        public WorkersConsumptionReportCard WorkersConsumptionReportCard
        {
            get { return _workersConsumptionReportCard; }
            set { _workersConsumptionReportCard = value; }
        }
        private IWork _owner;

        public IWork Owner
        {
            get { return _owner; }
            set
            {
                SetProperty(ref _owner, value);

            }
        }

        public WorkerConsumption()
        {

        }
        public WorkerConsumption(string number, string name) : base(number, name)
        {

        }

    }
}
