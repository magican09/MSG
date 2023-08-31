namespace ExellAddInsLib.MSG
{
    public class MachineConsumption : Machine
    {
        public const int MCH_CONSUMPTIONS_FIRST_ROW_INDEX = 4;
        public const int MCH_CONSUMPTIONS_NUMBER_COL = 1;
        public const int MCH_CONSUMPTIONS_NAME_COL = 2;
        public const int MCH_CONSUMPTIONS_DATE_RAW = 3;
        public const int MCH_CONSUMPTIONS_FIRST_DATE_COL = 3;

        private decimal _quantity;

        public decimal Quantity
        {
            get { return _quantity; }
            set { SetProperty(ref _quantity, value); }
        }
        private MachinesConsumptionReportCard _machinesConsumptionReportCard = new MachinesConsumptionReportCard();

        public MachinesConsumptionReportCard MachinesConsumptionReportCard
        {
            get { return _machinesConsumptionReportCard; }
            set { _machinesConsumptionReportCard = value; }
        }
        //private IWork _owner;

        //public IWork Owner
        //{
        //    get { return _owner; }
        //    set
        //    {
        //        SetProperty(ref _owner, value);

        //    }
        //}

        public MachineConsumption()
        {

        }
        public MachineConsumption(string number, string name) : base(number, name)
        {

        }

    }
}
