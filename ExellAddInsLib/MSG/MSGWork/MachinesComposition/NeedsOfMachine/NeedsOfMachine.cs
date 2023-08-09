namespace ExellAddInsLib.MSG
{
    public class NeedsOfMachine : Post
    {
        private decimal _quantity;

        public decimal Quantity
        {
            get { return _quantity; }
            set { SetProperty(ref _quantity, value); }
        }
        private NeedOfMachinesReportCard _needsOfMachinesReportCard;

        public NeedOfMachinesReportCard NeedsOfMachinesReportCard
        {
            get { return _needsOfMachinesReportCard; }
            set { SetProperty(ref _needsOfMachinesReportCard, value); }
        }
        //[DontClone]
        // public IWork Owner { get; set; }

        public NeedsOfMachine() : base()
        {
            NeedsOfMachinesReportCard = new NeedOfMachinesReportCard(); ;
        }
        public NeedsOfMachine(string number, string name) : base(number, name)
        {
            NeedsOfMachinesReportCard = new NeedOfMachinesReportCard();
        }

    }
}
