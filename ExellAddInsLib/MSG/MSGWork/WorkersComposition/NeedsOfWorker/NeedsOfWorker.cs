namespace ExellAddInsLib.MSG
{
    public class NeedsOfWorker : Post
    {
        private decimal _quantity;

        public decimal Quantity
        {
            get { return _quantity; }
            set { SetProperty(ref _quantity, value); }
        }
        private NeedsOfWorkersReportCard _needsOfWorkersReportCard;

        public NeedsOfWorkersReportCard NeedsOfWorkersReportCard
        {
            get { return _needsOfWorkersReportCard; }
            set { SetProperty(ref _needsOfWorkersReportCard, value); }
        }
        //[DontClone]
        // public IWork Owner { get; set; }

        public NeedsOfWorker() : base()
        {
            NeedsOfWorkersReportCard = new NeedsOfWorkersReportCard();
        }
        public NeedsOfWorker(string number, string name) : base(number, name)
        {
            NeedsOfWorkersReportCard = new NeedsOfWorkersReportCard();
        }

    }
}
