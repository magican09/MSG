using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;

namespace ExellAddInsLib.MSG
{
    public abstract class Work : ExcelBindableBase, IWork
    {

        private int _rowIndex;

        public int RowIndex
        {
            get { return _rowIndex; }
            set { SetProperty(ref _rowIndex, value); }
        }

        public Dictionary<string, int> PropertyColumnMap = new Dictionary<string, int>();


        private string _number;

        public string Number
        {
            get { return _number; }
            set { SetProperty(ref _number, value); }
        }//Номер работы
        private string _name;

        public string Name
        {
            get { return _name; }
            set { SetProperty(ref _name, value); }
        }//Наименование работы
        private decimal _quantity;

        public decimal Quantity
        {
            get { return _quantity; }
            set { SetProperty(ref _quantity, value); }
        }//Выполенный объем работ

        private decimal _previousComplatedQuantity;

        public decimal PreviousComplatedQuantity
        {
            get { return _previousComplatedQuantity; }
            set { SetProperty(ref _previousComplatedQuantity, value); }
        }//Ранее выполненые объемы

        private decimal _projectQuantity;

        public decimal ProjectQuantity
        {
            get { return _projectQuantity; }
            set { SetProperty(ref _projectQuantity, value); }
        }//Проектный объем работ

        private UnitOfMeasurement _unitOfMeasurement;
        [DontClone]
        public UnitOfMeasurement UnitOfMeasurement
        {
            get { return _unitOfMeasurement; }
            set { SetProperty(ref _unitOfMeasurement, value); }
        } //Ед. изм.
        private decimal _laboriousness;
        public decimal Laboriousness
        {
            get { return _laboriousness; }
            set { SetProperty(ref _laboriousness, value); }
        }//Трудоемкость  чел.час/ед.изм

        private WorkReportCard _reportCard ;
      
        [NonGettinInReflection]
        [DontClone]
        [NonRegisterInUpCellAddresMap]
        public WorkReportCard ReportCard
        {
            get { return _reportCard; }
            set { SetProperty(ref _reportCard, value); }
        }
      
        //public WorkReportCard ReportCard { get; set; }
        private MSGExellModel _ownerExellModel;



        public MSGExellModel OwnerExellModel
        {
            get { return _ownerExellModel; }
            set { _ownerExellModel = value; }
        }

        private IWork _owner;

        //public IWork Owner
        //{
        //    get { return _owner; }
        //    set { _owner = value; }
        //}

        private WorkersComposition _workersComposition;

        public WorkersComposition WorkersComposition
        {
            get { return _workersComposition; }
            set { SetProperty(ref _workersComposition, value); }
        }



        private ObservableCollection<IWork> _children = new ObservableCollection<IWork>();

        public ObservableCollection<IWork> Children
        {
            get { return _children; }
            set { _children = value; }
        }
        public Work()
        {
            WorkersComposition = new WorkersComposition();
            ReportCard = new WorkReportCard();
            Children.CollectionChanged += OnChildrenAdd;

        }

        private void OnChildrenAdd(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Add)
            {
                //foreach (IWork child in e.NewItems)
                //    child.Owner = this;
            }
            if (e.Action == NotifyCollectionChangedAction.Remove)
            {
            //    foreach (IWork child in e.OldItems)
            //        child.Owner = null;
            }
        }
        public  virtual  void SetSectionNumber(string section_number)
        {
            Number = Number.Substring(Number.IndexOf('.'), Number.Length - Number.IndexOf('.'));
            Number = section_number + Number;
        }

      new  public object Clone()
        {
            var new_work =(IWork)this.MemberwiseClone();
            foreach(var kvp in this.CellAddressesMap)
            {
                new_work.CellAddressesMap.Add(kvp.Key, new ExellPropAddress(kvp.Value));
            }


            return new_work; 
        }

        public void ClearCalculatesFields()
        {
            this.Laboriousness = 0;
            this.Quantity = 0;
        }
    }
}
