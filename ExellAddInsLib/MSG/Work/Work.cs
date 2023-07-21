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

        private decimal _projectQuantity;

        public decimal ProjectQuantity
        {
            get { return _projectQuantity; }
            set { SetProperty(ref _projectQuantity, value); }
        }//Проектный объем работ

        private UnitOfMeasurement _unitOfMeasurement;

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

        private WorkReportCard _reportCard;

        public WorkReportCard ReportCard
        {
            get { return _reportCard; }
            set { SetProperty(ref _reportCard, value); }
        }

        private MSGExellModel _ownerExellModel;



        public MSGExellModel OwnerExellModel
        {
            get { return _ownerExellModel; }
            set { _ownerExellModel = value; }
        }

        private IWork _owner;

        public IWork Owner
        {
            get { return _owner; }
            set { _owner = value; }
        }


        private ObservableCollection<IWork> _children = new ObservableCollection<IWork>();

        public ObservableCollection<IWork> Children
        {
            get { return _children; }
            set { _children = value; }
        }
        public Work()
        {
            Children.CollectionChanged += OnChildrenAdd;
        }

        private void OnChildrenAdd(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Add)
            {
                foreach (IWork child in e.NewItems)
                    child.Owner = this;
            }
            if (e.Action == NotifyCollectionChangedAction.Remove)
            {
                foreach (IWork child in e.OldItems)
                    child.Owner = null;
            }
        }

        public void ClearCalculatesFields()
        {
            this.Laboriousness = 0;
            this.Quantity = 0;
        }
    }
}
