using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public abstract class Work : ExcelBindableBase, IWork
    {
        private Excel.Worksheet _worksheet;

        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public override Excel.Worksheet Worksheet
        {
            get { return base.Worksheet; }
            set
            {

                this.WorkersComposition.Worksheet = value;
                this.MachinesComposition.Worksheet = value;
                if (this.ReportCard != null) this.ReportCard.Worksheet = value;
                base.Worksheet = value;
            }
        }

        private int _rowIndex;

        public int RowIndex
        {
            get { return _rowIndex; }
            set { SetProperty(ref _rowIndex, value); }
        }

        public Dictionary<string, int> PropertyColumnMap = new Dictionary<string, int>();


        private string _number;

        public override string Number
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

        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public WorkReportCard ReportCard
        {
            get { return _reportCard; }
            set
            {
                SetProperty(ref _reportCard, value);
                if (_reportCard != null)
                    _reportCard.Worksheet = this.Worksheet;
            }
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

        private MachinesComposition _machinesComposition;

        public MachinesComposition MachinesComposition
        {
            get { return _machinesComposition; }
            set { SetProperty(ref _machinesComposition, value); }
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
            WorkersComposition.Owner = this;
            MachinesComposition = new MachinesComposition();
            MachinesComposition.Owner = this;
            //  ReportCard = new WorkReportCard();
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
        public virtual void SetSectionNumber(string section_number)
        {
            Number = setSectionNumber(section_number, Number);

            foreach (var nw in this.WorkersComposition)
                nw.Number = setSectionNumber(section_number, nw.Number);
        }
        private string setSectionNumber(string section_number, string number)
        {
            number = number.Substring(number.IndexOf('.'), number.Length - number.IndexOf('.'));
            return section_number + number;
        }

        public override object Clone()
        {
            var new_work = (Work)base.Clone();
            new_work.UnitOfMeasurement = this.UnitOfMeasurement;
            new_work.WorkersComposition = (WorkersComposition)this.WorkersComposition.Clone();
            new_work.MachinesComposition = (MachinesComposition)this.MachinesComposition.Clone();

            return new_work;
        }
        public override Range GetRange()
        {
            Excel.Range range = base.GetRange();
            Excel.Range report_card_range = null;

            Excel.Range workers_composition_range = this.WorkersComposition.GetRange();
            Excel.Range machine_composition_range = this.MachinesComposition.GetRange();
            if (this.ReportCard != null && this.ReportCard.GetRange() != null)
            {
                report_card_range = this.ReportCard.GetRange();
                range = Worksheet.Application.Union(range, report_card_range);
            }
            range = Worksheet.Application.Union(new List<Excel.Range>() { range, workers_composition_range, machine_composition_range });

            return range;
        }

        public void ClearCalculatesFields()
        {
            this.Laboriousness = 0;
            this.Quantity = 0;
        }
    }
}
