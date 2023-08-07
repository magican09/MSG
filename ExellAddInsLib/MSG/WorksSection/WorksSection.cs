namespace ExellAddInsLib.MSG.Section
{
    public class WorksSection : ExcelBindableBase
    {
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
        /// <summary>
        /// Коллекция с работами типа МСГ модели
        /// </summary>
        public ExcelNotifyChangedCollection<MSGWork> MSGWorks { get; private set; } = new ExcelNotifyChangedCollection<MSGWork>();

        public override object Clone()
        {
            WorksSection new_obj = (WorksSection)base.Clone();
            new_obj.MSGWorks = (ExcelNotifyChangedCollection<MSGWork>)this.MSGWorks.Clone();
            new_obj.MSGWorks.Owner = new_obj;
            return new_obj;
        }

    }
}
