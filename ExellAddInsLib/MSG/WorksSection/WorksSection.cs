namespace ExellAddInsLib.MSG.Section
{
    public class WorksSection : ExcelBindableBase
    {
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
        /// <summary>
        /// Коллекция с работами типа МСГ модели
        /// </summary>
        public ExcelNotifyChangedCollection<MSGWork> MSGWorks { get; private set; } = new ExcelNotifyChangedCollection<MSGWork>();

        new public object Clone()
        {
            WorksSection new_obj = (WorksSection)this.Clone();
            new_obj.MSGWorks = (ExcelNotifyChangedCollection<MSGWork>)this.MSGWorks.Clone();
            return new_obj;
        }

    }
}
