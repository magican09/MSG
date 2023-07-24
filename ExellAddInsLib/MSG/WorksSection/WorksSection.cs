using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG.Section
{
    public  class WorksSection: ExcelBindableBase
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
        public ObservableCollection<MSGWork> MSGWorks { get; private set; } = new ObservableCollection<MSGWork>();
    }
}
