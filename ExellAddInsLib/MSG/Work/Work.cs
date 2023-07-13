using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public class Work : IWork
    {
        private string _number;

        public string Number
        {
            get { return _number; }
            set { _number = value; }
        }//Номер работы
        private string _name;

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }//Наименование работы
        private decimal _quantity;

        public decimal Quantity
        {
            get { return _quantity; }
            set { _quantity = value; }
        }//Выполенный объем работ

        private decimal _projectQuantity;

        public decimal ProjectQuantity
        {
            get { return _projectQuantity; }
            set { _projectQuantity = value; }
        }//Проектный объем работ

        private UnitOfMeasurement _unitOfMeasurement;

        public UnitOfMeasurement UnitOfMeasurement
        {
            get { return _unitOfMeasurement; }
            set { _unitOfMeasurement = value; }
        } //Ед. изм.
        private decimal _laboriousness;
        public decimal Laboriousness
        {
            get { return _laboriousness; }
            set { _laboriousness = value; }
        }//Трудоемкость  чел.час/ед.изм

        private WorkReportCard _reportCard;

        public WorkReportCard ReportCard
        {
            get { return _reportCard; }
            set { _reportCard = value; }
        }

    }
}
