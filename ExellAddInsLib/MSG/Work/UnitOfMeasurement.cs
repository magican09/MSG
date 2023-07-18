using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG 
{
    public  class UnitOfMeasurement
    {
        private int _number;

        public int Number
        {
            get { return _number; }
            set { _number = value; }
        }

        private string _name;

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }//Наименование

   
        public UnitOfMeasurement(string name)
        {
            Name = name;
        }
        public UnitOfMeasurement(int number, string name):this(name)
        {
            Number = number;
        }

    }
}
