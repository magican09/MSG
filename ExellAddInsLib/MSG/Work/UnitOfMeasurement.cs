using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG 
{
    public  class UnitOfMeasurement
    {
        private string _name;

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }//Наименование

        private string _shortName;
        public string ShortName
        {
            get { return _shortName; }
            set {  _shortName =  value; }
        }
        private string _fullName;
        public string FullName
        {
            get { return _fullName; }
            set {  _fullName= value; }
        }

        public UnitOfMeasurement(string name)
        {
            Name = name;
        }

    }
}
