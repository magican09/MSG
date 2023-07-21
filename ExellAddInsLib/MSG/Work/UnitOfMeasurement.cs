namespace ExellAddInsLib.MSG
{
    public class UnitOfMeasurement : BindableBase, INameable
    {
        private int _number;

        public int Number
        {
            get { return _number; }
            set { SetProperty(ref _number, value); }
        }

        private string _name;

        public string Name
        {
            get { return _name; }
            set { SetProperty(ref _name, value); }
        }//Наименование


        public UnitOfMeasurement(string name)
        {
            Name = name;
        }
        public UnitOfMeasurement(int number, string name) : this(name)
        {
            Number = number;
        }

    }
}
