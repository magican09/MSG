namespace ExellAddInsLib.MSG
{
    public class Person : ExcelBindableBase
    {
        private string _number;

        public string Number
        {
            get { return _number; }
            set { SetProperty(ref _number, value); }
        }
        private string _name;

        public string Name
        {
            get { return _name; }
            set { SetProperty(ref _name, value); }
        }
        public Person(string number, string name)
        {
            Number = number;

            Name = name;

        }
        public Person()
        {

        }
    }
}
