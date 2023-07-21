namespace ExellAddInsLib.MSG
{
    public class Person
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
        }
        public Person(int number, string name)
        {
            Number = number;

            Name = name;

        }
    }
}
