namespace ExellAddInsLib.MSG
{
    public class Post:ExcelBindableBase
    {
        private string _number;

        public string Number
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
        public Post(string number, string name)
        {
            Number = number;
            Name = name;
        }
        public Post()
        {

        }
    }
}
