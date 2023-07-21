namespace ExellAddInsLib.MSG
{
    public class Employer : Person
    {
        private Post _post;

        public Post Post
        {
            get { return _post; }
            set { _post = value; }
        }
        public Employer(int number, string name, Post post) : base(number, name)
        {
            Post = post;
        }
    }
}
