namespace ExellAddInsLib.MSG
{
    public class Employer : Person
    {
        private Post _post;

        public Post Post
        {
            get { return _post; }
            set { SetProperty(ref _post, value); }
        }
        public Employer(string number, string name, Post post) : base(number, name)
        {
            Post = post;
        }
        public Employer()
        {

        }
    }
}
