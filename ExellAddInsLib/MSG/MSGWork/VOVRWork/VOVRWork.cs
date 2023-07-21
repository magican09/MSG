using System.Collections.ObjectModel;

namespace ExellAddInsLib.MSG
{
    public class VOVRWork : Work
    {
        private ObservableCollection<KSWork> _kSWorks = new ObservableCollection<KSWork>();

        public ObservableCollection<KSWork> KSWorks
        {
            get { return _kSWorks; }
            set { SetProperty(ref _kSWorks, value); }
        }


    }
}
