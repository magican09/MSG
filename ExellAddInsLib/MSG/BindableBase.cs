using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExellAddInsLib.MSG
{
    public class BindableBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void SetProperty<T>(ref T member, T new_val, [CallerMemberName] string property_name = "")
        {

            member = new_val;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property_name));

        }
    }
}
