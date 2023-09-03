using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public class ExellCellSubsciption : IDisposable
    {
        public  IObserver<PropertyChangeState> Observer { get; private set; }
        public IObservable<PropertyChangeState> Observable { get; private set; }
        List<IObserver<PropertyChangeState>> _observers;
        public ExellCellSubsciption(IObserver<PropertyChangeState> observer,IObservable<PropertyChangeState> observable,
            List<IObserver<PropertyChangeState>> observers)
        {
            Observer = observer;
            Observable = observable;
            _observers=observers;
        }
        public void Dispose()
        {
           if(_observers.Contains(Observer))
                _observers.Remove(Observer);
        }
    }
}
