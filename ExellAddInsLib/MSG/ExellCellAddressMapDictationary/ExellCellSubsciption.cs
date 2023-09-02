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
        public ExellCellSubsciption(IObserver<PropertyChangeState> observer,IObservable<PropertyChangeState> observable)
        {
            Observer = observer;
            Observable = observable;
        }
        public void Dispose()
        {
           
        }
    }
}
