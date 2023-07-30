using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG 
{
    public class WorkColletion:ObservableCollection<IWork>
    {
        protected override void ClearItems()
        {
            //foreach (IWork item in this)
            //    item.Parent = null;
            base.ClearItems();
        }
        protected override void InsertItem(int index, IWork item)
        {
            //item.Parent = this.Owner;
            base.InsertItem(index, item);

        }
        protected override void RemoveItem(int index)
        {
           // this[index].Parent = null;
            base.RemoveItem(index);
        }
    }
}
