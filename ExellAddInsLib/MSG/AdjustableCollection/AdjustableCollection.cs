using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public class AdjustableCollection<T> : ExcelNotifyChangedCollection<T> where T:IExcelBindableBase
    {
        protected override void SetItem(int index, T item)
        {

            base.SetItem(index, item);
        }
        protected override void ClearItems()
        {

            base.ClearItems();
        }
        protected override void InsertItem(int index, T item)
        {


            if (item.Number != null)
            {
                item.Owner = this.Owner;
                var _subsequent_items = this.Where(itm => this.IndexOf(itm) > index).ToList();
                int ii = index + 2;

                int number_level = item.Number.Split('.').Length - 1;
                if (_subsequent_items.Count > 0)
                    item.SetNumberItem(number_level, ii.ToString());
              
                foreach (T itm in _subsequent_items)
                {
                    ii++;
                    itm.SetNumberItem(number_level, ii.ToString());
                 
                }
            }

            base.InsertItem(index, item);

        }
        protected override void RemoveItem(int index)
        {

            base.RemoveItem(index);
        }
    }
}
