using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    public class AdjustableCollection<T> : ExcelNotifyChangedCollection<T> where T : IExcelBindableBase
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


            if (item.Number != null && this.Owner!=null)
            {
                item.Owner = this.Owner;
                var _subsequent_items = this.Where(itm => this.IndexOf(itm) >= index).ToList();
                var _previous_items = this.Where(itm => this.IndexOf(itm) < index).ToList();

                int item_suffix_num = _previous_items.Count+1;

                string item_number = $"{item.Owner.Number}.{item_suffix_num}";
                string[] item_numbers  = item_number.Split('.');
                int num_loc_indx = 0;
                foreach (string num in item_numbers)
                    item.SetNumberItem(num_loc_indx++, num);
         
                string[] item_prefix_numbers = item.NumberPrefix.Split('.');
                foreach (T itm in _subsequent_items)
                {
                    item_suffix_num++;
                    int subsq_itm_indx = 0;
                    foreach (string num in item_prefix_numbers)
                        itm.SetNumberItem(subsq_itm_indx++, num);
                    itm.SetNumberItem(subsq_itm_indx, (item_suffix_num).ToString());
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
