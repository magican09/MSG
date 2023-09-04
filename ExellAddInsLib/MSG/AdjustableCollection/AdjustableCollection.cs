using System.Linq;

namespace ExellAddInsLib.MSG
{
    public class AdjustableCollection<T> : ExcelNotifyChangedCollection<T> where T : IObservableExcelBindableBase
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


            if (item.Number != null && this.Owner != null)
            {
                item.Owner = this.Owner;
                var _subsequent_items = this.Where(itm => this.IndexOf(itm) >= index).ToList();
                var _previous_items = this.Where(itm => this.IndexOf(itm) < index).ToList();
                if (index == this.Count - 1 && _subsequent_items.Count > 0)
                {
                    _previous_items.Add(_subsequent_items[0]);
                    _subsequent_items.Remove(_subsequent_items[0]);
                }

                int item_suffix_num = _previous_items.Count + 1;
                string item_number = "";
                if (item.Owner.Number != null && item.Owner.Number != "")
                    item_number = $"{item.Owner.Number}.{item_suffix_num}";
                else
                    item_number = $"{item_suffix_num}";
                string[] item_numbers = item_number.Split('.');
                int num_loc_indx = 0;
                foreach (string num in item_numbers)
                    item.SetNumberItem(num_loc_indx++, num);

                string[] item_prefix_numbers = { };

                if (item.NumberPrefix != null)
                    item_prefix_numbers = item.NumberPrefix.Split('.');


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
