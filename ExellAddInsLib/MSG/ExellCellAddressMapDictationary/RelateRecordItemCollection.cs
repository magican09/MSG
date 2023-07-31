using System.Collections.ObjectModel;

namespace ExellAddInsLib.MSG
{
    public class RelateRecordItemCollection : ObservableCollection<RelateRecord>
    {
        private RelateRecord _owner;

        public RelateRecord Owner
        {
            get { return _owner; }
            set { _owner = value; }
        }
        public RelateRecordItemCollection(RelateRecord owner)
        {
            Owner = owner;
        }
        protected override void SetItem(int index, RelateRecord item)
        {
            item.Parent = this.Owner;
            base.SetItem(index, item);
        }
        protected override void ClearItems()
        {
            foreach (RelateRecord item in this)
                item.Parent = null;
            base.ClearItems();
        }
        protected override void InsertItem(int index, RelateRecord item)
        {
            item.Parent = this.Owner;
            base.InsertItem(index, item);

        }
        protected override void RemoveItem(int index)
        {
            this[index].Parent = null;
            base.RemoveItem(index);
        }

    }
}
