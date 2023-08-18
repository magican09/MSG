namespace ExellAddInsLib.MSG
{
    public class RelateRecord
    {
        private IExcelBindableBase _entity;

        public IExcelBindableBase Entity
        {
            get { return _entity; }
            set { _entity = value; }
        }

        private ExcelPropAddress _exellPropAddress;

        public ExcelPropAddress ExellPropAddress
        {
            get { return _exellPropAddress; }
            set { _exellPropAddress = value; }
        }


        private RelateRecord _parent;


        public RelateRecord Parent
        {
            get { return _parent; }
            set { _parent = value; }
        }
        private RelateRecordItemCollection _items;

        public RelateRecordItemCollection Items
        {
            get { return _items; }
            set { _items = value; }
        }
        private string _propertyName;

        public string PropertyName
        {
            get { return _propertyName; }
            set { _propertyName = value; }
        }

        public RelateRecord(IExcelBindableBase entity)
        {
            Entity = entity;
            _items = new RelateRecordItemCollection(this);
        }

    }

}
