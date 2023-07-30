using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class ExcelNotifyChangedCollection<T> : ObservableCollection<T>, IExcelBindableBase
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private Guid _id = Guid.NewGuid();

        public Guid Id
        {
            get { return _id; }
        }
        public ObservableCollection<IExcelBindableBase> Owners { get; set; } = new ObservableCollection<IExcelBindableBase>();
        public void PropertyChange(object sender, string property_name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property_name));
        }
        public void SetProperty<T>(ref T member, T new_val, [CallerMemberName] string property_name = "")
        {
            if (new_val is IExcelBindableBase excell_bindable_new_val && !excell_bindable_new_val.Owners.Contains(this))
                excell_bindable_new_val.Owners.Add(this);
            if (member is IExcelBindableBase excell_bindable_member && excell_bindable_member.Owners.Contains(this))
                excell_bindable_member.Owners.Remove(this);
            member = new_val;
            PropertyChange(this, property_name);

        }



        public ExellCellAddressMapDictationary CellAddressesMap { get; set; } = new ExellCellAddressMapDictationary();

        public ExcelNotifyChangedCollection()
        {
            CellAddressesMap.AddEvent += OnCellAdressAdd;
            CellAddressesMap.OnSetWorksheet += OnCellAddressesMapWorksheet_Change;
            this.CollectionChanged += OnElementCollectionChanged;
        }

        private void OnElementCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Add)
                foreach (T element in e.NewItems)
                {
                    if (element is IExcelBindableBase excel_bindable_element && !excel_bindable_element.Owners.Contains(this))
                        excel_bindable_element.Owners.Add(this);
                }
            if (e.Action == NotifyCollectionChangedAction.Remove)
                foreach (T element in e.OldItems)
                {
                    if (element is IExcelBindableBase excel_bindable_element && excel_bindable_element.Owners.Contains(this))
                        excel_bindable_element.Owners.Remove(this);
                }
        }


        private void OnCellAddressesMapWorksheet_Change(Worksheet worksheet)
        {
            foreach (T element in this)
            {
                if (element is IExcelBindableBase excel_bindable_element)
                    excel_bindable_element.CellAddressesMap.SetWorksheet(worksheet);
            }
        }

        private void OnCellAdressAdd(ExellCellAddressMapDictationary.AddEventArgs pAddEventArgs)
        {
            if (pAddEventArgs != null)
            {
                Excel.Worksheet worksheet = pAddEventArgs.Value.Worksheet;
                worksheet.Cells[pAddEventArgs.Value.Row, pAddEventArgs.Value.Column].Interior.Color
                                = XlRgbColor.rgbAliceBlue;
            }
        }

        public object Clone()
        {
            //var new_collecion = this.MemberwiseClone();
            var new_collecion = new ExcelNotifyChangedCollection<T>();
            foreach (T element in this)
                if (element is ICloneable clanable_element)
                    new_collecion.Add((T)clanable_element.Clone());
            return new_collecion;
        }
    }
}
