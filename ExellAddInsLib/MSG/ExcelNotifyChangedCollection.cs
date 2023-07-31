using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
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
            {
                excell_bindable_new_val.Owners.Add(this);
                foreach (var kvp in excell_bindable_new_val.CellAddressesMap)
                {
                    string key_str = $"{excell_bindable_new_val.Id.ToString()}_{kvp.Value.ProprertyName}";
                    if (!this.CellAddressesMap.ContainsKey(key_str))
                        this.CellAddressesMap.Add(key_str, kvp.Value);
                }
                excell_bindable_new_val.CellAddressesMap.AddEvent += OnCellAdressAdd;
            }
            if (member is IExcelBindableBase excell_bindable_member && excell_bindable_member.Owners.Contains(this))
            {
                excell_bindable_member.Owners.Remove(this);
                foreach (var kvp in excell_bindable_member.CellAddressesMap)
                {
                    string key_str = $"{excell_bindable_member.Id.ToString()}_{kvp.Value.ProprertyName}";
                    if (this.CellAddressesMap.ContainsKey(key_str))
                        this.CellAddressesMap.Remove(key_str);
                }

                excell_bindable_member.CellAddressesMap.AddEvent -= OnCellAdressAdd;
            }
            member = new_val;
            PropertyChange(this, property_name);

        }



        public ExellCellAddressMapDictationary CellAddressesMap { get; set; }

        public ExcelNotifyChangedCollection()
        {
            CellAddressesMap = new ExellCellAddressMapDictationary();
            CellAddressesMap.Owner = this;
            CellAddressesMap.OnSetWorksheet += OnCellAddressesMapWorksheet_Change;
            this.CollectionChanged += OnElementCollectionChanged;
        }

        private void OnElementCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            //if (e.Action == NotifyCollectionChangedAction.Add)
            //    foreach (T element in e.NewItems)
            //    {if (element is MSGWork)
            //            ;
            //        if (element is IExcelBindableBase excel_bindable_element && !excel_bindable_element.Owners.Contains(this))
            //        {
            //            excel_bindable_element.Owners.Add(this);
            //            foreach (var kvp in excel_bindable_element.CellAddressesMap)
            //            { string key_str = $"{excel_bindable_element.Id.ToString()}_{kvp.Value.ProprertyName}";
            //                if (!this.CellAddressesMap.ContainsKey(key_str))
            //                    this.CellAddressesMap.Add(key_str, kvp.Value);
            //            }
            //            excel_bindable_element.CellAddressesMap.AddEvent += OnCellAdressAdd;
            //        }
            //    }
            //if (e.Action == NotifyCollectionChangedAction.Remove)
            //    foreach (T element in e.OldItems)
            //    {
            //        if (element is IExcelBindableBase excel_bindable_element && excel_bindable_element.Owners.Contains(this))
            //        {
            //            excel_bindable_element.Owners.Remove(this);
            //            foreach (var kvp in excel_bindable_element.CellAddressesMap)
            //            {
            //                string key_str = $"{excel_bindable_element.Id.ToString()}_{kvp.Value.ProprertyName}";
            //                if (this.CellAddressesMap.ContainsKey(key_str))
            //                    this.CellAddressesMap.Remove(key_str);
            //            }
            //            excel_bindable_element.CellAddressesMap.AddEvent -= OnCellAdressAdd;
            //        }
            //    }
        }

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
            if (item is IExcelBindableBase excel_bindable_element && !excel_bindable_element.Owners.Contains(this))
            {
                excel_bindable_element.Owners.Add(this);
                foreach (var kvp in excel_bindable_element.CellAddressesMap)
                {
                    string key_str = $"{excel_bindable_element.Id.ToString()}_{kvp.Value.ProprertyName}";
                    if (!this.CellAddressesMap.ContainsKey(key_str))
                        this.CellAddressesMap.Add(key_str, kvp.Value);
                }
                excel_bindable_element.CellAddressesMap.AddEvent += OnCellAdressAdd;
            }

            base.InsertItem(index, item);

        }
        protected override void RemoveItem(int index)
        {
            if (this[index] is IExcelBindableBase excel_bindable_element && excel_bindable_element.Owners.Contains(this))
            {
                excel_bindable_element.Owners.Remove(this);
                foreach (var kvp in excel_bindable_element.CellAddressesMap)
                {
                    string key_str = $"{excel_bindable_element.Id.ToString()}_{kvp.Value.ProprertyName}";
                    if (this.CellAddressesMap.ContainsKey(key_str))
                        this.CellAddressesMap.Remove(key_str);
                }
                excel_bindable_element.CellAddressesMap.AddEvent -= OnCellAdressAdd;
            }
            //  this[index].Parent = null;
            base.RemoveItem(index);
        }

        private void OnCellAddressesMapWorksheet_Change(Worksheet worksheet)
        {
            foreach (T element in this)
            {
                if (element is IExcelBindableBase excel_bindable_element)
                    excel_bindable_element.CellAddressesMap.SetWorksheet(worksheet);
            }
        }

        private void OnCellAdressAdd(IExcelBindableBase sender, ExellCellAddressMapDictationary.AddEventArgs pAddEventArgs)
        {
            if (pAddEventArgs != null)
            {
                Excel.Worksheet worksheet = pAddEventArgs.Value.Worksheet;
                string key_str = $"{sender.Id.ToString()}_{pAddEventArgs.Value.ProprertyName}";
                this.CellAddressesMap.Add(key_str, pAddEventArgs.Value);
            }
        }
        public Excel.Range GetRange(Excel.Worksheet worksheet, int right_border, int low_borde = 100000000, int left_border = 0, int up_border = 0)
        {
            Excel.Range range = null;
            var cell_maps = this.CellAddressesMap.Where(cm => cm.Value.Worksheet.Name == worksheet.Name);
            if (cell_maps.Any())
            {
                int upper_row = cell_maps.OrderBy(c => c.Value.Row).First().Value.Row;
                int lower_row = cell_maps.OrderBy(c => c.Value.Row).Last().Value.Row;
                int left_col = cell_maps.OrderBy(c => c.Value.Column).First().Value.Column;
                int right_col = cell_maps.OrderBy(c => c.Value.Column).Last().Value.Column;
                if (lower_row > low_borde) lower_row = low_borde;
                if (upper_row < up_border) upper_row = up_border;
                if (left_col < left_border) left_col = left_border;
                if (right_col > right_border) right_col = right_border;

                var left_upper_cell = worksheet.Cells[upper_row, left_col];
                var rigth_lower_cell = worksheet.Cells[lower_row, right_col]; ;

                range = worksheet.Range[left_upper_cell, rigth_lower_cell];
            }
            return range;
        }
        public Excel.Range GetRange(Excel.Worksheet worksheet)
        {
            Excel.Range range = null;
            var cell_maps = this.CellAddressesMap.Where(cm => cm.Value.Worksheet.Name == worksheet.Name);
            if (cell_maps.Any())
            {
                var left_upper_cell = cell_maps.OrderBy(c => c.Value.Row).OrderBy(c => c.Value.Column).First().Value.Cell;
                var rigth_lower_cell = cell_maps.OrderBy(c => c.Value.Row).OrderBy(c => c.Value.Column).Last().Value.Cell;
                range = worksheet.Range[left_upper_cell, rigth_lower_cell];
            }
            return range;
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
