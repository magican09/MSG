﻿using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class ExcelNotifyChangedCollection<T> : ObservableCollection<T>, IExcelNotifyChangedCollection
        where T : IExcelBindableBase, ICloneable
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private bool _isValid = true;

        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public bool IsValid
        {

            get
            {
                if (this.CellAddressesMap.Where(cm => cm.Value.IsValid == false).Any())
                    _isValid = false;
                return _isValid;
            }

            set { _isValid = value; }
        }

        public string Name { get; set; }
        private Guid _id = Guid.NewGuid();
        public virtual string Number { get; set; }
        public string NumberSuffix
        {
            get
            {

                try
                {
                    if (this.Number != null && this.Number.Contains('.'))
                    {
                        var str = this.Number.Split('.');
                        List<string> out_str_arr = new List<string>();
                        for (int ii = 0; ii < str.Length - 1; ii++)
                            out_str_arr.Add(str[ii]);

                        string out_str = "";
                        foreach (string s in out_str_arr)
                            out_str += s;
                        return out_str;
                    }
                    else return null;
                }
                catch
                {
                    return null;
                }

            }
        }
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
            if (new_val is IExcelBindableBase excell_bindable_new_val/* && !excell_bindable_new_val.Owners.Contains(this)*/)
            {
                this.RegisterNewValInCellAddresMap(excell_bindable_new_val, property_name);
            }
            if (member is IExcelBindableBase excell_bindable_member /*&& excell_bindable_member.Owners.Contains(this)*/)
            {
                this.UnregisterMemberValInCellAddresMap(excell_bindable_member, property_name);
            }
            member = new_val;
            PropertyChange(this, property_name);

        }
        private IExcelBindableBase _owner;

        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public IExcelBindableBase Owner
        {
            get { return _owner; }
            set
            {
                foreach (T elm in this)
                    elm.Owner = value;
                _owner = value;
            }
        }

        public ExellCellAddressMapDictationary CellAddressesMap { get; set; }

        public ExcelNotifyChangedCollection()
        {
            CellAddressesMap = new ExellCellAddressMapDictationary();
            CellAddressesMap.Owner = this;
            CellAddressesMap.OnSetWorksheet += OnCellAddressesMapWorksheet_Change;
        }

        protected override void SetItem(int index, T item)
        {

            base.SetItem(index, item);
        }
        protected override void ClearItems()
        {
            foreach (T element in this)
            {
                if (element is IExcelBindableBase excel_bindable_element/* && excel_bindable_element.Owners.Contains(this)*/)
                {
                    excel_bindable_element.Owner = null;
                    foreach (var kvp in excel_bindable_element.CellAddressesMap)
                    {
                        string key_str = $"{excel_bindable_element.Id.ToString()}_{kvp.Value.ProprertyName}";
                        if (this.CellAddressesMap.ContainsKey(key_str))
                            this.CellAddressesMap.Remove(key_str);
                    }
                    excel_bindable_element.CellAddressesMap.AddEvent -= OnCellAdressAdd;
                }
            }
            base.ClearItems();
        }
        protected override void InsertItem(int index, T item)
        {
            if (item is IExcelBindableBase excel_bindable_element/* && !excel_bindable_element.Owners.Contains(this)*/)
            {
                excel_bindable_element.Owner = this.Owner;
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
            if (this[index] is IExcelBindableBase excel_bindable_element/* && excel_bindable_element.Owners.Contains(this)*/)
            {
                excel_bindable_element.Owner = null;
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
                var lu = cell_maps.OrderBy(c => c.Value.Row).OrderBy(c => c.Value.Column).First().Value;
                var rl = cell_maps.OrderBy(c => c.Value.Column).OrderBy(c => c.Value.Row).Last().Value;
                var left_upper_cell = cell_maps.OrderBy(c => c.Value.Row).OrderBy(c => c.Value.Column).First().Value.Cell;
                var rigth_lower_cell = cell_maps.OrderBy(c => c.Value.Column).OrderBy(c => c.Value.Row).Last().Value.Cell;
                range = worksheet.Range[left_upper_cell, rigth_lower_cell];
            }
            return range;
        }
        public void SetInvalidateCellsColor(XlRgbColor color)
        {
            var invalide_cells = this.CellAddressesMap.Where(cm => cm.Value.IsValid == false);
            foreach (var kvp in invalide_cells)
            {
                kvp.Value.Cell.Interior.Color = color;
            }
        }
        public void ChangeTopRow(int row)
        {
            int top_row = this.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).First().Value.Row;
            int row_delta = row - top_row;
            if (top_row + row_delta <= 0) row_delta = 0;
            foreach (var kvp in this.CellAddressesMap)
            {
                kvp.Value.Row += row_delta;
            }
        }
        public int GetRowsCount()
        {
            int top_row = this.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).First().Value.Row;
            int bottom_row = this.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).Last().Value.Row;
            return bottom_row - top_row;
        }
        public int GetBottomRow()
        {
            int top_row = this.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).First().Value.Row;
            int bottom_row = this.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).Last().Value.Row;
            return bottom_row;
        }
        public int GetTopRow()
        {
            int top_row = this.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).First().Value.Row;
            return top_row;
        }
        private void RegisterNewValInCellAddresMap(IExcelBindableBase excell_bindable_new_val, string property_name)
        {
            //    excell_bindable_new_val.Owners.Add(this);
            var non_reg_in_upper_attribute = this.GetType().GetProperty(property_name).GetCustomAttribute(typeof(NonRegisterInUpCellAddresMapAttribute));
            if (non_reg_in_upper_attribute == null)
            {
                //  excell_bindable_new_val.Owners.Add(this);
                foreach (var kvp in excell_bindable_new_val.CellAddressesMap)
                {
                    string key_str = $"{excell_bindable_new_val.Id.ToString()}_{kvp.Value.ProprertyName}";
                    if (!this.CellAddressesMap.ContainsKey(key_str))
                        this.CellAddressesMap.Add(key_str, kvp.Value);
                }
                excell_bindable_new_val.CellAddressesMap.AddEvent += OnCellAdressAdd;
            }
        }
        private void UnregisterMemberValInCellAddresMap(IExcelBindableBase excell_bindable_member, string property_name)
        {
            //     excell_bindable_member.Owners.Remove(this);
            foreach (var kvp in excell_bindable_member.CellAddressesMap)
            {
                string key_str = $"{excell_bindable_member.Id.ToString()}_{kvp.Value.ProprertyName}";
                if (this.CellAddressesMap.ContainsKey(key_str))
                    this.CellAddressesMap.Remove(key_str);
            }

            excell_bindable_member.CellAddressesMap.AddEvent -= OnCellAdressAdd;
        }
        private List<IExcelBindableBase> nambered_objects = new List<IExcelBindableBase>();
        public void SetNumberItem(int possition, string number, bool first_itaration = true)
        {
            if (this.Number == null || !this.Number.Contains(".")) return;
            string[] str = this.Number.Split('.');
            str[possition] = number;
            string out_str = "";
            foreach (string s in str)
                out_str += $"{s}.";
            out_str = out_str.TrimEnd('.');
            this.Number = out_str;
            if (first_itaration) nambered_objects.Clear();
            var prop_infoes = this.GetType().GetRuntimeProperties().Where(pr => pr.GetIndexParameters().Length == 0
                                            && pr.GetValue(this) is IExcelBindableBase);
            foreach (PropertyInfo prop_inf in prop_infoes)
            {
                var prop_val = prop_inf.GetValue(this);
                if (prop_val is IExcelBindableBase exbb_prop_value)
                {
                    if (!nambered_objects.Contains(exbb_prop_value))
                    {
                        nambered_objects.Add(exbb_prop_value);
                        exbb_prop_value.SetNumberItem(possition, number);
                    }
                }
            }
            foreach (T itm in this)
            {
                if (!nambered_objects.Contains(itm))
                {
                    nambered_objects.Add(itm);
                    itm.SetNumberItem(possition, number);
                }
            }
        }
        public string GetSelfNamber()
        {
            if (this.Number != null && !Name.Contains('.')) return this.Number;
            if (this.Number == null) return "";
            string[] str_array = this.Number.Split('.');
            return str_array[str_array.Length - 1];
        }
        public virtual object Clone()
        {

            var new_collecion = (IExcelNotifyChangedCollection)Activator.CreateInstance(this.GetType());
            var prop_infoes = new_collecion.GetType().GetProperties().Where(pr => pr.GetIndexParameters().Length == 0
                                                             && pr.CanWrite
                                                             && pr.PropertyType.FullName.Contains("System.")
                                                            && pr.GetValue(this) != null
                                                            && !(pr.GetValue(this) is IList));

            foreach (PropertyInfo prop_info in prop_infoes)
            {
                var this_obj_prop_value = prop_info.GetValue(this);
                if (prop_info.CanWrite)
                    prop_info.SetValue(new_collecion, this_obj_prop_value);
            }

            new_collecion.CellAddressesMap = new ExellCellAddressMapDictationary();
            new_collecion.CellAddressesMap.Owner = new_collecion;
            Dictionary<Guid, T> map_objects = new Dictionary<Guid, T>();
            foreach (var kvp in this.CellAddressesMap.Where(k => k.Key.Contains('_')))
            {
                string prop_name = "";
                Guid guid;
                string prop_full_name = kvp.Key;
                string[] props = prop_full_name.Split('_');
                guid = Guid.Parse(props[0]);
                prop_name = props[1];
                var item = this.FirstOrDefault(elm => elm.Id == guid);
                if (item != null)
                {
                    T itm;
                    if (map_objects.ContainsKey(guid))
                    {
                        itm = map_objects[guid];
                    }
                    else
                    {
                        itm = (T)((T)item).Clone();
                        map_objects.Add(guid, itm);
                    }
                    if (!new_collecion.Contains(itm))
                    {
                        new_collecion.Add(itm);
                        itm.Owner = this.Owner;
                    }
                }
            }

            foreach (var kvp in this.CellAddressesMap.Where(k => !k.Key.Contains('_')))
                new_collecion.CellAddressesMap.Add(kvp.Key, new ExellPropAddress(kvp.Value));

            return new_collecion;
        }
    }
}
