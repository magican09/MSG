using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
        public event BeforePropertyChangeEventHandler BeforePropertyChange;

        private bool _isValid = true;
        private Excel.Worksheet _worksheet;

        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public Excel.Worksheet Worksheet
        {
            get { return _worksheet; }
            set
            {
                _worksheet = value;
                this.CellAddressesMap.SetWorksheet(_worksheet);
                foreach (T itm in this)
                {
                    itm.Worksheet = _worksheet;
                    itm.CellAddressesMap.SetWorksheet(_worksheet);
                }
            }
        }

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
        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public string NumberPrefix
        {
            get
            {

                try
                {
                    if (Number != null && Number.Contains('.'))
                    {
                        var str = this.Number.Split('.');
                        List<string> out_str_arr = new List<string>();
                        for (int ii = 0; ii < str.Length - 1; ii++)
                            out_str_arr.Add(str[ii]);

                        string out_str = "";
                        foreach (string s in out_str_arr)
                            out_str += $"{s}.";
                        out_str = out_str.TrimEnd('.');
                        return out_str;
                    }
                    else return null;
                }
                catch (Exception exp)
                {
                    throw new Exception($"Ошибка при получении свойства  ExcelNotifyChangedCollection<T>..NumberPrefix:{this.ToString()}:{this.Number}.Ошибка:{exp.Message}");

                }

            }
        }
        public string NumberSuffix
        {
            get
            {
                try
                {
                    if (Number != null)
                    {
                        var str = this.Number.Split('.');
                        return str[str.Length]; ;
                    }
                    else return null;
                }
                catch (Exception exp)
                {
                    throw new Exception($"Ошибка при получении свойства ExcelNotifyChangedCollection<T>.NumberSuffix.{this.ToString()}:{this.Number}.Ошибка:{exp.Message}");
                }
            }
        }
        public Guid Id
        {
            get { return _id; }
        }

        public bool IsPropertyChangedHaveSubsctribers()
        {
            return PropertyChanged != null;
        }

        public ObservableCollection<IExcelBindableBase> Owners { get; set; } = new ObservableCollection<IExcelBindableBase>();
        public void PropertyChange(object sender, string property_name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property_name));
        }
        public void SetProperty<T>(ref T member, T new_val, [CallerMemberName] string property_name = "")
        {
            var validate_out = BeforePropertyChange?.Invoke(this, new PropertyChangedEventArgs(property_name),  new_val);
       
            if (new_val is IExcelBindableBase excell_bindable_new_val/* && !excell_bindable_new_val.Owners.Contains(this)*/)
            {
                this.RegisterNewValInCellAddresMap(excell_bindable_new_val, property_name);
            }
            if (member is IExcelBindableBase excell_bindable_member /*&& excell_bindable_member.Owners.Contains(this)*/)
            {
                this.UnregisterMemberValInCellAddresMap(excell_bindable_member, property_name);
            }

            if (validate_out != null && validate_out.Value.Item1)
                member = (T)validate_out.Value.Item2;
            else
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
                    string key_str = $"{excel_bindable_element.Id}_{kvp.Value.ProprertyName}";
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
        public virtual Excel.Range GetRange(int right_border, int low_border = 100000000, int left_border = 0, int up_border = 0)
        {
            Excel.Range _range = this.GetRange();
            Excel.Range range = null;

            Excel.Worksheet worksheet = this.Worksheet;

            foreach (Excel.Range r in _range.Columns)
            {
                if (r.Column >= left_border && r.Column <= right_border && r.Row >= up_border && r.Row <= low_border)
                {
                    if (range == null) range = r;
                    range = Worksheet.Application.Union(range, r);
                }

            }

            return range;
        }

        public virtual Excel.Range GetRange()
        {
            Excel.Range range = null;
            Excel.Worksheet worksheet = this.Worksheet;
            var cell_maps = this.CellAddressesMap.Where(cm => cm.Value.Worksheet.Name == worksheet.Name && !cm.Key.Contains('_')); //Выбираем записи элементов коллекции
            if (cell_maps.Any())
            {
                var left_upper_cell = cell_maps.OrderBy(c => c.Value.Row).OrderBy(c => c.Value.Column).First().Value.Cell;
                var rigth_lower_cell = cell_maps.OrderBy(c => c.Value.Column).OrderBy(c => c.Value.Row).Last().Value.Cell;
                range = worksheet.Range[left_upper_cell, rigth_lower_cell];
            }

            foreach (T itm in this)
            {
                Excel.Range rg = itm.GetRange();
                if (rg != null)
                {
                    if (range == null) range = rg;
                    else
                        range = Worksheet.Application.Union(range, rg);
                }
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
            if (this.CellAddressesMap.Count == 0) return;
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
            if (this.CellAddressesMap.Count == 0) return 0;
            int top_row = this.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).First().Value.Row;
            int bottom_row = this.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).Last().Value.Row;
            return bottom_row - top_row;
        }
        public int GetBottomRow()
        {
            if (this.CellAddressesMap.Count == 0) return 0;
            int top_row = this.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).First().Value.Row;
            int bottom_row = this.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).Last().Value.Row;
            return bottom_row;
        }
        public int GetTopRow()
        {
            if (this.CellAddressesMap.Count == 0) return 0;
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

        public void UpdateExcelRepresetation()
        {
            this.UpdateExellBindableObject();
            foreach (T itm in this)
                itm.UpdateExcelRepresetation();
        }
        public int AdjustExcelRepresentionTree(int row)
        {
            this.ChangeTopRow(row);

            return row;
        }
        public virtual void SetStyleFormats(int col)
        {
            foreach (T itm in this)
                itm.SetStyleFormats(col);

        }

        /// <summary>
        /// Функция обновляет документальное представление объетка (рукурсивно проходит по всем объектам 
        /// реализующим интерфейс IExcelBindableBase). 
        /// </summary>
        /// <param name="obj">Связанный с докуметом Worksheet объект рализующий IExcelBindableBase </param>
        internal void UpdateExellBindableObject()
        {
            var obj = this;
            var prop_infoes = obj.GetType().GetProperties().Where(pr => pr.GetIndexParameters().Length == 0);

            foreach (var kvp in obj.CellAddressesMap.Where(k => !k.Key.Contains('_')))
            {
                var val = this.GetPropertyValueByPath(obj, kvp.Value.ProprertyName);
                if (val != null)
                    kvp.Value.Cell.Value = val;
            }
        }
        private object GetPropertyValueByPath(IExcelBindableBase obj, string full_prop_name)
        {
            string[] prop_names = full_prop_name.Split('.');
            foreach (string name in prop_names)
            {
                string rest_prop_name_part = full_prop_name;
                if (full_prop_name.Contains(".")) rest_prop_name_part = full_prop_name.Replace($"{name}.", "");
                if (obj.GetType().GetProperty(name).GetCustomAttribute(typeof(NonGettinInReflectionAttribute)) != null)
                    return null;
                var prop_value = obj.GetType().GetProperty(name).GetValue(obj);

                if (prop_value is IExcelBindableBase excel_bimdable_prop_value)
                {
                    return this.GetPropertyValueByPath(excel_bimdable_prop_value, rest_prop_name_part);
                }
                else if (prop_value != null && prop_value.GetType().FullName.Contains("System."))
                {

                    //if (prop_value is DateTime date_val)
                    //    return date_val.ToString("d");
                    //else
                    //    return prop_value.ToString();
                    return prop_value;
                }
                else
                    return "";
            }
            return null;
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
