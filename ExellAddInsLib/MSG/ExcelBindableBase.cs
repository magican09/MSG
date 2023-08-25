using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public delegate (bool, object) BeforePropertyChangeEventHandler(object sender, PropertyChangedEventArgs e, object new_val);

    public abstract class ExcelBindableBase : INotifyPropertyChanged, IExcelBindableBase
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public event BeforePropertyChangeEventHandler BeforePropertyChange;
        private Excel.Worksheet _worksheet;

        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public virtual Excel.Worksheet Worksheet
        {
            get { return _worksheet; }
            set
            {
                _worksheet = value;
                this.CellAddressesMap.SetWorksheet(_worksheet);

            }
        }

        public string Name { get; set; }

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
                catch
                {
                    throw new Exception($"Ошибка при получении свойства ExelBindableBase.NumberPrfix:{this.ToString()}:{this.Number}");
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
                        return str[str.Length - 1]; ;
                    }
                    else return null;
                }
                catch
                {
                    throw new Exception($"Ошибка при получении свойства ExelBindableBase.NumberSuffix:{this.ToString()}:{this.Number}");
                }
            }
        }
        private Guid _id = Guid.NewGuid();

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

       public  bool IsChanged { get; set; } 
        public Guid Id
        {

            get { return _id; }
        }
        public void PropertyChange(object sender, string property_name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property_name));
        }
        public void SetProperty<T>(ref T member, T new_val, [CallerMemberName] string property_name = "")
        {
            var validate_out = BeforePropertyChange?.Invoke(this, new PropertyChangedEventArgs(property_name), new_val);


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
                member =  new_val;

            PropertyChange(this, property_name);

        }
        public void SetProperty(string property_name, object new_val)
        {

        }
        public bool IsPropertyChangedHaveSubsctribers()
        {

            return PropertyChanged != null;
        }

        // public ObservableCollection<IExcelBindableBase> Owners { get; set; } = new ObservableCollection<IExcelBindableBase>();
        private IExcelBindableBase _owner;

        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public IExcelBindableBase Owner
        {
            get { return _owner; }
            set
            {

                _owner = value;
            }
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


        public ExellCellAddressMapDictationary CellAddressesMap { get; set; }
        public ExcelBindableBase()
        {
            CellAddressesMap = new ExellCellAddressMapDictationary();
            CellAddressesMap.Owner = this;
            //  CellAddressesMap.AddEvent += OnCellAdressAdd;
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
            var cell_maps = this.CellAddressesMap.Where(cm => cm.Value.Worksheet.Name == worksheet.Name);
            if (cell_maps.Any())
            {
                var left_upper_cell = cell_maps.OrderBy(c => c.Value.Row).OrderBy(c => c.Value.Column).First().Value.Cell;
                var rigth_lower_cell = cell_maps.OrderBy(c => c.Value.Row).OrderBy(c => c.Value.Column).Last().Value.Cell;

                range = worksheet.Range[left_upper_cell, rigth_lower_cell];
            }
            return range;
        }
        public virtual void SetInvalidateCellsColor(XlRgbColor color)
        {
            var invalide_cells = this.CellAddressesMap.Where(cm => cm.Value.IsValid == false);
            foreach (var kvp in invalide_cells)
            {
                kvp.Value.Cell.Interior.Color = color;
            }
        }
        public virtual void ChangeTopRow(int row)
        {
            int top_row = this.CellAddressesMap.Where(adr => !adr.Key.Contains("_")).OrderBy(kvp => kvp.Value.Row).First().Value.Row;
            int row_delta = row - top_row;
            if (top_row + row_delta <= 0) row_delta = 0;
            //  foreach(var kvp in this.CellAddressesMap.Where(k=>k.Key.Contains('_')))
            foreach (var kvp in this.CellAddressesMap.Where(adr => !adr.Key.Contains("_")))
            {
                kvp.Value.Row += row_delta;
            }
        }
        public virtual int GetRowsCount()
        {
            int top_row = this.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).First().Value.Row;
            int bottom_row = this.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).Last().Value.Row;
            return bottom_row - top_row;
        }
        public virtual int GetBottomRow()
        {
            int top_row = this.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).First().Value.Row;
            int bottom_row = this.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).Last().Value.Row;
            return bottom_row;
        }
        public virtual int GetTopRow()
        {
            int top_row = this.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).First().Value.Row;
            return top_row;
        }
        private List<IExcelBindableBase> nambered_objects = new List<IExcelBindableBase>();
        public void SetNumberItem(int possition, string number, bool first_itaration = true)
        {
            if (this.Number == null) return;

            string[] str = this.Number.Split('.');
            if(!this.Number.Contains("."))
                str = new string[] { this.Number};
            str[possition] = number;
            string out_str = "";
            foreach (string s in str)
                out_str += $"{s}.";
            out_str = out_str.TrimEnd('.');
            this.Number = out_str;
            if (first_itaration) nambered_objects.Clear();
            var prop_infoes = this.GetType().GetRuntimeProperties().Where(pr => pr.GetIndexParameters().Length == 0
                                            && pr.GetCustomAttribute(typeof(NonGettinInReflectionAttribute)) == null
                                            && pr.GetValue(this) is IExcelBindableBase);
            foreach (PropertyInfo prop_inf in prop_infoes)
            {
                var prop_val = prop_inf.GetValue(this);
                if (prop_val is WorkersComposition)
                    ;
                if (prop_val is IExcelBindableBase)
                    ;
                if (prop_val is IExcelNotifyChangedCollection)
                    ;

                if (prop_val is IExcelBindableBase exbb_prop_value)
                {

                    if (prop_val is IList exbb_list)
                    {
                        foreach (IExcelBindableBase elm in exbb_list)
                        {
                            if (!nambered_objects.Contains(elm))
                            {
                                nambered_objects.Add(elm);
                                elm.SetNumberItem(possition, number);
                            }
                        }
                    }
                    else if (!nambered_objects.Contains(exbb_prop_value))
                    {
                        nambered_objects.Add(exbb_prop_value);
                        exbb_prop_value.SetNumberItem(possition, number);

                    }
                }
            }
        }
        public virtual void UpdateExcelRepresetation()
        {
            this.UpdateExellBindableObject();
        }
        public virtual int AdjustExcelRepresentionTree(int row)
        {
            this.ChangeTopRow(row);
            return row;
        }

        public virtual void SetStyleFormats(int col)
        {


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
            CellAddressesMap.SetCellNumberFormat();
            foreach (var kvp in obj.CellAddressesMap.Where(k => !k.Key.Contains('_')))
            {
                var val = this.GetPropertyValueByPath(obj, kvp.Value.ProprertyName);
                if (val != null)
                {
                    kvp.Value.Cell.NumberFormat = kvp.Value.CellNumberFormat;
                    kvp.Value.Cell.Value = val;

                }
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
     
        internal void LoadExellBindableObjectFromField()
        {
            var obj = this;
            var prop_infoes = obj.GetType().GetProperties().Where(pr => pr.GetIndexParameters().Length == 0);
            CellAddressesMap.SetCellNumberFormat();
            foreach (var kvp in obj.CellAddressesMap.Where(k => !k.Key.Contains('_')))
            {
                var val = kvp.Value.Cell.Value;
                var owner = kvp.Value.Owner;
               var prop_info =  GetPropertyInfoByPath(owner, kvp.Value.ProprertyName);

                if (prop_info.CanWrite) prop_info.SetValue(owner, val,null);
            }
        }
        private PropertyInfo GetPropertyInfoByPath(IExcelBindableBase obj, string full_prop_name)
        {
            string[] prop_names = full_prop_name.Split('.');
            foreach (string name in prop_names)
            {
                string rest_prop_name_part = full_prop_name;
                if (full_prop_name.Contains(".")) rest_prop_name_part = full_prop_name.Replace($"{name}.", "");
                if (obj.GetType().GetProperty(name).GetCustomAttribute(typeof(NonGettinInReflectionAttribute)) != null)
                    return null;
                var prop_info = obj.GetType().GetProperty(name);
                var prop_value = prop_info.GetValue(obj);

                if (prop_value is IExcelBindableBase excel_bimdable_prop_value)
                {
                    return this.GetPropertyInfoByPath(excel_bimdable_prop_value, rest_prop_name_part);
                }
                else if (prop_value != null && prop_value.GetType().FullName.Contains("System."))
                {

                    //if (prop_value is DateTime date_val)
                    //    return date_val.ToString("d");
                    //else
                    //    return prop_value.ToString();
                    return prop_info;
                }
                else
                    return  null;
            }
            return null;
        }
        public string GetSelfNamber()
        {
            if (this.Number != null && !this.Number.Contains('.')) return this.Number;
            if (this.Number == null) return "";
            string[] str_array = this.Number.Split('.');
            return str_array[str_array.Length - 1];
        }



        public virtual object Clone()
        {
            IExcelBindableBase new_obj = (IExcelBindableBase)Activator.CreateInstance(this.GetType());
            new_obj.CellAddressesMap = new ExellCellAddressMapDictationary();
            new_obj.CellAddressesMap.Owner = new_obj;
            foreach (var kvp in this.CellAddressesMap.Where(k => !k.Key.Contains('_')))
            {
                var excell_address = new ExcelPropAddress(kvp.Value);
                excell_address.CellNumberFormat = kvp.Value.CellNumberFormat;
                new_obj.CellAddressesMap.Add(kvp.Key, excell_address);
            }

            var prop_infoes = new_obj.GetType().GetProperties().Where(pr => pr.GetIndexParameters().Length == 0
                                                                           && pr.CanWrite
                                                                           && pr.PropertyType.FullName.Contains("System.")
                                                                          && pr.GetValue(this) != null
                                                                            && !(pr.GetValue(this) is IList));
            foreach (PropertyInfo prop_info in prop_infoes)
            {
                var this_obj_prop_value = prop_info.GetValue(this);
                if (prop_info.CanWrite)
                    prop_info.SetValue(new_obj, this_obj_prop_value);
            }


            return new_obj;

        }

    }
}
