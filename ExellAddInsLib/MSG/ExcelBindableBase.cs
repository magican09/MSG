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

    public abstract class ExcelBindableBase : INotifyPropertyChanged, IObservableExcelBindableBase
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
                foreach (var observer in this._observers.Select(s => s as ExcelPropAddress))
                    observer.Worksheet = _worksheet;

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
                if (this._observers.Where(cm => (cm as ExcelPropAddress).IsValid == false).Any())
                    _isValid = false;
                return _isValid;
            }
            set { _isValid = value; }
        }

        public bool IsChanged { get; set; }
        public Guid Id
        {

            get { return _id; }
        }
        public void PropertyChange(object sender, string property_name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property_name));
            foreach (var observer in _observers)
                observer.OnNext(new PropertyChangeState(this, property_name));
        }
        public void SetProperty<T>(ref T member, T new_val, [CallerMemberName] string property_name = "")
        {
            member = new_val;
            PropertyChange(this, property_name);

        }
        public void SetProperty(string property_name, object new_val)
        {

        }
        public bool IsPropertyChangedHaveSubsctribers()
        {

            return PropertyChanged != null;
        }

        // public ObservableCollection<IObservableExcelBindableBase> Owners { get; set; } = new ObservableCollection<IObservableExcelBindableBase>();
        private IObservableExcelBindableBase _owner;

        [NonGettinInReflection]
        [NonRegisterInUpCellAddresMap]
        public IObservableExcelBindableBase Owner
        {
            get { return _owner; }
            set
            {

                _owner = value;
            }
        }

        //private void RegisterNewValInCellAddresMap(IObservableExcelBindableBase excell_bindable_new_val, string property_name)
        //{
        //    //    excell_bindable_new_val.Owners.Add(this);
        //    var non_reg_in_upper_attribute = this.GetType().GetProperty(property_name).GetCustomAttribute(typeof(NonRegisterInUpCellAddresMapAttribute));
        //    if (non_reg_in_upper_attribute == null)
        //    {
        //        //  excell_bindable_new_val.Owners.Add(this);
        //        foreach (var kvp in excell_bindable_new_val._observers)
        //        {
        //            string key_str = $"{excell_bindable_new_val.Id.ToString()}_{observer.ProprertyName}";
        //            if (!this._observers.ContainsKey(key_str))
        //                this._observers.Add(key_str, observer);
        //        }
        //        excell_bindable_new_val._observers.AddEvent += OnCellAdressAdd;
        //    }
        //}
        //private void UnregisterMemberValInCellAddresMap(IObservableExcelBindableBase excell_bindable_member, string property_name)
        //{
        //    //     excell_bindable_member.Owners.Remove(this);
        //    foreach (var kvp in excell_bindable_member._observers)
        //    {
        //        string key_str = $"{excell_bindable_member.Id.ToString()}_{observer.ProprertyName}";
        //        if (this._observers.ContainsKey(key_str))
        //            this._observers.Remove(key_str);
        //    }

        //    excell_bindable_member._observers.AddEvent -= OnCellAdressAdd;
        //}


        //private void OnCellAdressAdd(IObservableExcelBindableBase sender, ExellCellAddressMapDictationary.AddEventArgs pAddEventArgs)
        //{
        //    if (pAddEventArgs != null)
        //    {
        //        Excel.Worksheet worksheet = pAddEventArgs.Value.Worksheet;
        //        string key_str = $"{sender.Id.ToString()}_{pAddEventArgs.Value.ProprertyName}";
        //        this._observers.Add(key_str, pAddEventArgs.Value);
        //    }

        //}


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
            var cell_maps = this._observers.Where(cm => (cm as ExcelPropAddress).Worksheet.Name == worksheet.Name).Select(s => s as ExcelPropAddress);
            if (cell_maps.Any())
            {
                var left_upper_cell = cell_maps.OrderBy(c => c.Row).OrderBy(c => c.Column).First().Cell;
                var rigth_lower_cell = cell_maps.OrderBy(c => c.Row).OrderBy(c => c.Column).Last().Cell;

                range = worksheet.Range[left_upper_cell, rigth_lower_cell];
            }
            return range;
        }
        public virtual void SetInvalidateCellsColor(XlRgbColor color)
        {
            var invalide_cells = this._observers.Select(s => s as ExcelPropAddress).Where(cm => (cm as ExcelPropAddress).IsValid == false);
            foreach (var observer in invalide_cells)
            {
                observer.Cell.Interior.Color = color;
            }
        }
        public virtual void ChangeTopRow(int row)
        {
            int top_row = this._observers.Select(s => s as ExcelPropAddress).Where(observer => !observer.ProprertyName.Contains("_")).OrderBy(observer => observer.Row).First().Row;
            int row_delta = row - top_row;
            if (top_row + row_delta <= 0) row_delta = 0;
            //  foreach(var kvp in this._observers.Where(k=>k.Key.Contains('_')))
            foreach (var observer in this._observers.Select(s => s as ExcelPropAddress).Where(observer => !observer.ProprertyName.Contains("_")))
            {
                observer.Row += row_delta;
            }
        }
        public virtual int GetRowsCount()
        {
            int top_row = this._observers.Select(s => s as ExcelPropAddress).OrderBy(observer => observer.Row).First().Row;
            int bottom_row = this._observers.Select(s => s as ExcelPropAddress).OrderBy(observer => observer.Row).Last().Row;
            return bottom_row - top_row;
        }
        public virtual int GetBottomRow()
        {
            int top_row = this._observers.Select(s => s as ExcelPropAddress).OrderBy(observer => observer.Row).First().Row;
            int bottom_row = this._observers.Select(s => s as ExcelPropAddress).OrderBy(observer => observer.Row).Last().Row;
            return bottom_row;
        }
        public virtual int GetTopRow()
        {
            int top_row = this._observers.Select(s => s as ExcelPropAddress).OrderBy(kvp => kvp.Row).First().Row;
            return top_row;
        }
        public virtual int GetLeftColumn()
        {
            int left_col = this._observers.Select(s => s as ExcelPropAddress).OrderBy(kvp => kvp.Column).First().Column;
            return left_col;
        }

        private List<IObservableExcelBindableBase> nambered_objects = new List<IObservableExcelBindableBase>();
        public void SetNumberItem(int possition, string number, bool first_itaration = true)
        {
            if (this.Number == null) return;

            string[] str = this.Number.Split('.');
            if (!this.Number.Contains("."))
                str = new string[] { this.Number };
            str[possition] = number;
            string out_str = "";
            foreach (string s in str)
                out_str += $"{s}.";
            out_str = out_str.TrimEnd('.');
            this.Number = out_str;
            if (first_itaration) nambered_objects.Clear();
            var prop_infoes = this.GetType().GetRuntimeProperties().Where(pr => pr.GetIndexParameters().Length == 0
                                            && pr.GetCustomAttribute(typeof(NonGettinInReflectionAttribute)) == null
                                            && pr.GetValue(this) is IObservableExcelBindableBase);
            foreach (PropertyInfo prop_inf in prop_infoes)
            {
                var prop_val = prop_inf.GetValue(this);
                if (prop_val is IObservableExcelBindableBase exbb_prop_value)
                {

                    if (prop_val is IList exbb_list)
                    {
                        foreach (IObservableExcelBindableBase elm in exbb_list)
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
        /// реализующим интерфейс IObservableExcelBindableBase). 
        /// </summary>
        /// <param name="obj">Связанный с докуметом Worksheet объект рализующий IObservableExcelBindableBase </param>
        internal void UpdateExellBindableObject()
        {

            var prop_infoes = this.GetType().GetProperties().Where(pr => pr.GetIndexParameters().Length == 0);

            foreach (var observer in _observers.Select(s => s as ExcelPropAddress))
                observer.SetCellNumberFormat();

            foreach (var observer in this._observers.Select(s => s as ExcelPropAddress).Where(obs => !obs.ProprertyName.Contains('_')))
            {
                var val = this.GetPropertyValueByPath(this, observer.ProprertyName);
                if (val != null)
                {
                    observer.Cell.NumberFormat = observer.CellNumberFormat;
                    observer.Cell.Value = val;

                }
            }
        }
        private object GetPropertyValueByPath(IObservableExcelBindableBase obj, string full_prop_name)
        {
            string[] prop_names = full_prop_name.Split('.');
            foreach (string name in prop_names)
            {
                string rest_prop_name_part = full_prop_name;
                if (full_prop_name.Contains(".")) rest_prop_name_part = full_prop_name.Replace($"{name}.", "");
                if (obj.GetType().GetProperty(name).GetCustomAttribute(typeof(NonGettinInReflectionAttribute)) != null)
                    return null;
                var prop_value = obj.GetType().GetProperty(name).GetValue(obj);

                if (prop_value is IObservableExcelBindableBase excel_bimdable_prop_value)
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
            foreach (var observer in _observers.Select(s => s as ExcelPropAddress))
                observer.SetCellNumberFormat();

            foreach (var observer in obj._observers.Select(s => s as ExcelPropAddress).Where(obs => !obs.ProprertyName.Contains('_')))
            {
                var val = observer.Cell.Value;
                var owner = observer.Owner;
                var prop_info = GetPropertyInfoByPath(owner as IObservableExcelBindableBase, observer.ProprertyName);

                if (prop_info.CanWrite) prop_info.SetValue(owner, val, null);
            }
        }
        private PropertyInfo GetPropertyInfoByPath(IObservableExcelBindableBase obj, string full_prop_name)
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

                if (prop_value is IObservableExcelBindableBase excel_bimdable_prop_value)
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
                    return null;
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
            IObservableExcelBindableBase new_obj = (IObservableExcelBindableBase)Activator.CreateInstance(this.GetType());
            foreach (var observer in this._observers.Select(s => s as ExcelPropAddress).Where(obs => !obs.ProprertyName.Contains('_')))
            {
                var excell_address = new ExcelPropAddress(observer);
                excell_address.CellNumberFormat = observer.CellNumberFormat;
                //  new_obj._observers.Add(kvp.Key, excell_address);
                new_obj.Subscribe(excell_address);
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
            new_obj.Worksheet = this.Worksheet;

            return new_obj;

        }
        private List<IObserver<PropertyChangeState>> _observers = new List<IObserver<PropertyChangeState>>();
        public List<IDisposable> Subscribers { get; private set; } = new List<IDisposable>();
        public IDisposable Subscribe(IObserver<PropertyChangeState> observer)
        {
            if (!_observers.Contains(observer as ExcelPropAddress))
            {
                _observers.Add(observer as ExcelPropAddress);
                IDisposable subscriber = new ExellCellSubsciption(observer as ExcelPropAddress, this, _observers);
                Subscribers.Add(subscriber);
                return subscriber;
            }

            return null;
        }
        public ExcelPropAddress this[string i]
        {
            get
            {
                return this._observers.Select(s => s as ExcelPropAddress).FirstOrDefault(obs => obs.ProprertyName == i); ;
            }

        }
        public void SetPropertyValidStatus(string prop_name, bool isValid)
        {
            foreach (var observer in this._observers)
                observer.OnNext(new PropertyChangeState(this, prop_name, isValid));
        }
        public Excel.Range GetCell(string prop_name)
        {
            return this._observers.Select(s => s as ExcelPropAddress).FirstOrDefault(obs => obs.ProprertyName == prop_name).Cell;
        }
        public ExcelPropAddress GetPropAddress(string prop_name)
        {
            return this._observers.Select(s => s as ExcelPropAddress).FirstOrDefault(obs => obs.ProprertyName == prop_name);
        }
    }
}
