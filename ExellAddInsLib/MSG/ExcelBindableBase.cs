using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public abstract class ExcelBindableBase : INotifyPropertyChanged, IExcelBindableBase
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private Guid _id = Guid.NewGuid();

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

            if (new_val is IExcelBindableBase excell_bindable_new_val /*&& !excell_bindable_new_val.Owners.Contains(this)*/)
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
            if (member is IExcelBindableBase excell_bindable_member /*&& excell_bindable_member.Owners.Contains(this)*/)
            {
            //    excell_bindable_member.Owners.Remove(this);
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
        // public ObservableCollection<IExcelBindableBase> Owners { get; set; } = new ObservableCollection<IExcelBindableBase>();
        [DontClone]
        [NonGettinInReflection]
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

        public void ChangeTopRow(int row)
        {
            int top_row = this.CellAddressesMap.OrderBy(kvp=>kvp.Value.Row).First().Value.Row;
            int row_delta = row-top_row;
            if (top_row+ row_delta <= 0) row_delta = 0;
            foreach(var kvp in this.CellAddressesMap)
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
        public object Clone()
        {
            if (this is UnitOfMeasurement)
                ;
            if (this is RCWork)
                ;
            if (this is WorkReportCard)
                ;
            var new_obj =  Activator.CreateInstance(this.GetType());
            var prop_infoes = new_obj.GetType().GetProperties().Where(pr => pr.GetIndexParameters().Length == 0
                                                                 && pr.CanWrite
                                                                && pr.GetValue(this) != null);
            foreach (PropertyInfo prop_info in prop_infoes)
            {
                var this_prop_value = prop_info.GetValue(this);
                var new_obj_prop_value = prop_info.GetValue(new_obj);
                if (prop_info.Name == "WorkReportCard")
                    ;
                if (prop_info.GetCustomAttribute(typeof(NonGettinInReflectionAttribute)) == null)
                {
                    if (!prop_info.PropertyType.FullName.Contains("System."))
                    {
                        if(this_prop_value is ICloneable clonable_prop_value && prop_info.GetCustomAttribute(typeof(DontCloneAttribute)) == null)
                        {
                            new_obj_prop_value = clonable_prop_value.Clone();
                        }
                        else if (prop_info.GetCustomAttribute(typeof(DontCloneAttribute)) != null)
                        {
                            new_obj_prop_value = this_prop_value;
                        }
                        else
                        {
                            var constr_method = new_obj_prop_value.GetType().GetConstructor(BindingFlags.Instance | BindingFlags.Public, null, new Type[0], null);
                            new_obj_prop_value = constr_method.Invoke(null);
                        }
                    }
                    else
                    {
                       if(prop_info.CanWrite)
                            prop_info.SetValue(new_obj, this_prop_value);
                    }
                }
                else
                    ;
            }

            return new_obj;

        }
    }
}
