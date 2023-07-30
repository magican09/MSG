using Microsoft.Office.Interop.Excel;
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
            if (new_val is IExcelBindableBase excell_bindable_new_val && !excell_bindable_new_val.Owners.Contains(this))
                excell_bindable_new_val.Owners.Add(this);
            if (member is IExcelBindableBase excell_bindable_member && excell_bindable_member.Owners.Contains(this))
                excell_bindable_member.Owners.Remove(this);
            member = new_val;
            PropertyChange(this, property_name);

        }
        public ObservableCollection<IExcelBindableBase> Owners { get; set; } = new ObservableCollection<IExcelBindableBase>();
        public ExellCellAddressMapDictationary CellAddressesMap { get; set; } 
        public ExcelBindableBase()
        {
            CellAddressesMap = new ExellCellAddressMapDictationary();
            CellAddressesMap.Owner = this;
            CellAddressesMap.AddEvent += OnCellAdressAdd;
        }
        
        private void OnCellAdressAdd(IExcelBindableBase sender,ExellCellAddressMapDictationary.AddEventArgs pAddEventArgs)
        {
            if (pAddEventArgs != null)
            {
                Excel.Worksheet worksheet = pAddEventArgs.Value.Worksheet;
            }

        }

        public Excel.Range GetRange(Excel.Worksheet worksheet)
        {
            Excel.Range range =null;
            var cell_maps = this.CellAddressesMap.Where(cm => cm.Value.Worksheet.Name == worksheet.Name);
            if(cell_maps.Any())
            {
                int upper_row = cell_maps.OrderBy(c=>c.Value.Row).First().Value.Row;
                int lower_row = cell_maps.OrderBy(c => c.Value.Row).Last().Value.Row;
                int left_col = cell_maps.OrderBy(c => c.Value.Column).First().Value.Column;
                int right_col = cell_maps.OrderBy(c => c.Value.Column).Last().Value.Column;
                var left_upper_cell = cell_maps.OrderBy(c=>c.Value.Row).OrderBy(c=>c.Value.Column).First().Value.Cell;
                var rigth_lower_cell = cell_maps.OrderBy(c => c.Value.Row).OrderBy(c => c.Value.Column).Last().Value.Cell;

                range = worksheet.Range[left_upper_cell, rigth_lower_cell];

            }
            return range;
        }

        public object Clone()
        {
            var new_obj = this.MemberwiseClone();
            var prop_infoes = new_obj.GetType().GetProperties().Where(pr => pr.GetIndexParameters().Length == 0
                                                                && !pr.PropertyType.FullName.Contains("System.")
                                                                && pr.GetValue(this) != null);
            //foreach(PropertyInfo prop_info in prop_infoes)
            //  {
            //      var prop_val = prop_info.GetValue(this);
            //     if(prop_val is ICloneable clonable_prop_val)
            //      {
            //         prop_val = clonable_prop_val.Clone();
            //      }
            //     else
            //      {
            //          var constr_method = prop_val.GetType().GetConstructor(BindingFlags.Instance | BindingFlags.Public, null, new Type[0], null);
            //          prop_val = constr_method.Invoke(null);
            //      }

            //  }

            foreach (PropertyInfo prop_info in prop_infoes)
            {
                var prop_val = prop_info.GetValue(this);
                var constr_method = prop_val.GetType().GetConstructor(BindingFlags.Instance | BindingFlags.Public, null, new Type[0], null);
                prop_val = constr_method.Invoke(null);


            }

            return new_obj;

        }
    }
}
