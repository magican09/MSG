using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.Common;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class ExellModelBase : ExcelBindableBase
    {
        public const int HASH_FUNCTION_COL = 1;
        public const int HASH_FUNCTION_ROW = 7;
        public const int MAX_HASH_FUNCTION_ROW = 10000;
        public const int MAX_HASH_FUNCTION_COL = 37;

        public Dictionary<Tuple<int, int>, ExcelPropAddress> AllHashDictationary = new Dictionary<Tuple<int, int>, ExcelPropAddress>();
        public List<int> RowsHashValues = new List<int>();
        public List<int> ColumnsHashValues = new List<int>();
        public bool IsHasEnabled = false;

        public ObservableCollection<Excel.Worksheet> AllWorksheets = new ObservableCollection<Excel.Worksheet>();

    
        public List<ExellCellSubsciption> ExcelSubsriptions = new List<ExellCellSubsciption>();
        /// <summary>
        /// Функция для регистрации объекта реализующего интрефейс INotifyPropertyChanged 
        /// для обработки событий изменения полей объета и соотвествующего изменения связанной с 
        /// с этим полем ячейки в документе Worksheet
        /// </summary>
        /// <param name="work"></param>
        public void Register(IObservable<PropertyChangeState> notified_object, string prop_name, int row, int column,
            Excel.Worksheet worksheet, Func<object, bool> validate_value_call_back = null,
               Func<object, object> coerce_value_call_back = null, RelateRecord register = null)
        {

            //  try
            {
                var prop_names_chain = prop_name.Split(new char[] { '.' });
                Type prop_type = notified_object.GetType().GetProperty(prop_names_chain[prop_names_chain.Length-1]).PropertyType;

                var address = new ExcelPropAddress(row, column, worksheet, prop_type, prop_name, validate_value_call_back, coerce_value_call_back);
                address.Owner = notified_object;
                ExcelSubsriptions.Add(notified_object.Subscribe(address) as ExellCellSubsciption);
            }
            //  catch (Exception ex)
            {
                //     throw  new Exception($"Ошибка при регистрации объектов в MSGExelModel. MSHExcelModel.Register(..): {ex.Message}");
            }

        }

       
        private bool IsRegistered(IExcelBindableBase obj, string prop_name)
        {
            if (this.ExcelSubsriptions.FirstOrDefault(r => (r.Observable as IObservableExcelBindableBase).Id == obj.Id && (r.Observer as ExcelPropAddress).ProprertyName == prop_name) != null)
                return true;
            else
                return false;
        }

        /// <summary>
        /// Удаления регистрации объекта из системы отслеживания
        /// </summary>
        /// <param name="notified_object"></param>
        /// <param name="first_iteration"></param>
        public void Unregister(IObservableExcelBindableBase notified_object, bool first_iteration = true)
        {
            var subscriptions = this.ExcelSubsriptions.Where(subs => (subs.Observable as IObservableExcelBindableBase).Id == notified_object.Id);
           foreach(var subs in subscriptions)
                 subs.Dispose();
        }


        public override void UpdateExcelRepresetation()
        {

        }




    }
}
