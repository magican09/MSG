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

        /// <summary>
        /// Реестр зарегистрированных в системе основных объектов
        /// </summary>
        public ObservableCollection<RelateRecord> RegistedObjects = new ObservableCollection<RelateRecord>();
        /// <summary>
        /// Реест и основных и внутренных отлеживаемых обхектов и их совойств
        /// </summary>
        public ObservableCollection<RelateRecord> ObjectPropertyNameRegister = new ObservableCollection<RelateRecord>();

        /// <summary>
        /// Времення колеекция для предотвращения зацикливания в рекурсии Register(..)/
        /// </summary>
        private Dictionary<IExcelBindableBase, string> RegisterTemporalStopList = new Dictionary<IExcelBindableBase, string>();

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
                // notified_object.CellAddressesMap.Add(prop_name, address);
                ExcelSubsriptions.Add(notified_object.Subscribe(address) as ExellCellSubsciption);
            }
            //  catch (Exception ex)
            {
                //     throw  new Exception($"Ошибка при регистрации объектов в MSGExelModel. MSHExcelModel.Register(..): {ex.Message}");
            }

        }

       
        private bool IsRegistered(IExcelBindableBase obj, string prop_name)
        {
            if (this.RegistedObjects.FirstOrDefault(r => r.Entity.Id == obj.Id && r.PropertyName == prop_name) != null)
                return true;
            else
                return false;
        }

        ObservableCollection<IExcelBindableBase> unregistedObjects = new ObservableCollection<IExcelBindableBase>();
        /// <summary>
        /// Удаления регистрации объекта из системы отслеживания
        /// </summary>
        /// <param name="notified_object"></param>
        /// <param name="first_iteration"></param>
        public void Unregister(IExcelBindableBase notified_object, bool first_iteration = true)
        {
            //if (first_iteration) unregistedObjects.Clear();
            //if (unregistedObjects.Contains(notified_object)) return;
            //var all_registed_rrecords = this.RegistedObjects.Where(ro => ro.Entity.Id == notified_object.Id).ToList();
            //foreach (var r_obj in all_registed_rrecords)
            //{
            //    notified_object.PropertyChanged -= OnPropertyChanged;
            //    notified_object.BeforePropertyChange -= OnBeforPropertyChanged;
            //    this.RegistedObjects.Remove(r_obj);
            //}
            //if (notified_object is IList exbb_list)
            //    foreach (IExcelBindableBase elm in exbb_list)
            //        this.Unregister(elm);

            //var all_object_prop_names_registed_rrecords = new ObservableCollection<RelateRecord>(
            //    this.ObjectPropertyNameRegister.Where(op => op.Entity.Id == notified_object.Id).ToList());

            //foreach (var rr in all_object_prop_names_registed_rrecords)
            //    this.ObjectPropertyNameRegister.Remove(rr);

            //var prop_infoes = notified_object.GetType().GetRuntimeProperties().Where(pr => pr.GetIndexParameters().Length == 0
            //                                                         && pr.GetCustomAttribute(typeof(NonGettinInReflectionAttribute)) == null
            //                                                                             && pr.GetValue(notified_object) is IExcelBindableBase);
            //foreach (PropertyInfo property_info in prop_infoes)
            //{
            //    var property_val = property_info.GetValue(notified_object);
            //    if (property_val is IExcelBindableBase exbb_prop_val)
            //    {
            //        this.Unregister(exbb_prop_val, false);
            //    }
            //}
        }

        ObservableCollection<IExcelBindableBase> registed_objects = new ObservableCollection<IExcelBindableBase>();
        /// <summary>
        /// Регистрация всего дерева IExcelBindableBase объектов в системе отслеживания
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="firt_itaration"></param>
      

        /// <summary>
        /// Функция для получаения самого высоско вдевере регистрации объектов объекта
        /// </summary>
        /// <param name="relateRecord"></param>
        /// <returns></returns>
        private RelateRecord GetFirstParentRelateRecord(RelateRecord relateRecord)
        {
            if (relateRecord.Parent != null)
                GetFirstParentRelateRecord(relateRecord.Parent);
            else
                return relateRecord;
            return null;
        }
        /// <summary>
        /// Функция для получения всех самых нижни в дереве регистрации объектов зависимых от данного объекта.
        /// </summary>
        /// <param name="relateRecord"></param>
        /// <param name="childrenRecords"></param>
        private void GetChildrenRelateRecords(RelateRecord relateRecord, ObservableCollection<Tuple<RelateRecord, string>> childrenRecords)
        {
            string prop_name = "";
            if (relateRecord.Items.Count == 0)
                childrenRecords.Add(new Tuple<RelateRecord, string>(relateRecord, $"{relateRecord.PropertyName}"));
            foreach (RelateRecord rr in relateRecord.Items)
            {
                if (rr.Items.Count == 0)
                    childrenRecords.Add(new Tuple<RelateRecord, string>(rr, $"{relateRecord.PropertyName}.{rr.PropertyName}"));
                else
                    this.GetChildrenRelateRecords(rr, childrenRecords);
            }

        }
        /// <summary>
        /// Получение значения по пути к свойству из объекта
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="prop_path"></param>
        /// <returns></returns>
        private object GetValueFromObject(IExcelBindableBase obj, string prop_path)
        {
            string rest_prop_name_part = prop_path;

            if (prop_path.Contains("."))
                rest_prop_name_part = prop_path.Substring(prop_path.IndexOf('.') + 1, prop_path.Length - prop_path.IndexOf('.') - 1);
            string prop_name = prop_path.Replace($".{rest_prop_name_part}", "");
            if (prop_name != "")
            {
                var prop_val = obj.GetType().GetProperty(prop_name).GetValue(obj);
                if (prop_val is IExcelBindableBase ex_n_prop_val)
                    return GetValueFromObject(ex_n_prop_val, rest_prop_name_part);
                else
                {
                    var prop_non_object_val = obj.GetType().GetProperty(prop_path).GetValue(obj);
                    return prop_non_object_val;
                }
            }

            return null;
        }

        /// <summary>
        /// Обработчик собиытия изменений в зарегистрированных объетах.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            object lock_obj = new object();

            if (sender is IExcelBindableBase bindable_object)
            {
                lock (lock_obj)
                {
                    var ralated_records = this.RegistedObjects
                        .Where(rr => rr.Entity.Id == bindable_object.Id)
                        .Where(r =>
                        {
                            var prop_names_array = r.PropertyName.Split('.');
                            foreach (string name in prop_names_array)
                                if (name == e.PropertyName)
                                    return true;
                            return false;
                        });
                    foreach (RelateRecord related_rec in ralated_records) //Находим все зависимые записиыыы
                    {
                        var parent_rrecord = this.GetFirstParentRelateRecord(related_rec);
                        ObservableCollection<Tuple<RelateRecord, string>> all_children_records = new ObservableCollection<Tuple<RelateRecord, string>>(); ;
                        this.GetChildrenRelateRecords(parent_rrecord, all_children_records); //Находим все зависяцщие дочерние записи
                        var children_for_read_props = all_children_records.Where(ch => ch.Item2 == parent_rrecord.ExellPropAddress.ProprertyName); //Находим объект находящийся по зарегисрированному в реестре пути
                        foreach (Tuple<RelateRecord, string> rr_tuple in children_for_read_props)
                        {
                            var val = GetValueFromObject(parent_rrecord.Entity, rr_tuple.Item2);
                            if (parent_rrecord.ExellPropAddress.Cell.Value == null
                               || parent_rrecord.ExellPropAddress.Cell.Value.ToString() != val.ToString())
                            {
                                parent_rrecord.ExellPropAddress.Cell.Value = val;
                                parent_rrecord.ExellPropAddress.Cell.Interior.Color = XlRgbColor.rgbAquamarine;

                            }


                        }
                    }
                }
            }
        }

        public override void UpdateExcelRepresetation()
        {

        }




    }
}
