﻿using ExellAddInsLib.MSG.Section;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.Common;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using static System.Collections.Specialized.BitVector32;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class MSGExellModel : ExcelBindableBase
    {
        public const int COMMON_PARAMETRS_VALUE_COL = 3;

        public const int CONTRACT_CODE_ROW = 2;
        public const int CONSTRUCTION_OBJECT_CODE_ROW = 3;
        public const int CONSTRUCTION_SUBOBJECT_CODE_ROW = 4;


        public const int WORKS_START_DATE_ROW = 1;
        public const int WORKS_TART_DATE_COL = 3;
        public const int WORKS_END_DATE_ROW = 2;
        public const int WORKS_END_DATE_COL = 3;

        public const int FIRST_ROW_INDEX = 7;

        public const int WSEC_NUMBER_COL = 2;
        public const int WSEC_NAME_COL = WSEC_NUMBER_COL + 1;

        public const int MSG_NUMBER_COL = 4;
        public const int MSG_NAME_COL = MSG_NUMBER_COL + 1;
        public const int MSG_MEASURE_COL = MSG_NUMBER_COL + 2;
        public const int MSG_QUANTITY_COL = MSG_NUMBER_COL + 3;
        public const int MSG_QUANTITY_FACT_COL = MSG_NUMBER_COL + 4;
        public const int MSG_LABOURNESS_COL = MSG_NUMBER_COL + 5;
        public const int MSG_START_DATE_COL = MSG_NUMBER_COL + 6;
        public const int MSG_END_DATE_COL = MSG_NUMBER_COL + 7;

        public const int MSG_NEEDS_OF_WORKERS_NUMBER_COL = MSG_END_DATE_COL + 1;
        public const int MSG_NEEDS_OF_WORKERS_NAME_COL = MSG_NEEDS_OF_WORKERS_NUMBER_COL + 1;
        public const int MSG_NEEDS_OF_WORKERS_QUANTITY_COL = MSG_NEEDS_OF_WORKERS_NAME_COL + 1;


        public const int VOVR_NUMBER_COL = 15;
        public const int VOVR_NAME_COL = VOVR_NUMBER_COL + 1;
        public const int VOVR_MEASURE_COL = VOVR_NUMBER_COL + 2;
        public const int VOVR_QUANTITY_COL = VOVR_NUMBER_COL + 3;
        public const int VOVR_QUANTITY_FACT_COL = VOVR_NUMBER_COL + 4;
        public const int VOVR_LABOURNESS_COL = VOVR_NUMBER_COL + 5;


        public const int KS_NUMBER_COL = 21;
        public const int KS_CODE_COL = KS_NUMBER_COL + 1;
        public const int KS_NAME_COL = KS_NUMBER_COL + 2;
        public const int KS_MEASURE_COL = KS_NUMBER_COL + 3;
        public const int KS_QUANTITY_COL = KS_NUMBER_COL + 4;
        public const int KS_QUANTITY_FACT_COL = KS_NUMBER_COL + 5;
        public const int KS_LABOURNESS_COL = KS_NUMBER_COL + 6;

        public const int RC_NUMBER_COL = KS_LABOURNESS_COL + 1;
        public const int RC_CODE_COL = RC_NUMBER_COL + 1;
        public const int RC_NAME_COL = RC_NUMBER_COL + 2;
        public const int RC_MEASURE_COL = RC_NUMBER_COL + 3;
        public const int RC_QUANTITY_COL = RC_NUMBER_COL + 4;
        public const int RC_QUANTITY_FACT_COL = RC_NUMBER_COL + 5;
        public const int RC_LABOURNESS_COEFFICIENT_COL = RC_NUMBER_COL + 6;
        public const int RC_LABOURNESS_COL = RC_NUMBER_COL + 7;


        public const int WRC_DATE_ROW = 6;

        public const int WRC_NUMBER_COL = RC_LABOURNESS_COL + 1;
        public const int WRC_PC_QUANTITY_COL = WRC_NUMBER_COL + 1;
        public const int WRC_DATE_COL = WRC_PC_QUANTITY_COL + 1;

        public const int W_CONSUMPTIONS_FIRST_ROW_INDEX = 4;
        public const int W_CONSUMPTIONS_NUMBER_COL = 1;
        public const int W_CONSUMPTIONS_NAME_COL = 2;
        public const int W_CONSUMPTIONS_DATE_RAW = 3;
        public const int W_CONSUMPTIONS_FIRST_DATE_COL = 3;

        public const int _SECTIONS_GAP = 3;


        private int null_str_count = 0;
        /// <summary>
        /// Дата начала ведомости 
        /// </summary>
        public DateTime WorksStartDate { get; set; }
        /// <summary>
        /// Дата окончания работ в данной ведомости в соотвествии с планируемыми в  сроками отраженнным в части МСГ ведомости.
        /// (в части WorkSchedules работ MSGWork)
        /// </summary>
        public DateTime WorksEndDate
        {
            get
            {
                DateTime end_date = DateTime.MinValue;
                var last_ended_work = this.MSGWorks.OrderBy(w => w.WorkSchedules.EndDate).LastOrDefault();
                if (last_ended_work != null)
                    end_date = last_ended_work.WorkSchedules.EndDate;
                return end_date;
            }

        }



        /// <summary>
        /// Коллекция с разделами работ
        /// </summary>
        private ExcelNotifyChangedCollection<WorksSection> _worksSections;

        public ExcelNotifyChangedCollection<WorksSection> WorksSections
        {
            get { return _worksSections; }
            set { SetProperty(ref _worksSections, value); }
        }
        /// <summary>
        /// Коллекция с работами типа МСГ модели
        /// </summary>

        private ExcelNotifyChangedCollection<MSGWork> _mSGWorks;

        public ExcelNotifyChangedCollection<MSGWork> MSGWorks
        {
            get { return _mSGWorks; }
            set { SetProperty(ref _mSGWorks, value); }
        }
        /// <summary>
        /// Коллекция с работами типа ВОВР модели
        /// </summary>
        private ExcelNotifyChangedCollection<VOVRWork> _vOVRWorks;

        public ExcelNotifyChangedCollection<VOVRWork> VOVRWorks
        {
            get { return _vOVRWorks; }
            set { SetProperty(ref _vOVRWorks, value); }
        }
        /// <summary>
        /// Коллекция с работами типа КС-2 модели
        /// </summary>

        private ExcelNotifyChangedCollection<KSWork> _kSWorks;

        public ExcelNotifyChangedCollection<KSWork> KSWorks
        {
            get { return _kSWorks; }
            set { SetProperty(ref _kSWorks, value); }
        }
        /// <summary>
        /// Коллекция с работами типа ждя учета модели
        /// </summary>

        private ExcelNotifyChangedCollection<RCWork> _rCWorks;

        public ExcelNotifyChangedCollection<RCWork> RCWorks
        {
            get { return _rCWorks; }
            set { SetProperty(ref _rCWorks, value); }
        }
        /// <summary>
        /// Коллекция  табелей выполненных работ
        /// </summary>
        private ExcelNotifyChangedCollection<WorkReportCard> _workReportCards;

        public ExcelNotifyChangedCollection<WorkReportCard> WorkReportCards
        {
            get { return _workReportCards; }
            set { SetProperty(ref _workReportCards, value); }
        }
        /// <summary>
        /// Коллекция с единицами измерения модели
        /// </summary>
        private ExcelNotifyChangedCollection<UnitOfMeasurement> _unitOfMeasurements;

        public ExcelNotifyChangedCollection<UnitOfMeasurement> UnitOfMeasurements
        {
            get { return _unitOfMeasurements; }
            set { SetProperty(ref _unitOfMeasurements, value); }
        }
        //public WorkersCompositionReportCard WorkersCompositionReportCard = new MSG.WorkersCompositionReportCard();

        private WorkersComposition _WorkersComposition;

        public WorkersComposition WorkersComposition
        {
            get { return _WorkersComposition; }
            set { SetProperty(ref _WorkersComposition, value); }
        }
        private WorkerConsumptions _workerConsumptions;

        public WorkerConsumptions WorkerConsumptions
        {
            get { return _workerConsumptions; }
            set { SetProperty(ref _workerConsumptions, value); }
        }

        /// <summary>
        ///Шифр объекта или договора
        /// </summary>
        public string ContractCode { get; set; }
        /// <summary>
        ///Наименоваение объекта/договора
        /// </summary
        public string ContructionObjectCode { get; set; }
        /// <summary>
        ///Наименование подобъекта
        /// </summary
        public string ConstructionSubObjectCode { get; set; }
        /// <summary>
        /// Ссылка на родительскую модель
        /// </summary>
        public MSGExellModel Owner { get; set; }
        /// <summary>
        /// Коллекия дочерних моделей
        /// </summary>
        public ObservableCollection<MSGExellModel> Children { get; set; } = new ObservableCollection<MSGExellModel>();
        /// <summary>
        /// Прикрепленный к модели лист ведомости  Worksheet
        /// </summary>
        private Excel.Worksheet _registerSheet;

        public Excel.Worksheet RegisterSheet
        {
            get
            {
                return _registerSheet;
            }
            set
            {
                if (!AllWorksheets.Contains(value))
                {
                    if (AllWorksheets.Contains(_registerSheet))
                        AllWorksheets.Remove(_registerSheet);
                    AllWorksheets.Add(value);
                }
                _registerSheet = value;

            }
        }
        /// <summary>
        /// Прикрепленный к модели лист  Людских ресурсов Worksheet
        /// </summary>
        ///    
        private Excel.Worksheet _workerConsumptionsSheet;


        public Excel.Worksheet WorkerConsumptionsSheet
        {
            get
            {
                return _workerConsumptionsSheet;
            }
            set
            {
                if (!AllWorksheets.Contains(value))
                {
                    if (AllWorksheets.Contains(_workerConsumptionsSheet))
                        AllWorksheets.Remove(_workerConsumptionsSheet);
                    AllWorksheets.Add(value);
                }
                _workerConsumptionsSheet = value;

            }
        }



        /// <summary>
        /// Прикрепленный к модели лист общих данных  Worksheet
        /// </summary>
        private Excel.Worksheet _commonSheet;

        public Excel.Worksheet CommonSheet
        {
            get
            {
                return _commonSheet;
            }
            set
            {
                if (!AllWorksheets.Contains(value))
                {
                    if (AllWorksheets.Contains(_commonSheet))
                        AllWorksheets.Remove(_commonSheet);
                    AllWorksheets.Add(value);
                }
                _commonSheet = value;
            }
        }

        public ObservableCollection<Excel.Worksheet> AllWorksheets = new ObservableCollection<Excel.Worksheet>();
        /// <summary>
        /// Отвественных за работы отраженных в работах данной модели
        /// </summary>
        public Employer Employer { get; set; }
        public MSGExellModel()
        {
            WorksSections = new ExcelNotifyChangedCollection<WorksSection>();
            MSGWorks = new ExcelNotifyChangedCollection<MSGWork>();
            VOVRWorks = new ExcelNotifyChangedCollection<VOVRWork>();
            KSWorks = new ExcelNotifyChangedCollection<KSWork>();
            RCWorks = new ExcelNotifyChangedCollection<RCWork>();
            WorkReportCards = new ExcelNotifyChangedCollection<WorkReportCard>();
            WorkersComposition = new WorkersComposition();
            WorkerConsumptions = new WorkerConsumptions();
            UnitOfMeasurements = new ExcelNotifyChangedCollection<UnitOfMeasurement>();


        }
        // private ObservableCollection<IExcelBindableBase> RegistedObjects = new ObservableCollection<IExcelBindableBase>();
        // private EventedDictationary<RalateRecord, Tuple<string, ExellPropAddress>> RegistedObjects = new EventedDictationary<RalateRecord, Tuple<string, ExellPropAddress>>();
        // private EventedDictationary<RalateRecord, ExellPropAddress> RegistedObjects = new EventedDictationary<RalateRecord, ExellPropAddress>();
        public ObservableCollection<RelateRecord> RegistedObjects = new ObservableCollection<RelateRecord>();
        // private EventedDictationary<IExcelBindableBase, string> ObjectPropertyNameRegister = new EventedDictationary<IExcelBindableBase, string>();
        public ObservableCollection<RelateRecord> ObjectPropertyNameRegister = new ObservableCollection<RelateRecord>();
        public Dictionary<IExcelBindableBase, string> RegisterTemporalStopList = new Dictionary<IExcelBindableBase, string>();
        /// <summary>
        /// Функция для регистрации объекта реализующего интрефейс INotifyPropertyChanged 
        /// для обработки событий изменения полей объета и соотвествующего изменения связанной с 
        /// с этим полем ячейки в документе Worksheet
        /// </summary>
        /// <param name="work"></param>
        public void Register(IExcelBindableBase notified_object, string prop_name, int row, int column, Excel.Worksheet worksheet, RelateRecord register = null)
        {

            if (notified_object is MSGWork)
                ;
            try
            {
                var prop_names = prop_name.Split(new char[] { '.' });
                RelateRecord local_register = new RelateRecord(notified_object);
                if (register == null)
                {
                    register = local_register;

                    if (notified_object.CellAddressesMap.ContainsKey(prop_name))
                        local_register.ExellPropAddress = notified_object.CellAddressesMap[prop_name];
                    else
                    {
                        local_register.ExellPropAddress = new ExellPropAddress(row, column, worksheet, prop_name);
                        notified_object.CellAddressesMap.Add(prop_name, local_register.ExellPropAddress);
                    }
                    //  if (!notified_object.CellAddressesMap.ContainsKey(prop_name))
                    //      notified_object.CellAddressesMap.Add(prop_name, local_register.ExellPropAddress);

                    this.RegistedObjects.Add(local_register);
                    RegisterTemporalStopList.Clear();
                    RegisterTemporalStopList.Add(local_register.Entity, prop_names[0]);
                    local_register.PropertyName = prop_names[0];
                    this.ObjectPropertyNameRegister.Add(local_register);
                }
                else
                    register.Items.Add(local_register);

                foreach (string name in prop_names)
                {
                    string rest_prop_name_part = prop_name;
                    if (prop_name.Contains(".")) rest_prop_name_part = prop_name.Replace($"{name}.", "");

                    if (!RegisterTemporalStopList.ContainsKey(local_register.Entity))
                    {
                        RegisterTemporalStopList.Add(local_register.Entity, name);
                        local_register.PropertyName = name;
                        ObjectPropertyNameRegister.Add(local_register);
                    }

                    var prop_value = notified_object.GetType().GetProperty(name).GetValue(notified_object);
                    if (prop_value is IExcelBindableBase excel_bimdable_prop_value)
                    {
                        this.Register(excel_bimdable_prop_value, rest_prop_name_part, row, column, worksheet, local_register);
                    }

                    notified_object.PropertyChanged += OnPropertyChange;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при геристрации объектов в MSGExelModel: {ex.Message}");
            }

        }
        public void Unregister(IExcelBindableBase notified_object)
        {
            // var all_registed_rrecords = this.RegistedObjects.Where(ro => ro.Entity.Id == notified_object.Id);
            var all_object_prop_names_registed_rrecords = new ObservableCollection<RelateRecord>(
                this.ObjectPropertyNameRegister.Where(op => op.Entity.Id == notified_object.Id).ToList());
            foreach (var rr in all_object_prop_names_registed_rrecords)
                this.ObjectPropertyNameRegister.Remove(rr);
        }

        ObservableCollection<IExcelBindableBase> registed_objects = new ObservableCollection<IExcelBindableBase>();
        public void RegisterObjectInObjectPropertyNameRegister(IExcelBindableBase obj, bool firt_itaration = true)
        {
            if (firt_itaration == true) registed_objects.Clear();

            if (!registed_objects.Contains(obj))
            {
                var cell_eddr_maps = obj.CellAddressesMap.Where(kvp => !kvp.Key.Contains('_'));
                foreach (var kvp in cell_eddr_maps)
                {
                    string prop_name = kvp.Key;
                    string kvp_worksheet_name = kvp.Value.Worksheet.Name;
                    string sheet_root_name = kvp_worksheet_name.Substring(0, kvp_worksheet_name.IndexOf('_'));
                    var work_sheet = this.AllWorksheets.FirstOrDefault(wh => wh.Name.Contains(sheet_root_name));
                    this.Register(obj, prop_name, kvp.Value.Row, kvp.Value.Column, work_sheet);

                }
                registed_objects.Add(obj);
                var prop_infoes = obj.GetType().GetProperties().Where(pr => pr.GetIndexParameters().Length == 0
                                                                            && pr.GetValue(obj) is IExcelBindableBase);
                foreach (PropertyInfo prop_info in prop_infoes)
                {
                    var prop_val = prop_info.GetValue(obj) as IExcelBindableBase;
                    if (prop_val is IList list_prop_val)
                    {
                        foreach (var elm in list_prop_val)
                            if (elm is IExcelBindableBase exb_elm)
                                this.RegisterObjectInObjectPropertyNameRegister(exb_elm, false);
                    }
                    else
                        this.RegisterObjectInObjectPropertyNameRegister(prop_val, false);

                }

            }
            //if (firt_itaration == true) registed_objects.Clear();
            //if(!registed_objects.Contains(obj))
            //{
            //    var cell_eddr_maps = obj.CellAddressesMap.Where(kvp => !kvp.Key.Contains('_'));
            //    foreach (var kvp in cell_eddr_maps)
            //    {
            //        RelateRecord local_register = new RelateRecord(obj);
            //        local_register.ExellPropAddress = kvp.Value;
            //        local_register.PropertyName = kvp.Key.Split('.')[0];
            //        obj.PropertyChanged += OnPropertyChange;
            //        this.ObjectPropertyNameRegister.Add(local_register);

            //    }
            //    registed_objects.Add(obj);
            //    var prop_infoes = obj.GetType().GetProperties().Where(pr => pr.GetIndexParameters().Length == 0
            //                                                                && pr.GetValue(obj) is IExcelBindableBase);
            //    foreach (PropertyInfo prop_info in prop_infoes)
            //    {
            //        var prop_val = prop_info.GetValue(obj) as IExcelBindableBase;
            //        this.RegisterObjectInObjectPropertyNameRegister(prop_val, false);
            //    }
            //}
        }

        private RelateRecord GetFirstParentRelateRecord(RelateRecord relateRecord)
        {
            if (relateRecord.Parent != null)
                GetFirstParentRelateRecord(relateRecord.Parent);
            else
                return relateRecord;
            return null;
        }
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
        private void OnPropertyChange(object sender, PropertyChangedEventArgs e)
        {
            if (sender is IExcelBindableBase bindable_object)
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
                        parent_rrecord.ExellPropAddress.Cell.Value = val;
                        parent_rrecord.ExellPropAddress.Cell.Interior.Color = XlRgbColor.rgbAquamarine;
                    }
                }
            }
        }


        /// <summary>
        /// Функция из части РАЗДЕЛЫ  листа Worksheet создает и помещает в модель  разделы работ
        /// </summary>
        public void LoadWorksSections()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;

            WorksStartDate = DateTime.Parse(registerSheet.Cells[WORKS_START_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());
            int rowIndex = FIRST_ROW_INDEX;
            null_str_count = 0;
            foreach (var section in this.WorksSections)
                this.Unregister(section);
            this.WorksSections.Clear();
            while (null_str_count < 100)
            {
                if (registerSheet.Cells[rowIndex, WSEC_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    WorksSection w_section = new WorksSection();

                    w_section.Number = registerSheet.Cells[rowIndex, WSEC_NUMBER_COL].Value.ToString();
                    w_section.Name = registerSheet.Cells[rowIndex, WSEC_NAME_COL].Value;
                    this.Register(w_section, "Number", rowIndex, WSEC_NUMBER_COL, registerSheet);
                    this.Register(w_section, "Name", rowIndex, WSEC_NAME_COL, registerSheet);
                    if (!this.WorksSections.Contains(w_section))
                        this.WorksSections.Add(w_section);
                }

                rowIndex++;
            }
        }
        /// <summary>
        /// Функция из части МСГ листа Worksheet создает и помещает в модель работы типа MSGWork 
        /// </summary>
        public void LoadMSGWorks()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;

            int rowIndex = FIRST_ROW_INDEX;
            null_str_count = 0;
            foreach (var work in this.MSGWorks)
                this.Unregister(work);
            this.MSGWorks.Clear();

            while (null_str_count < 100)
            {
                if (registerSheet.Cells[rowIndex, MSG_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    MSGWork msg_work = new MSGWork();
                    msg_work.Number = registerSheet.Cells[rowIndex, MSG_NUMBER_COL].Value.ToString();
                    msg_work.Name = registerSheet.Cells[rowIndex, MSG_NAME_COL].Value;

                    if (registerSheet.Cells[rowIndex, MSG_MEASURE_COL].Value != null)
                    {
                        string un_name = registerSheet.Cells[rowIndex, MSG_MEASURE_COL].Value.ToString();
                        UnitOfMeasurement unitOfMeasurement = UnitOfMeasurements.FirstOrDefault(um => um.Name == un_name);
                        if (unitOfMeasurement != null)
                        {
                            msg_work.UnitOfMeasurement = unitOfMeasurement;
                            this.Register(msg_work, "UnitOfMeasurement.Name", rowIndex, MSG_MEASURE_COL, registerSheet);
                        }
                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_MEASURE_COL], registerSheet.Cells[rowIndex, MSG_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, MSG_QUANTITY_COL].Value != null)
                        msg_work.ProjectQuantity = Decimal.Parse(registerSheet.Cells[rowIndex, MSG_QUANTITY_COL].Value.ToString());
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_QUANTITY_COL], registerSheet.Cells[rowIndex, MSG_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value != null)
                    {
                        var fdf = registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value.ToString();
                        decimal res;
                        Decimal.TryParse(registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value.ToString(), out res);
                        msg_work.Laboriousness = res;//Decimal.Parse(registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value.ToString());
                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL], registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    this.Register(msg_work, "Number", rowIndex, MSG_NUMBER_COL, this.RegisterSheet);
                    this.Register(msg_work, "Name", rowIndex, MSG_NAME_COL, this.RegisterSheet);
                    this.Register(msg_work, "ProjectQuantity", rowIndex, MSG_QUANTITY_COL, this.RegisterSheet);
                    this.Register(msg_work, "Quantity", rowIndex, MSG_QUANTITY_FACT_COL, this.RegisterSheet);
                    this.Register(msg_work, "Laboriousness", rowIndex, MSG_LABOURNESS_COL, this.RegisterSheet);

                    DateTime start_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_START_DATE_COL].Value.ToString());
                    DateTime end_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_END_DATE_COL].Value.ToString());

                    foreach (var sh in msg_work.WorkSchedules)
                        this.Unregister(sh);
                    msg_work.WorkSchedules.Clear();
                    WorkScheduleChunk work_sh_chunk = new WorkScheduleChunk(start_time, end_time);

                    this.Register(work_sh_chunk, "StartTime", rowIndex, MSG_START_DATE_COL, this.RegisterSheet);
                    this.Register(work_sh_chunk, "EndTime", rowIndex, MSG_END_DATE_COL, this.RegisterSheet);
                    msg_work.WorkSchedules.Add(work_sh_chunk);

                    while (registerSheet.Cells[rowIndex + 1, MSG_NUMBER_COL].Value == null
                                 && registerSheet.Cells[rowIndex + 1, MSG_START_DATE_COL].Value != null)
                    {
                        rowIndex++;
                        start_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_START_DATE_COL].Value.ToString());
                        end_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_END_DATE_COL].Value.ToString());
                        WorkScheduleChunk extra_work_sh_chunk = new WorkScheduleChunk(start_time, end_time);
                        this.Register(extra_work_sh_chunk, "StartTime", rowIndex, MSG_START_DATE_COL, this.RegisterSheet);
                        this.Register(extra_work_sh_chunk, "EndTime", rowIndex, MSG_END_DATE_COL, this.RegisterSheet);
                        msg_work.WorkSchedules.Add(extra_work_sh_chunk);
                    }

                    if (!this.MSGWorks.Contains(msg_work))
                        this.MSGWorks.Add(msg_work);
                    WorksSection w_section = this.WorksSections.Where(ws => ws.Number.StartsWith(msg_work.Number.Remove(msg_work.Number.LastIndexOf(".")))).FirstOrDefault();
                    if (w_section != null)
                        w_section.MSGWorks.Add(msg_work);

                }
                rowIndex++;
            }
        }
        /// <summary>
        /// Функция из части МСГ листа Worksheet 
        /// </summary>
        public void LoadMSGWorkerCompositions()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;

            int rowIndex = FIRST_ROW_INDEX;
            foreach (var work_coposition in this.WorkersComposition)
                this.Unregister(work_coposition);
            this.WorkersComposition.Clear();
            null_str_count = 0;


            while (null_str_count < 100)
            {
                if (registerSheet.Cells[rowIndex, MSG_NEEDS_OF_WORKERS_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    NeedsOfWorker msg_needs_of_workers = new NeedsOfWorker();

                    msg_needs_of_workers.Number = registerSheet.Cells[rowIndex, MSG_NEEDS_OF_WORKERS_NUMBER_COL].Value.ToString();

                    msg_needs_of_workers.Name = registerSheet.Cells[rowIndex, MSG_NEEDS_OF_WORKERS_NAME_COL].Value;

                    if (registerSheet.Cells[rowIndex, MSG_NEEDS_OF_WORKERS_QUANTITY_COL].Value != null)
                    {
                        msg_needs_of_workers.Quantity = Int32.Parse(registerSheet.Cells[rowIndex, MSG_NEEDS_OF_WORKERS_QUANTITY_COL].Value.ToString());
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_NEEDS_OF_WORKERS_QUANTITY_COL], registerSheet.Cells[rowIndex, MSG_NEEDS_OF_WORKERS_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;

                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_NEEDS_OF_WORKERS_QUANTITY_COL], registerSheet.Cells[rowIndex, MSG_NEEDS_OF_WORKERS_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbRed;
                    this.Register(msg_needs_of_workers, "Number", rowIndex, MSG_NEEDS_OF_WORKERS_NUMBER_COL, this.RegisterSheet);
                    this.Register(msg_needs_of_workers, "Name", rowIndex, MSG_NEEDS_OF_WORKERS_NAME_COL, this.RegisterSheet);
                    this.Register(msg_needs_of_workers, "Quantity", rowIndex, MSG_NEEDS_OF_WORKERS_QUANTITY_COL, this.RegisterSheet);

                    MSGWork msg_work = this.MSGWorks.Where(w => w.Number.StartsWith(msg_needs_of_workers.Number.Remove(msg_needs_of_workers.Number.LastIndexOf(".")))).FirstOrDefault();
                    if (msg_work != null)
                    {
                        msg_work.WorkersComposition.Add(msg_needs_of_workers);
                        //    msg_needs_of_workers.Owner = msg_work;
                        foreach (WorkScheduleChunk chunk in msg_work.WorkSchedules)
                        {
                            for (DateTime date = chunk.StartTime; date <= chunk.EndTime; date = date.AddDays(1))
                            {
                                NeedsOfWorkersDay needsOfWorkersDay = new NeedsOfWorkersDay();
                                needsOfWorkersDay.Date = date;
                                needsOfWorkersDay.Quantity = msg_needs_of_workers.Quantity;
                                msg_needs_of_workers.NeedsOfWorkersReportCard.Add(needsOfWorkersDay);
                            }
                        }
                    }

                    NeedsOfWorker global_needs_of_worker = this.WorkersComposition.FirstOrDefault(nw => nw.Name == msg_needs_of_workers.Name);
                    if (global_needs_of_worker == null)
                    {
                        global_needs_of_worker = new NeedsOfWorker();
                        global_needs_of_worker.Number = msg_needs_of_workers.Number;
                        global_needs_of_worker.Name = msg_needs_of_workers.Name;
                        foreach (NeedsOfWorkersDay needsOfWorkersDay in msg_needs_of_workers.NeedsOfWorkersReportCard)
                            global_needs_of_worker.NeedsOfWorkersReportCard.Add(needsOfWorkersDay);
                        this.WorkersComposition.Add(global_needs_of_worker);
                    }
                    else
                    {
                        foreach (NeedsOfWorkersDay needsOfWorkersDay in msg_needs_of_workers.NeedsOfWorkersReportCard)
                        {
                            var nw_day = global_needs_of_worker.NeedsOfWorkersReportCard.FirstOrDefault(nwd => nwd.Date == needsOfWorkersDay.Date);
                            if (nw_day != null)
                            {
                                nw_day.Quantity += needsOfWorkersDay.Quantity;
                            }
                            else
                            {
                                NeedsOfWorkersDay new_nw_day = new NeedsOfWorkersDay(needsOfWorkersDay);
                                global_needs_of_worker.NeedsOfWorkersReportCard.Add(new_nw_day);
                            }
                        }

                    }
                }
                rowIndex++;
            }
        }
        /// <summary>
        /// Функция из части  ВОВР листа Worksheet создает и помещает в модель работы типа VOVRWork 
        /// </summary>
        public void LoadVOVRWorks()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;
            foreach (var work in this.VOVRWorks)
                this.Unregister(work);
            this.VOVRWorks.Clear();
            null_str_count = 0;

            while (null_str_count < 100)
            {
                if (registerSheet.Cells[rowIndex, VOVR_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    VOVRWork vovr_work = new VOVRWork();

                    vovr_work.Number = registerSheet.Cells[rowIndex, VOVR_NUMBER_COL].Value.ToString();
                    vovr_work.Name = registerSheet.Cells[rowIndex, VOVR_NAME_COL].Value.ToString();
                    if (registerSheet.Cells[rowIndex, VOVR_MEASURE_COL].Value != null)
                    {
                        vovr_work.UnitOfMeasurement = new UnitOfMeasurement(registerSheet.Cells[rowIndex, VOVR_MEASURE_COL].Value);
                        this.Register(vovr_work, "UnitOfMeasurement.Name", rowIndex, VOVR_MEASURE_COL, this.RegisterSheet);
                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_MEASURE_COL], registerSheet.Cells[rowIndex, VOVR_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL].Value != null)
                    {
                        vovr_work.ProjectQuantity = Decimal.Parse(registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL].Value.ToString());

                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL], registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL].Value != null)
                    {
                        vovr_work.Laboriousness = Decimal.Parse(registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL].Value.ToString());
                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL], registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    this.Register(vovr_work, "Number", rowIndex, VOVR_NUMBER_COL, this.RegisterSheet);
                    this.Register(vovr_work, "Name", rowIndex, VOVR_NAME_COL, this.RegisterSheet);
                    this.Register(vovr_work, "ProjectQuantity", rowIndex, VOVR_QUANTITY_COL, this.RegisterSheet);
                    this.Register(vovr_work, "Quantity", rowIndex, VOVR_QUANTITY_FACT_COL, this.RegisterSheet);
                    this.Register(vovr_work, "Laboriousness", rowIndex, VOVR_LABOURNESS_COL, this.RegisterSheet);

                    if (!this.VOVRWorks.Contains(vovr_work))
                        this.VOVRWorks.Add(vovr_work);

                    MSGWork msg_work = this.MSGWorks.Where(w => w.Number.StartsWith(vovr_work.Number.Remove(vovr_work.Number.LastIndexOf(".")))).FirstOrDefault();
                    if (msg_work != null)
                        msg_work.VOVRWorks.Add(vovr_work);




                }

                rowIndex++;
            }
        }

        /// <summary>
        /// Функция из части КС-2 листа Worksheet создает и помещает в модель работы типа KSWork 
        /// </summary>
        public void LoadKSWorks()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;
            foreach (var work in this.KSWorks)
                this.Unregister(work);
            this.KSWorks.Clear();
            null_str_count = 0;
            while (null_str_count < 100)
            {
                if (registerSheet.Cells[rowIndex, KS_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    KSWork ks_work = new KSWork();

                    ks_work.Number = registerSheet.Cells[rowIndex, KS_NUMBER_COL].Value.ToString();
                    ks_work.Code = registerSheet.Cells[rowIndex, KS_CODE_COL].Value.ToString();
                    ks_work.Name = registerSheet.Cells[rowIndex, KS_NAME_COL].Value.ToString();

                    if (registerSheet.Cells[rowIndex, KS_MEASURE_COL].Value != null)
                    {
                        ks_work.UnitOfMeasurement = new UnitOfMeasurement(registerSheet.Cells[rowIndex, KS_MEASURE_COL].Value);
                        this.Register(ks_work, "UnitOfMeasurement.Name", rowIndex, KS_MEASURE_COL, this.RegisterSheet);
                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, KS_MEASURE_COL], registerSheet.Cells[rowIndex, KS_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, KS_QUANTITY_COL].Value != null)
                        ks_work.ProjectQuantity = Decimal.Parse(registerSheet.Cells[rowIndex, KS_QUANTITY_COL].Value.ToString());
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, KS_QUANTITY_COL], registerSheet.Cells[rowIndex, KS_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbRed;


                    if (registerSheet.Cells[rowIndex, KS_LABOURNESS_COL].Value != null)
                        ks_work.Laboriousness = Decimal.Parse(registerSheet.Cells[rowIndex, KS_LABOURNESS_COL].Value.ToString());
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, KS_LABOURNESS_COL], registerSheet.Cells[rowIndex, KS_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbRed;



                    this.Register(ks_work, "Number", rowIndex, KS_NUMBER_COL, this.RegisterSheet);
                    this.Register(ks_work, "Code", rowIndex, KS_CODE_COL, this.RegisterSheet);
                    this.Register(ks_work, "Name", rowIndex, KS_NAME_COL, this.RegisterSheet);
                    this.Register(ks_work, "ProjectQuantity", rowIndex, KS_QUANTITY_COL, this.RegisterSheet);
                    this.Register(ks_work, "Quantity", rowIndex, KS_QUANTITY_FACT_COL, this.RegisterSheet);
                    this.Register(ks_work, "Laboriousness", rowIndex, KS_LABOURNESS_COL, this.RegisterSheet);


                    if (!this.KSWorks.Contains(ks_work))
                        this.KSWorks.Add(ks_work);
                    VOVRWork vovr_work = VOVRWorks.Where(w => w.Number.StartsWith(ks_work.Number.Remove(ks_work.Number.LastIndexOf(".")))).FirstOrDefault();
                    if (vovr_work != null)
                        vovr_work.KSWorks.Add(ks_work);
                }
                rowIndex++;
            }
        }
        /// <summary>
        /// Функция из части КС-2 листа Worksheet создает и помещает в модель работы типа KSWork 
        /// </summary>
        public void LoadRCWorks()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            int rowIndex = FIRST_ROW_INDEX;
            foreach (var work in this.RCWorks)
                this.Unregister(work);
            this.RCWorks.Clear();
            null_str_count = 0;
            while (null_str_count < 100)
            {
                if (registerSheet.Cells[rowIndex, RC_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    RCWork rc_work = new RCWork();

                    rc_work.Number = registerSheet.Cells[rowIndex, RC_NUMBER_COL].Value.ToString();
                    rc_work.Code = registerSheet.Cells[rowIndex, RC_CODE_COL].Value.ToString();
                    rc_work.Name = registerSheet.Cells[rowIndex, RC_NAME_COL].Value.ToString();

                    if (registerSheet.Cells[rowIndex, RC_MEASURE_COL].Value != null)
                    {
                        rc_work.UnitOfMeasurement = new UnitOfMeasurement(registerSheet.Cells[rowIndex, RC_MEASURE_COL].Value);
                        this.Register(rc_work, "UnitOfMeasurement.Name", rowIndex, RC_MEASURE_COL, this.RegisterSheet);
                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, RC_MEASURE_COL], registerSheet.Cells[rowIndex, RC_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, RC_QUANTITY_COL].Value != null)
                        rc_work.ProjectQuantity = Decimal.Parse(registerSheet.Cells[rowIndex, RC_QUANTITY_COL].Value.ToString());
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, RC_QUANTITY_COL], registerSheet.Cells[rowIndex, RC_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, RC_LABOURNESS_COEFFICIENT_COL].Value != null)
                        rc_work.LabournessCoefficient = Decimal.Parse(registerSheet.Cells[rowIndex, RC_LABOURNESS_COEFFICIENT_COL].Value.ToString());
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, RC_LABOURNESS_COEFFICIENT_COL], registerSheet.Cells[rowIndex, RC_LABOURNESS_COEFFICIENT_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, RC_LABOURNESS_COL].Value != null)
                        rc_work.Laboriousness = Decimal.Parse(registerSheet.Cells[rowIndex, RC_LABOURNESS_COL].Value.ToString());
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, RC_LABOURNESS_COL], registerSheet.Cells[rowIndex, RC_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    this.Register(rc_work, "Number", rowIndex, RC_NUMBER_COL, this.RegisterSheet);
                    this.Register(rc_work, "Code", rowIndex, RC_CODE_COL, this.RegisterSheet);
                    this.Register(rc_work, "Name", rowIndex, RC_NAME_COL, this.RegisterSheet);
                    this.Register(rc_work, "ProjectQuantity", rowIndex, RC_QUANTITY_COL, this.RegisterSheet);
                    this.Register(rc_work, "Quantity", rowIndex, RC_QUANTITY_FACT_COL, this.RegisterSheet);
                    this.Register(rc_work, "LabournessCoefficient", rowIndex, RC_LABOURNESS_COEFFICIENT_COL, this.RegisterSheet);
                    this.Register(rc_work, "Laboriousness", rowIndex, RC_LABOURNESS_COL, this.RegisterSheet);

                    if (!this.RCWorks.Contains(rc_work))
                        this.RCWorks.Add(rc_work);
                    KSWork ks_work = this.KSWorks.Where(w => w.Number.StartsWith(rc_work.Number.Remove(rc_work.Number.LastIndexOf(".")))).FirstOrDefault();
                    if (ks_work != null)
                        ks_work.RCWorks.Add(rc_work);
                }
                rowIndex++;
            }
        }
        /// <summary>
        /// Функция из календарной части (левой части) листа Worksheet создает и помещает в соответсвующие  
        /// работы типа KSWork табели выполненных работ ReportCard с объектами типа WorkDay с даной и количеством 
        /// выполенной работы
        /// </summary>
        public void LoadWorksReportCards()
        {
            Excel.Worksheet registerSheet = this.RegisterSheet;
            foreach (var rc in this.WorkReportCards)
                this.Unregister(rc);
            this.WorkReportCards.Clear();
            int rowIndex = FIRST_ROW_INDEX;

            null_str_count = 0;
            if (this.Owner != null)
                while (null_str_count < 100)
                {
                    if (registerSheet.Cells[rowIndex, WRC_NUMBER_COL].Value == null)
                        null_str_count++;
                    else
                    {
                        null_str_count = 0;
                        string rc_number = registerSheet.Cells[rowIndex, WRC_NUMBER_COL].Value.ToString();
                        WorkReportCard report_card = new WorkReportCard();

                        DateTime end_date = DateTime.Parse(registerSheet.Cells[WORKS_END_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());
                        report_card.Number = registerSheet.Cells[rowIndex, WRC_NUMBER_COL].Value.ToString();
                        this.Register(report_card, "Number", rowIndex, WRC_NUMBER_COL, this.RegisterSheet);

                        if (registerSheet.Cells[rowIndex, WRC_PC_QUANTITY_COL].Value != null)
                            report_card.PreviousComplatedQuantity = Decimal.Parse(registerSheet.Cells[rowIndex, WRC_PC_QUANTITY_COL].Value.ToString());

                        this.Register(report_card, "PreviousComplatedQuantity", rowIndex, WRC_PC_QUANTITY_COL, this.RegisterSheet);

                        int date_index = 0;
                        while (registerSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value != null &&
                            DateTime.Parse(registerSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value.ToString()) < end_date)
                        {
                            DateTime current_date = DateTime.Parse(registerSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value.ToString());
                            decimal quantity = 0;
                            if (registerSheet.Cells[rowIndex, WRC_DATE_COL + date_index].Value != null)
                                quantity = Decimal.Parse(registerSheet.Cells[rowIndex, WRC_DATE_COL + date_index].Value.ToString());
                            if (quantity != 0)
                            {
                                WorkDay workDay = new WorkDay();
                                workDay.Date = current_date;
                                workDay.Quantity = quantity;
                                this.Register(workDay, "Quantity", rowIndex, WRC_DATE_COL + date_index, this.RegisterSheet);
                                report_card.Add(workDay);
                            }
                            date_index++;
                        }
                        RCWork rc_work = this.RCWorks.FirstOrDefault(w => w.Number == rc_number);
                        if (rc_work != null)
                            rc_work.ReportCard = report_card;

                        if (!this.WorkReportCards.Contains(report_card))
                            this.WorkReportCards.Add(report_card);



                    }
                    rowIndex++;
                }

        }
        public void LoadWorkerConsumptions()
        {
            Excel.Worksheet consumtionsSheet = this.WorkerConsumptionsSheet;
            int rowIndex = W_CONSUMPTIONS_FIRST_ROW_INDEX;
            foreach (var wc in this.WorkerConsumptions)
                this.Unregister(wc);
            this.WorkerConsumptions.Clear();
            null_str_count = 0;

            while (null_str_count < 100)
            {
                if (WorkerConsumptionsSheet.Cells[rowIndex, W_CONSUMPTIONS_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    WorkerConsumption worker_consumption = new WorkerConsumption();

                    worker_consumption.Number = consumtionsSheet.Cells[rowIndex, W_CONSUMPTIONS_NUMBER_COL].Value.ToString();

                    worker_consumption.Name = consumtionsSheet.Cells[rowIndex, W_CONSUMPTIONS_NAME_COL].Value;
                    worker_consumption.WorkersConsumptionReportCard.Clear();

                    int date_index = 0;
                    if (this.Owner != null)
                        while (consumtionsSheet.Cells[W_CONSUMPTIONS_DATE_RAW, W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value != null)
                        {
                            DateTime current_date = DateTime.Parse(consumtionsSheet.Cells[W_CONSUMPTIONS_DATE_RAW, W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString());
                            decimal quantity = 0;
                            if (consumtionsSheet.Cells[rowIndex, W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value != null)
                                quantity = Decimal.Parse(consumtionsSheet.Cells[rowIndex, W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString());

                            if (quantity != 0)
                            {
                                WorkerConsumptionDay w_consumption_Day = new WorkerConsumptionDay();
                                w_consumption_Day.Date = current_date;
                                // workDay.CellAddressesMap.Add("Date", new ExellPropAddress(WRC_DATE_ROW, WRC_DATE_COL + date_index));
                                w_consumption_Day.Quantity = quantity;
                                this.Register(w_consumption_Day, "Quantity", rowIndex, W_CONSUMPTIONS_FIRST_DATE_COL + date_index, consumtionsSheet);
                                worker_consumption.WorkersConsumptionReportCard.Add(w_consumption_Day);
                            }

                            date_index++;
                        }
                    if (!this.WorkerConsumptions.Contains(worker_consumption))
                        this.WorkerConsumptions.Add(worker_consumption);
                    this.Register(worker_consumption, "Number", rowIndex, W_CONSUMPTIONS_NUMBER_COL, consumtionsSheet);
                    this.Register(worker_consumption, "Name", rowIndex, W_CONSUMPTIONS_NAME_COL, consumtionsSheet);

                }
                rowIndex++;
            }

        }

        /// <summary>
        /// Функция перезагружает все объекты из всех Worksheet в соотвествующие модели. 
        /// </summary>
        public void ReloadAllSheetsInModel()
        {
            foreach (MSGExellModel model in Children)
                model.ReloadSheetModel();
            this.ReloadSheetModel();
        }
        /// <summary>
        /// Функция перезагружает все объектные модели с соответсвующих листов
        /// </summary>
        public void ReloadSheetModel()
        {
            this.UpdateCellAddressMapsWorkSheets();
            this.ContractCode = this.CommonSheet.Cells[CONTRACT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL].Value.ToString();
            this.ContructionObjectCode = this.CommonSheet.Cells[CONSTRUCTION_OBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL].Value.ToString();
            this.ConstructionSubObjectCode = this.CommonSheet.Cells[CONSTRUCTION_SUBOBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL].Value.ToString();

            //this.CellAddressesMap.Add("ContractCode", new ExellPropAddress<int, int, Worksheet>(CONTRACT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL, this.CommonSheet));
            //this.CellAddressesMap.Add("ContructionObjectCode", new ExellPropAddress<int, int, Worksheet>(CONSTRUCTION_OBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL, this.CommonSheet));
            //this.CellAddressesMap.Add("ConstructionSubObjectCode", new ExellPropAddress<int, int, Worksheet>(CONSTRUCTION_SUBOBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL, this.CommonSheet));
            this.WorksStartDate = DateTime.Parse(this.RegisterSheet.Cells[WORKS_START_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());
            // this.WorksEndDate = DateTime.Parse(this.RegisterSheet.Cells[WORKS_END_DATE_COL, WORKS_END_DATE_COL].Value.ToString());

            //this.CellAddressesMap.Add("WorksStartDate", new ExellPropAddress<int, int, Worksheet>(WORKS_START_DATE_ROW, WORKS_END_DATE_COL, this.RegisterSheet));
            if (this.Owner == null)
            {
                this.LoadWorksSections();
                this.LoadMSGWorks();
                this.LoadMSGWorkerCompositions();
                this.LoadVOVRWorks();
                this.LoadKSWorks();
                this.LoadRCWorks();
                this.LoadWorksReportCards();
                this.LoadWorkerConsumptions();
            }
            else
            {
                this.LoadWorksReportCards();
                this.LoadWorkerConsumptions();
            }
        }

        public void SetStyleFormats()
        {
            int selectin_col = 33;
            this.SetBordersBoldLine(this.WorksSections.GetRange(this.RegisterSheet));
            int first_row = 0;
            int last_row = 0;
            foreach (WorksSection section in this.WorksSections)
            {
                var section_range = section.GetRange(this.RegisterSheet);
                section_range.Interior.ColorIndex = selectin_col;
                first_row = section_range.Row;
                this.SetBordersBoldLine(section.MSGWorks.GetRange(this.RegisterSheet));
                int msg_work_col = selectin_col + 1;
                foreach (MSGWork msg_work in section.MSGWorks)
                {
                    var msg_work_range = msg_work.GetRange(this.RegisterSheet);
                    msg_work_range.Interior.ColorIndex = msg_work_col;
                    this.SetBordersBoldLine(msg_work.VOVRWorks.GetRange(this.RegisterSheet));
                    this.SetBordersBoldLine(msg_work.WorkSchedules.GetRange(this.RegisterSheet));
                    this.SetBordersBoldLine(msg_work.WorkersComposition.GetRange(this.RegisterSheet));
                    foreach (NeedsOfWorker need_of_worker in msg_work.WorkersComposition)
                    {
                        var need_of_worker_range = need_of_worker.GetRange(this.RegisterSheet);
                        need_of_worker_range.Interior.ColorIndex = msg_work_col;
                    }
                    int vovr_work_col = msg_work_col + 1;
                    foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                    {
                        var vovr_work_range = vovr_work.GetRange(this.RegisterSheet);
                        vovr_work_range.Interior.ColorIndex = vovr_work_col;
                        this.SetBordersBoldLine(vovr_work.KSWorks.GetRange(this.RegisterSheet));
                        int ks_work_col = vovr_work_col;
                        foreach (KSWork ks_work in vovr_work.KSWorks)
                        {
                            var ks_work_range = ks_work.GetRange(this.RegisterSheet);
                            ks_work_range.Interior.ColorIndex = ks_work_col;

                            this.SetBordersBoldLine(ks_work.RCWorks.GetRange(this.RegisterSheet, RC_LABOURNESS_COL));

                            int rc_works_top_row = ks_work.RCWorks.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).First().Value.Row;
                            int rc_works_bottom_row = ks_work.RCWorks.CellAddressesMap.OrderBy(kvp => kvp.Value.Row).Last().Value.Row;

                            foreach (RCWork rc_work in ks_work.RCWorks)
                            {
                                var rc_work_range = rc_work.GetRange(this.RegisterSheet, RC_LABOURNESS_COL);
                                rc_work_range.Interior.ColorIndex = ks_work_col;
                                if (rc_work.ReportCard != null && rc_work.ReportCard.CellAddressesMap.Count > 0)
                                {
                                    var cr_range = rc_work.ReportCard.GetRange(this.RegisterSheet);
                                    //this.SetBordersBoldLine(cr_range, XlLineStyle.xlLineStyleNone,XlLineStyle.xlDouble,
                                    //    XlLineStyle.xlContinuous, XlLineStyle.xlContinuous);
                                    if (cr_range != null) cr_range.Interior.ColorIndex = ks_work_col;
                                    Excel.Range last_cell = this.RegisterSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);

                                    var days_row_range = this.RegisterSheet.Range[
                                           this.RegisterSheet.Cells[rc_work.ReportCard.CellAddressesMap["Number"].Cell.Row, WRC_NUMBER_COL],
                                           this.RegisterSheet.Cells[rc_work.ReportCard.CellAddressesMap["Number"].Cell.Row, last_cell.Column]];
                                    days_row_range.Interior.ColorIndex = ks_work_col;

                                    this.SetBordersBoldLine(days_row_range,
                                        XlLineStyle.xlLineStyleNone, XlLineStyle.xlContinuous,
                                        XlLineStyle.xlContinuous, XlLineStyle.xlContinuous);


                                }
                                if (last_row < rc_work_range.Row) last_row = rc_work_range.Row;
                            }

                            var report_cards_range = this.RegisterSheet.Range[this.RegisterSheet.Cells[rc_works_top_row, WRC_NUMBER_COL],
                                                                            this.RegisterSheet.Cells[rc_works_bottom_row, WRC_PC_QUANTITY_COL]];
                            this.SetBordersBoldLine(report_cards_range, XlLineStyle.xlLineStyleNone);
                            this.SetBordersBoldLine(ks_work_range, XlLineStyle.xlLineStyleNone);
                        }
                        vovr_work_col++;
                    }
                }

                Excel.Range range = this.RegisterSheet.Range[this.RegisterSheet.Rows[first_row + 1],
                     this.RegisterSheet.Rows[last_row + _SECTIONS_GAP]];
                range.Group();
            }

            this.SetBordersBoldLine(this.WorkerConsumptions.GetRange(this.WorkerConsumptionsSheet));
            int w_consumption_col = 33;
            foreach (WorkerConsumption consumption in this.WorkerConsumptions)
            {
                consumption.GetRange(this.WorkerConsumptionsSheet).Interior.ColorIndex = w_consumption_col++;

            }
            //this.SetFormulas();
        }
        public void SetFormulas()
        {
            int days_number = (this.WorksEndDate - this.WorksStartDate).Days;
            foreach (WorksSection section in this.WorksSections)
            {
                foreach (MSGWork msg_work in section.MSGWorks)
                {
                    foreach (NeedsOfWorker need_of_worker in msg_work.WorkersComposition)
                    {
                    }
                    string msg_works_labourness_sum_formula = "";
                    foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                    {
                        string vovr_works_labourness_sum_formula = "";
                        foreach (KSWork ks_work in vovr_work.KSWorks)
                        {
                            string rc_works_labourness_sum_formula = "";
                            foreach (RCWork rc_work in ks_work.RCWorks)
                            {
                                var first_cell = this.RegisterSheet.Cells[rc_work.CellAddressesMap["Number"].Row, WRC_PC_QUANTITY_COL];
                                var lastt_cell = this.RegisterSheet.Cells[rc_work.CellAddressesMap["Number"].Row, WRC_PC_QUANTITY_COL + 1 + (this.WorksEndDate - this.WorksStartDate).Days];

                                rc_work.CellAddressesMap["Quantity"].Cell.FormulaLocal =
                                    $"=СУММ({Func.RangeAddress(first_cell)}:{Func.RangeAddress(lastt_cell)})";
                                //   rc_work.CellAddressesMap["Quantity"].Cell.NumberFormat = "#,##0";

                                rc_works_labourness_sum_formula +=
                                    $"{Func.RangeAddress(rc_work.CellAddressesMap["Quantity"].Cell)}*{Func.RangeAddress(rc_work.CellAddressesMap["Laboriousness"].Cell)}+";
                                if (this.Owner == null)
                                {

                                    int day_iterator = 0;
                                    while (day_iterator <= days_number)
                                    {
                                        string w_day_quantity_furmula = "";

                                        foreach (MSGExellModel model in this.Children)
                                        {
                                            var child_rc = model.WorkReportCards.FirstOrDefault(rc => rc.Number == rc_work.Number);
                                            if (child_rc != null)
                                            {
                                                var w_day_range =
                                                    model.RegisterSheet.Cells[child_rc.CellAddressesMap["Number"].Row, WRC_DATE_COL + day_iterator];
                                                w_day_quantity_furmula += $"{model.RegisterSheet.Name}!{Func.RangeAddress(w_day_range)}+";
                                            }
                                        }
                                        w_day_quantity_furmula = w_day_quantity_furmula.TrimEnd('+');

                                        this.RegisterSheet.Cells[rc_work.CellAddressesMap["Number"].Row, WRC_DATE_COL + day_iterator].Formula =
                                                                                                     $"={w_day_quantity_furmula}";
                                        day_iterator++;
                                    }
                                    string rc_previous_quantity_furmula = "";
                                    foreach (MSGExellModel model in this.Children)
                                    {

                                        var child_rc = model.WorkReportCards.FirstOrDefault(rc => rc.Number == rc_work.Number);
                                        if (child_rc != null)
                                        {
                                            var child_rc_pr_quantiti_range = child_rc.CellAddressesMap["PreviousComplatedQuantity"].Cell;
                                            rc_previous_quantity_furmula += $"{model.RegisterSheet.Name}!{Func.RangeAddress(child_rc_pr_quantiti_range)}+";
                                        }
                                    }
                                    if (rc_work.ReportCard == null)
                                    {
                                        rc_work.ReportCard = new WorkReportCard();
                                        this.Register(rc_work.ReportCard, "Number", rc_work.CellAddressesMap["Number"].Row, WRC_NUMBER_COL, this.RegisterSheet);
                                        this.Register(rc_work.ReportCard, "PreviousComplatedQuantity", rc_work.CellAddressesMap["Number"].Row, WRC_PC_QUANTITY_COL, this.RegisterSheet);
                                        rc_work.ReportCard.Number = rc_work.Number;
                                    }

                                    rc_previous_quantity_furmula = rc_previous_quantity_furmula.TrimEnd('+');
                                        var rc_work_previous_range = rc_work.ReportCard.CellAddressesMap["PreviousComplatedQuantity"].Cell.Formula =
                                                                                                     $"={rc_previous_quantity_furmula}";

                                }

                            }
                            rc_works_labourness_sum_formula = rc_works_labourness_sum_formula.TrimEnd('+');
                            string ks_quantity_formula = $"=({rc_works_labourness_sum_formula})/{Func.RangeAddress(ks_work.CellAddressesMap["Laboriousness"].Cell)}";
                            ks_work.CellAddressesMap["Quantity"].Cell.Formula = ks_quantity_formula;

                            vovr_works_labourness_sum_formula +=
                                $"{Func.RangeAddress(ks_work.CellAddressesMap["Quantity"].Cell)}*{Func.RangeAddress(ks_work.CellAddressesMap["Laboriousness"].Cell)}+";
                        }
                        vovr_works_labourness_sum_formula = vovr_works_labourness_sum_formula.TrimEnd('+');
                        string vovr_quantity_formula = $"=({vovr_works_labourness_sum_formula})/{Func.RangeAddress(vovr_work.CellAddressesMap["Laboriousness"].Cell)}";
                        vovr_work.CellAddressesMap["Quantity"].Cell.Formula = vovr_quantity_formula;

                        msg_works_labourness_sum_formula +=
                                             $"{Func.RangeAddress(vovr_work.CellAddressesMap["Quantity"].Cell)}*{Func.RangeAddress(vovr_work.CellAddressesMap["Laboriousness"].Cell)}+";

                    }
                    msg_works_labourness_sum_formula = msg_works_labourness_sum_formula.TrimEnd('+');
                    string msg_quantity_formula = $"=({msg_works_labourness_sum_formula})/{Func.RangeAddress(msg_work.CellAddressesMap["Laboriousness"].Cell)}";
                    msg_work.CellAddressesMap["Quantity"].Cell.Formula = msg_quantity_formula;

                }

            }
        }
        private void SetBordersBoldLine(Excel.Range range)
        {
            if (range == null) return;
            //range.Borders.LineStyle = Excel.XlLineStyle.xlDot;
            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlDouble;
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlDouble;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlDouble;
        }
        private void SetBordersBoldLine(Excel.Range range, bool right = true, bool left = true, bool top = true, bool bottom = true)
        {
            if (range == null) return;

            if (left) range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlDouble;
            else range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            if (top) range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlDouble;
            else range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            if (bottom) range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            else range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            if (right) range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlDouble;
            else range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        }
        private void SetBordersBoldLine(Excel.Range range,
            Excel.XlLineStyle right = Excel.XlLineStyle.xlDouble,
            Excel.XlLineStyle left = Excel.XlLineStyle.xlDouble,
            Excel.XlLineStyle top = Excel.XlLineStyle.xlDouble,
            Excel.XlLineStyle bottom = Excel.XlLineStyle.xlDouble)
        {
            if (range == null) return;

            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = left;
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = top;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = bottom;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = right;
        }
        private void ClearStyleFormatsByObjects()
        {
            this.GetRange(this.RegisterSheet).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            this.GetRange(this.RegisterSheet).Interior.ColorIndex = 0;

        }
        /// <summary>
        /// Функиця пересчета трудоемкостей всех типов работ исходя из проставленных в трудоемкостей
        /// в работах типа КС-2
        /// </summary>
        public void CalcLabourness()
        {
            foreach (WorksSection section in this.WorksSections)
            {
                foreach (MSGWork msg_work in section.MSGWorks)
                {
                    //  if (msg_work.Laboriousness == 0)
                    {
                        decimal common_vovr_laboueness = 0;
                        foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                        {
                            // if (vovr_work.Laboriousness == 0)
                            {
                                decimal common_ks_laboueness = 0;
                                foreach (KSWork ks_work in vovr_work.KSWorks)
                                {
                                    common_ks_laboueness += ks_work.ProjectQuantity * ks_work.Laboriousness;

                                    this.CalcLabournessCoefficiens(ks_work);
                                    foreach (RCWork rc_work in ks_work.RCWorks)
                                        rc_work.Laboriousness = ks_work.ProjectQuantity * ks_work.Laboriousness * rc_work.LabournessCoefficient / rc_work.ProjectQuantity;
                                }
                                vovr_work.Laboriousness = common_ks_laboueness / vovr_work.ProjectQuantity;
                            }
                            common_vovr_laboueness += vovr_work.ProjectQuantity * vovr_work.Laboriousness;
                        }
                        msg_work.Laboriousness = common_vovr_laboueness / msg_work.ProjectQuantity;
                    }
                }
            }
        }
        private void CalcLabournessCoefficiens(KSWork ks_work)
        {
            var rc_works_with_notNull_labourness = ks_work.RCWorks.Where(rcw => rcw.Laboriousness != 0);
            decimal rc_laboriousness_coeffecients_sum = 0;
            foreach (RCWork rc_work in rc_works_with_notNull_labourness)
                rc_work.LabournessCoefficient = rc_work.Laboriousness * rc_work.ProjectQuantity / (ks_work.Laboriousness * ks_work.ProjectQuantity);

            var rc_works_with_notNull_labourness_coef = ks_work.RCWorks.Where(rcw => rcw.LabournessCoefficient != 0);
            foreach (RCWork rc_work in rc_works_with_notNull_labourness_coef)
                rc_laboriousness_coeffecients_sum += rc_work.LabournessCoefficient;

            var rc_works_with_Null_labourness_coef = ks_work.RCWorks.Where(rcw => rcw.LabournessCoefficient == 0).ToList();
            foreach (RCWork rc_work in rc_works_with_Null_labourness_coef)
            {
                decimal coef = (1 - rc_laboriousness_coeffecients_sum) / rc_works_with_Null_labourness_coef.ToList().Count;
                if (coef <= 0)
                {
                    rc_work.GetRange(this.RegisterSheet).Interior.Color = XlRgbColor.rgbRed;
                    throw new Exception("Кофицент должен быть больше нуля!");
                }
                rc_work.LabournessCoefficient = coef;
            }


        }
        /// <summary>
        /// Функцич подсчета объемов выполненных работ 
        /// </summary>
        public void CalcQuantity()
        {
            List<MSGExellModel> loaded_models = new List<MSGExellModel>();
            foreach (WorksSection section in this.WorksSections)
            {
                foreach (MSGWork msg_work in section.MSGWorks)
                {

                    var msg_work_all_rcWorks = this.RCWorks.Where(w => w.Number.StartsWith(msg_work.Number + "."));
                    foreach (RCWork rc_work in msg_work_all_rcWorks)
                    {
                        if (rc_work.ReportCard != null)
                        {
                            foreach (WorkDay ks_w_day in rc_work.ReportCard)
                            {
                                if (msg_work.ReportCard == null) { 
                                    msg_work.ReportCard = new WorkReportCard();
                                    msg_work.ReportCard.Number = msg_work.Number;
                                }
                                WorkDay msg_w_day = msg_work.ReportCard.FirstOrDefault(wd => wd.Date == ks_w_day.Date);
                                if (msg_w_day == null)
                                {
                                    msg_w_day = new WorkDay();
                                    msg_w_day.Date = ks_w_day.Date;
                                    msg_w_day.LaborСosts += ks_w_day.LaborСosts;

                                }
                                else
                                    msg_w_day.LaborСosts += ks_w_day.LaborСosts;
                                if (msg_work.Laboriousness != 0)
                                    msg_w_day.Quantity = msg_w_day.LaborСosts / msg_work.Laboriousness;
                                msg_work.ReportCard.Add(msg_w_day);
                            }
                        }

                    }


                }
            }
        }
        public void CalcQuantity_backup()
        {
            List<MSGExellModel> loaded_models = new List<MSGExellModel>();
            foreach (WorksSection section in this.WorksSections)
            {
                foreach (MSGWork msg_work in section.MSGWorks)
                {
                    msg_work.Quantity = 0;
                    decimal common_vovr_labour_quantity = 0;
                    decimal common_vovr_previos_complate_labour_quantity = 0;
                    foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                    {
                        vovr_work.Quantity = 0;
                        decimal common_ks_labour_quantity = 0;
                        decimal common_ks_previos_complate_labour_quantity = 0;
                        foreach (KSWork ks_work in vovr_work.KSWorks)
                        {
                            decimal common_rc_labour_quantity = 0;
                            decimal common_rc_previos_complate_labour_quantity = 0;
                            foreach (RCWork rc_work in ks_work.RCWorks)
                            {
                                rc_work.Quantity = 0;
                                if (this.Owner != null && rc_work.ReportCard != null)
                                {
                                    rc_work.Quantity = rc_work.ReportCard.Quantity + rc_work.ReportCard.PreviousComplatedQuantity;

                                }
                                else
                                {
                                    rc_work.PreviousComplatedQuantity = 0;

                                    rc_work.ReportCard = new WorkReportCard();
                                    this.Register(rc_work.ReportCard, "Number", rc_work.CellAddressesMap["Number"].Row, WRC_NUMBER_COL, this.RegisterSheet);
                                    this.Register(rc_work.ReportCard, "PreviousComplatedQuantity", rc_work.CellAddressesMap["Number"].Row, WRC_NUMBER_COL, this.RegisterSheet);
                                    rc_work.ReportCard.Number = rc_work.Number;
                                    int rc_work_common_quantity = 0;
                                    foreach (MSGExellModel model in this.Children)
                                    {
                                        if (!loaded_models.Contains(model))
                                        {
                                            model.LoadWorksReportCards();
                                            loaded_models.Add(model);
                                        }
                                        // RCWork child_rc_work = model.RCWorks.FirstOrDefault(w => w.Number == rc_work.Number);
                                        WorkReportCard child_rc = model.WorkReportCards.FirstOrDefault(rc => rc.Number == rc_work.Number);


                                        if (child_rc != null)
                                        {

                                            foreach (WorkDay child_w_day in child_rc)
                                            {
                                                WorkDay curent_w_day = rc_work.ReportCard.FirstOrDefault(wd => wd.Date == child_w_day.Date);
                                                if (curent_w_day != null)
                                                {
                                                    curent_w_day.Quantity += child_w_day.Quantity;
                                                    curent_w_day.LaborСosts = curent_w_day.Quantity * rc_work.Laboriousness;
                                                }
                                                else
                                                {
                                                    curent_w_day = new WorkDay();
                                                    curent_w_day.Date = child_w_day.Date;
                                                    curent_w_day.Quantity = child_w_day.Quantity;
                                                    curent_w_day.LaborСosts = child_w_day.Quantity * rc_work.Laboriousness;
                                                    DateTime end_date = DateTime.Parse(this.RegisterSheet.Cells[WORKS_END_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());

                                                    foreach (KeyValuePair<string, ExellPropAddress> map_item in child_w_day.CellAddressesMap)
                                                    {
                                                        int date_index = 0;
                                                        while (this.RegisterSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value != null &&
                                                                   DateTime.Parse(this.RegisterSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value.ToString()) < end_date)
                                                        {
                                                            if (DateTime.Parse(this.RegisterSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value.ToString()) == curent_w_day.Date)
                                                                break;
                                                            date_index++;
                                                        }
                                                        if (rc_work.ReportCard.CellAddressesMap.ContainsKey("Number"))
                                                        {
                                                            int curent_wrc_row = rc_work.ReportCard.CellAddressesMap["Number"].Row;
                                                            this.Register(curent_w_day, "Quantity", curent_wrc_row, WRC_DATE_COL + date_index, this.RegisterSheet);
                                                            curent_w_day.Quantity = curent_w_day.Quantity;
                                                        }

                                                    }
                                                    rc_work.ReportCard.Add(curent_w_day);
                                                }
                                            }
                                            rc_work.PreviousComplatedQuantity += child_rc.PreviousComplatedQuantity;
                                            rc_work.Quantity += child_rc.Quantity;
                                        }

                                    }
                                }
                                common_rc_labour_quantity += rc_work.Quantity * rc_work.Laboriousness;
                                common_rc_previos_complate_labour_quantity += rc_work.PreviousComplatedQuantity * rc_work.Laboriousness;
                            }

                            ks_work.Quantity = common_rc_labour_quantity / ks_work.Laboriousness;
                            ks_work.PreviousComplatedQuantity = common_rc_previos_complate_labour_quantity / ks_work.Laboriousness;

                            if (ks_work.Laboriousness != 0)
                            {
                                common_ks_labour_quantity += ks_work.Quantity * ks_work.Laboriousness;
                                common_ks_previos_complate_labour_quantity += ks_work.PreviousComplatedQuantity * ks_work.Laboriousness;
                            }
                        }

                        if (vovr_work.Laboriousness != 0)
                        {
                            vovr_work.Quantity = common_ks_labour_quantity / vovr_work.Laboriousness;
                            vovr_work.PreviousComplatedQuantity = common_ks_previos_complate_labour_quantity / vovr_work.Laboriousness;
                        }
                        common_vovr_labour_quantity += vovr_work.Quantity * vovr_work.Laboriousness;
                        common_vovr_previos_complate_labour_quantity += vovr_work.PreviousComplatedQuantity * vovr_work.Laboriousness;
                    }

                    if (msg_work.Laboriousness != 0)
                    {
                        msg_work.Quantity = common_vovr_labour_quantity / msg_work.Laboriousness;
                        msg_work.PreviousComplatedQuantity = common_vovr_previos_complate_labour_quantity / msg_work.Laboriousness;
                    }

                    var msg_work_all_rcWorks = this.RCWorks.Where(w => w.Number.StartsWith(msg_work.Number + "."));
                    foreach (RCWork rc_work in msg_work_all_rcWorks)
                    {
                        if (rc_work.ReportCard != null)
                        {
                            foreach (WorkDay ks_w_day in rc_work.ReportCard)
                            {
                                if (msg_work.ReportCard == null) msg_work.ReportCard = new WorkReportCard();
                                WorkDay msg_w_day = msg_work.ReportCard.FirstOrDefault(wd => wd.Date == ks_w_day.Date);
                                if (msg_w_day == null)
                                {
                                    msg_w_day = new WorkDay();
                                    msg_w_day.Date = ks_w_day.Date;
                                    msg_w_day.LaborСosts += ks_w_day.LaborСosts;

                                }
                                else
                                    msg_w_day.LaborСosts += ks_w_day.LaborСosts;
                                if (msg_work.Laboriousness != 0)
                                    msg_w_day.Quantity = msg_w_day.LaborСosts / msg_work.Laboriousness;
                                msg_work.ReportCard.Add(msg_w_day);
                            }
                        }

                    }


                }
            }
        }
        public void CalcWorkerConsumptions()
        {
            Excel.Worksheet consumtionsSheet = this.WorkerConsumptionsSheet;
            int rowIndex = W_CONSUMPTIONS_FIRST_ROW_INDEX;
            // this.WorkerConsumptions.Clear();
            null_str_count = 0;


            if (this.Owner == null)
                foreach (WorkerConsumption worker_consumption in this.WorkerConsumptions)
                {

                    foreach (MSGExellModel model in this.Children)
                    {

                        WorkerConsumption child_w_consumption = model.WorkerConsumptions.FirstOrDefault(w => w.Number == worker_consumption.Number);
                        if (child_w_consumption != null)
                        {

                            foreach (WorkerConsumptionDay child_w_day in child_w_consumption.WorkersConsumptionReportCard)
                            {
                                WorkerConsumptionDay curent_w_day = worker_consumption.WorkersConsumptionReportCard.FirstOrDefault(wd => wd.Date == child_w_day.Date);
                                if (curent_w_day != null)
                                {
                                    curent_w_day.Quantity += child_w_day.Quantity;
                                }
                                else
                                {
                                    curent_w_day = new WorkerConsumptionDay();
                                    //  this.Register(curent_w_day);
                                    curent_w_day.Date = child_w_day.Date;

                                    curent_w_day.Quantity = child_w_day.Quantity;

                                    DateTime end_date = DateTime.Parse(this.RegisterSheet.Cells[WORKS_END_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());

                                    foreach (KeyValuePair<string, ExellPropAddress> map_item in child_w_day.CellAddressesMap)
                                    {
                                        int date_index = 0;
                                        while (this.WorkerConsumptionsSheet.Cells[W_CONSUMPTIONS_DATE_RAW, W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value != null &&
                                          DateTime.Parse(this.WorkerConsumptionsSheet.Cells[W_CONSUMPTIONS_DATE_RAW, W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString()) < end_date)
                                        {
                                            if (DateTime.Parse(this.WorkerConsumptionsSheet.Cells[W_CONSUMPTIONS_DATE_RAW, W_CONSUMPTIONS_FIRST_DATE_COL + date_index].Value.ToString()) == curent_w_day.Date)
                                                break;
                                            date_index++;
                                        }
                                        int curent_w_consumption_row = worker_consumption.WorkersConsumptionReportCard.CellAddressesMap["WorkersConsumptionReportCard"].Row;

                                        curent_w_day.CellAddressesMap.Add(map_item.Key, new ExellPropAddress(curent_w_consumption_row, W_CONSUMPTIONS_FIRST_DATE_COL + date_index, this.WorkerConsumptionsSheet));
                                        curent_w_day.Quantity = curent_w_day.Quantity;
                                        //          this.WorkerConsumptionsSheet.Cells[curent_w_consumption_row, W_CONSUMPTIONS_FIRST_DATE_COL + date_index] =
                                        //      curent_w_day.Quantity.ToString();

                                        //   this.Register(curent_w_day);
                                    }
                                    worker_consumption.WorkersConsumptionReportCard.Add(curent_w_day);
                                }
                            }
                        }

                    }
                }


        }
        /// <summary>
        /// Вычисление всех вычисляемых величин внутри модели и всех его дочерних моделей.
        /// </summary>
        public void CalcAll()
        {
            if (this.Owner == null)
                this.ReloadSheetModel();
            else
                this.ReloadAllSheetsInModel();

            this.CalcLabourness();
            //    this.CalcQuantity();
            this.CalcWorkerConsumptions();
            this.SetStyleFormats();
            //   this.SetFormulas();
        }


        /// <summary>
        /// Функция сбрасывает в значение 0 все вычиляемые поля всех работ кроме  поля KSWork.Laboriousness
        /// </summary>
        public void ResetCalculatesFields()
        {
            foreach (MSGWork work in this.MSGWorks)
            {
                work.Quantity = 0;
                work.Laboriousness = 0;
            }
            foreach (VOVRWork work in this.VOVRWorks)
            {
                work.Quantity = 0;
                work.Laboriousness = 0;
            }
            foreach (KSWork work in this.KSWorks)
            {
                work.Quantity = 0;
            }
        }
        /// <summary>
        /// Функция обновляет разделы МСГ, ВОВР и КС-2 ведомости если модель является дочерней ( у нее есть владелец) 
        /// или если ведомость сама общая, то просто очищает у нее каледарную часть с записями выполенных объемов
        /// </summary>
        public void UpdateWorksheetCommonPart(bool erase_common_part)
        {
            if (this.Owner != null)
                this.CopyOwnerObjectModels();

            this.UpdateCellAddressMapsWorkSheets();

            this.ClearWorksheetCommonPart(erase_common_part);

            int last_row = FIRST_ROW_INDEX;
            foreach (WorksSection w_section in this.WorksSections.OrderBy(s => s.Number))
            {
                last_row = this.UpdateSectionWorksheetCommonPart(w_section, last_row) + _SECTIONS_GAP;
            }

        }
        public void CopyOwnerObjectModels()
        {

            this.WorksSections = (ExcelNotifyChangedCollection<WorksSection>)this.Owner.WorksSections.Clone();
            foreach (var section in this.WorksSections)
                this.RegisterObjectInObjectPropertyNameRegister(section);

        }
        public int UpdateSectionWorksheetCommonPart(WorksSection w_section, int top_row = FIRST_ROW_INDEX)
        {
            int section_row = top_row;
            int rc_row = top_row;
            int ks_row = top_row;
            int vovr_row = top_row;
            int msg_row = top_row;
            w_section.ChangeTopRow(section_row);
            this.UpdateExellBindableObject(w_section);
            foreach (MSGWork msg_work in w_section.MSGWorks)
            {
                msg_work.ChangeTopRow(msg_row);
                this.UpdateExellBindableObject(msg_work);
                int sh_ch_row_iterator = 0;
                foreach (WorkScheduleChunk w_ch in msg_work.WorkSchedules)
                {
                    w_ch.ChangeTopRow(msg_row + sh_ch_row_iterator);
                    this.UpdateExellBindableObject(w_ch);
                    sh_ch_row_iterator++;
                }
                int nw_row_iterator = 0;
                foreach (NeedsOfWorker n_w in msg_work.WorkersComposition)
                {
                    n_w.ChangeTopRow(msg_row + nw_row_iterator);
                    this.UpdateExellBindableObject(n_w);
                    nw_row_iterator++;
                }
                vovr_row = section_row;
                foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                {
                    vovr_work.ChangeTopRow(vovr_row);
                    this.UpdateExellBindableObject(vovr_work);
                    ks_row = vovr_row;
                    foreach (KSWork ks_work in vovr_work.KSWorks)
                    {
                        ks_work.ChangeTopRow(ks_row);
                        this.UpdateExellBindableObject(ks_work);
                        rc_row = ks_row;
                        foreach (RCWork rc_work in ks_work.RCWorks)
                        {
                            rc_work.ChangeTopRow(rc_row);
                            this.UpdateExellBindableObject(rc_work);

                            rc_work.ReportCard = this.WorkReportCards.Where(rc => rc.Number == rc_work.Number).FirstOrDefault();
                            if (rc_work.ReportCard != null)
                            {
                                rc_work.ReportCard.ChangeTopRow(rc_row);
                                this.UpdateExellBindableObject(rc_work.ReportCard);
                                foreach (WorkDay w_day in rc_work.ReportCard)
                                {
                                    w_day.ChangeTopRow(rc_work.CellAddressesMap["Number"].Row);
                                    this.UpdateExellBindableObject(w_day);
                                }
                            }


                            rc_row++;
                        }
                        ks_row = rc_row;
                        rc_row = ks_row;
                    }

                    vovr_row = ks_row;
                }
                msg_row = vovr_row;
            }
            section_row = msg_row + 1;
            return rc_row;
        }
        public void UpdateSectionWorksheetCommonPart_backup(WorksSection w_section, int top_row = FIRST_ROW_INDEX)
        {
            int section_row = top_row;
            int rc_row = top_row;
            int ks_row = top_row;
            int vovr_row = top_row;
            int msg_row = top_row;
            w_section.ChangeTopRow(section_row);
            this.UpdateExellBindableObject(w_section);
            foreach (MSGWork msg_work in w_section.MSGWorks)
            {
                msg_work.ChangeTopRow(msg_row);
                this.UpdateExellBindableObject(msg_work);
                int sh_ch_row_iterator = 0;
                foreach (WorkScheduleChunk w_ch in msg_work.WorkSchedules)
                {
                    w_ch.ChangeTopRow(msg_row + sh_ch_row_iterator);
                    this.UpdateExellBindableObject(w_ch);
                    sh_ch_row_iterator++;
                }
                int nw_row_iterator = 0;
                foreach (NeedsOfWorker n_w in msg_work.WorkersComposition)
                {
                    n_w.ChangeTopRow(msg_row + nw_row_iterator);
                    this.UpdateExellBindableObject(n_w);
                    nw_row_iterator++;
                }
                vovr_row = section_row;
                foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                {
                    vovr_work.ChangeTopRow(vovr_row);
                    this.UpdateExellBindableObject(vovr_work);
                    ks_row = vovr_row;
                    foreach (KSWork ks_work in vovr_work.KSWorks)
                    {
                        ks_work.ChangeTopRow(ks_row);
                        this.UpdateExellBindableObject(ks_work);
                        rc_row = ks_row;
                        foreach (RCWork rc_work in ks_work.RCWorks)
                        {
                            rc_work.ChangeTopRow(rc_row);
                            this.UpdateExellBindableObject(rc_work);

                            rc_work.ReportCard = this.WorkReportCards.Where(rc => rc.Number == rc_work.Number).FirstOrDefault();
                            if (rc_work.ReportCard != null)
                            {
                                rc_work.ReportCard.ChangeTopRow(rc_row);
                                this.UpdateExellBindableObject(rc_work.ReportCard);
                                foreach (WorkDay w_day in rc_work.ReportCard)
                                {
                                    w_day.ChangeTopRow(rc_row);
                                    this.UpdateExellBindableObject(w_day);
                                }
                            }


                            rc_row++;
                        }
                        ks_row = rc_row;
                        rc_row = ks_row;
                    }

                    vovr_row = ks_row;
                }
                msg_row = vovr_row;
            }
            section_row = msg_row + 1;
        }
        public void Update()
        {
            bool erase_common_part = false;

            if (this.Owner != null)
            {
                erase_common_part = true;

            }
            this.ReloadSheetModel();
            this.UpdateWorksheetCommonPart(erase_common_part);
            this.SetFormulas();
            this.SetStyleFormats();

        }
        /// <summary>
        /// Функция устанавливаетв объектах текущией модели соотвествующие worksheet-ы 
        /// Применяется в основном после применения Clone()  к объектоной модели MSGExcellModel
        /// </summary>
        public void UpdateCellAddressMapsWorkSheets()
        {


            this.WorksSections.CellAddressesMap.SetWorksheet(this.RegisterSheet);
            foreach (WorksSection w_section in this.WorksSections)
            {
                w_section.CellAddressesMap.SetWorksheet(this.RegisterSheet);
                foreach (MSGWork msg_work in w_section.MSGWorks)
                {
                    msg_work.CellAddressesMap.SetWorksheet(this.RegisterSheet);
                    if (!this.MSGWorks.Contains(msg_work)) this.MSGWorks.Add(msg_work);

                    foreach (WorkScheduleChunk w_ch in msg_work.WorkSchedules)
                        w_ch.CellAddressesMap.SetWorksheet(this.RegisterSheet);
                    foreach (NeedsOfWorker n_w in msg_work.WorkersComposition)
                        n_w.CellAddressesMap.SetWorksheet(this.RegisterSheet);

                    foreach (VOVRWork vovr_work in msg_work.VOVRWorks)
                    {
                        if (!this.VOVRWorks.Contains(vovr_work)) this.VOVRWorks.Add(vovr_work);
                        vovr_work.CellAddressesMap.SetWorksheet(this.RegisterSheet);
                        foreach (KSWork ks_work in vovr_work.KSWorks)
                        {
                            if (!this.KSWorks.Contains(ks_work)) this.KSWorks.Add(ks_work);
                            ks_work.CellAddressesMap.SetWorksheet(this.RegisterSheet);
                            foreach (RCWork rc_work in ks_work.RCWorks)
                            {
                                if (!this.RCWorks.Contains(rc_work)) this.RCWorks.Add(rc_work);
                                rc_work.ReportCard = this.WorkReportCards.Where(rc => rc.Number == rc_work.Number).FirstOrDefault();
                                rc_work.CellAddressesMap.SetWorksheet(this.RegisterSheet);

                            }
                        }
                    }
                }

            }
        }

        /// <summary>
        /// Функция обновляет документальное представление объетка (рукурсивно проходит по всем объектам 
        /// реализующим интерфейс IExcelBindableBase). 
        /// </summary>
        /// <param name="obj">Связанный с докуметом Worksheet объект рализующий IExcelBindableBase </param>
        private void UpdateExellBindableObject(IExcelBindableBase obj)
        {
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

                    if (prop_value is DateTime date_val)
                        return date_val.ToString("d");
                    else
                        return prop_value.ToString();
                }
                else
                    return "";
            }
            return null;
        }


        /// <summary>
        /// Фунция очищает календарную часть ведомости (очищает все записи выполненных работ)
        /// </summary>
        public void ClearWorksheetDaysPart()
        {
            //    Excel.Range last_cell = this.RegisterSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            Excel.Range common_area_range = this.RegisterSheet.Range[this.RegisterSheet.Cells[FIRST_ROW_INDEX, WRC_DATE_COL],
                this.RegisterSheet.Cells[this.KSWorks[this.KSWorks.Count - 1].CellAddressesMap["Laboriousness"].Row,
                                          this.KSWorks[this.KSWorks.Count - 1].CellAddressesMap["Laboriousness"].Column]];
            if (this.Owner != null)
                common_area_range.ClearContents();

            common_area_range.Interior.ColorIndex = 0;

            //  last_cell = this.WorkerConsumptionsSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            var last_work_consumption_RC = this.WorkerConsumptions[this.WorkerConsumptions.Count - 1].WorkersConsumptionReportCard;
            WorkerConsumptionDay last_work_consumption_Day = null;
            foreach (WorkerConsumption worker_consumption in this.WorkerConsumptions)
            {
                foreach (WorkerConsumptionDay worker_consumption_day in worker_consumption.WorkersConsumptionReportCard)
                {
                    if (last_work_consumption_Day == null || last_work_consumption_Day.Date < worker_consumption_day.Date)
                        last_work_consumption_Day = worker_consumption_day;
                }
            }

            if (last_work_consumption_Day != null)
            {
                common_area_range = this.WorkerConsumptionsSheet.Range[this.WorkerConsumptionsSheet.Cells[W_CONSUMPTIONS_FIRST_ROW_INDEX, W_CONSUMPTIONS_FIRST_DATE_COL],
                            this.WorkerConsumptionsSheet.Cells[last_work_consumption_Day.CellAddressesMap["Quantity"].Row, last_work_consumption_Day.CellAddressesMap["Quantity"].Column]];
                common_area_range.ClearContents();
                common_area_range.Interior.ColorIndex = 0;
            }


        }
        /// <summary>
        /// Функия очищает левую часть вдомости с МСГ, ВОВР и КС-2.
        /// </summary>
        public void ClearWorksheetCommonPart(bool erase_common_part)
        {
            Excel.Range last_cell = this.RegisterSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);

            Excel.Range common_area_range = this.RegisterSheet.Range[this.RegisterSheet.Cells[FIRST_ROW_INDEX, WSEC_NUMBER_COL],
                  this.RegisterSheet.Cells[last_cell.Row, WRC_NUMBER_COL - 1]];
            common_area_range.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            common_area_range.Interior.ColorIndex = 0;

            Excel.Range record_cards_area_range = this.RegisterSheet.Range[this.RegisterSheet.Cells[FIRST_ROW_INDEX, RC_NUMBER_COL],
                  this.RegisterSheet.Cells[last_cell.Row, last_cell.Column]];
            record_cards_area_range.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            record_cards_area_range.Interior.ColorIndex = 0;

            common_area_range.ClearContents();

            if (this.Owner == null)
                record_cards_area_range.ClearContents();



            Excel.Range all_rows = this.RegisterSheet.Cells.Rows;
            try
            {
                all_rows.Ungroup();

            }
            catch
            {

            }

        }


        public IExcelBindableBase GetObjectBySelection(Excel.Range section, Type object_type)
        {
            ObservableCollection<Tuple<double, IExcelBindableBase>> objects_distation = new ObservableCollection<Tuple<double, IExcelBindableBase>>();
            foreach (var kvp in this.RegistedObjects)
            {
                int obj_row = kvp.ExellPropAddress.Row;
                int obj_col = kvp.ExellPropAddress.Column;
                var dist = Math.Sqrt((obj_row = section.Row) ^ 2 + (obj_col = section.Column) ^ 2);
                objects_distation.Add(new Tuple<double, IExcelBindableBase>(dist, kvp.Entity));
            }
            IExcelBindableBase finded_obj = null;
            var tuple = objects_distation.Where(tp => tp.Item2.GetType() == object_type).OrderBy(el => el.Item1).FirstOrDefault();

            if (tuple != null)
                finded_obj = tuple.Item2 as IExcelBindableBase;

            return finded_obj;
        }
    }
}
