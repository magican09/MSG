using ExellAddInsLib.MSG.Section;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
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
        public const int KS_PC_QUANTITY_COL = WRC_NUMBER_COL + 1;

        public const int WRC_DATE_ROW = 6;

        public const int WRC_NUMBER_COL = 28;
        public const int WRC_DATE_COL = KS_PC_QUANTITY_COL + 1;

        public const int W_CONSUMPTIONS_FIRST_ROW_INDEX = 4;
        public const int W_CONSUMPTIONS_NUMBER_COL = 1;
        public const int W_CONSUMPTIONS_NAME_COL = 2;
        public const int W_CONSUMPTIONS_DATE_RAW = 3;
        public const int W_CONSUMPTIONS_FIRST_DATE_COL = 3;


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
        public ExcelNotifyChangedCollection<WorksSection> WorksSections { get; private set; } = new ExcelNotifyChangedCollection<WorksSection>();
        /// <summary>
        /// Коллекция с работами типа МСГ модели
        /// </summary>
        public ExcelNotifyChangedCollection<MSGWork> MSGWorks { get; private set; } = new ExcelNotifyChangedCollection<MSGWork>();
        /// <summary>
        /// Коллекция с работами типа ВОВР модели
        /// </summary>
        public ExcelNotifyChangedCollection<VOVRWork> VOVRWorks { get; private set; } = new ExcelNotifyChangedCollection<VOVRWork>();
        /// <summary>
        /// Коллекция с работами типа КС-2 модели
        /// </summary>
        public ObservableCollection<KSWork> KSWorks { get; private set; } = new ObservableCollection<KSWork>();
        /// <summary>
        /// Коллекция  табелей выполненных работ
        /// </summary>
        public ObservableCollection<WorkReportCard> WorkReportCards { get; private set; } = new ObservableCollection<WorkReportCard>();
        /// <summary>
        /// Коллекция с единицами измерения модели
        /// </summary>
        public ObservableCollection<UnitOfMeasurement> UnitOfMeasurements { get; set; } = new ObservableCollection<UnitOfMeasurement>();
        //public WorkersCompositionReportCard WorkersCompositionReportCard = new MSG.WorkersCompositionReportCard();
        public WorkersComposition WorkersComposition { get; set; } = new WorkersComposition();
        public WorkerConsumptions WorkerConsumptions { get; set; } = new WorkerConsumptions();

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
        public Excel.Worksheet RegisterSheet { get; set; }
        /// <summary>
        /// Прикрепленный к модели лист  Людских ресурсов Worksheet
        /// </summary>
        public Excel.Worksheet WorkerConsumptionsSheet { get; set; }

        /// <summary>
        /// Прикрепленный к модели лист общих данных  Worksheet
        /// </summary>
        public Excel.Worksheet CommonSheet { get; set; }
        /// <summary>
        /// Отвественных за работы отраженных в работах данной модели
        /// </summary>
        public Employer Employer { get; set; }
        public MSGExellModel()
        {

        }
        // private ObservableCollection<IExcelBindableBase> RegistedObjects = new ObservableCollection<IExcelBindableBase>();
        // private EventedDictationary<RalateRecord, Tuple<string, ExellPropAddress>> RegistedObjects = new EventedDictationary<RalateRecord, Tuple<string, ExellPropAddress>>();
        // private EventedDictationary<RalateRecord, ExellPropAddress> RegistedObjects = new EventedDictationary<RalateRecord, ExellPropAddress>();
        private ObservableCollection<RelateRecord> RegistedObjects = new ObservableCollection<RelateRecord>();
        // private EventedDictationary<IExcelBindableBase, string> ObjectPropertyNameRegister = new EventedDictationary<IExcelBindableBase, string>();
        private ObservableCollection<RelateRecord> ObjectPropertyNameRegister = new ObservableCollection<RelateRecord>();
        private Dictionary<IExcelBindableBase, string> RegisterTemporalStopList = new Dictionary<IExcelBindableBase, string>();
        /// <summary>
        /// Функция для регистрации объекта реализующего интрефейс INotifyPropertyChanged 
        /// для обработки событий изменения полей объета и соотвествующего изменения связанной с 
        /// с этим полем ячейки в документе Worksheet
        /// </summary>
        /// <param name="work"></param>
        public void Register(IExcelBindableBase notified_object, string prop_name, int row, int column, Excel.Worksheet worksheet, RelateRecord register = null)
        {

            RelateRecord local_register = new RelateRecord(notified_object);
            if (register == null)
            {
                register = local_register;
                local_register.ExellPropAddress = new ExellPropAddress(row, column, worksheet, prop_name);
                RegistedObjects.Add(local_register);
                RegisterTemporalStopList.Clear();

                RegisterTemporalStopList.Add(local_register.Entity, prop_name);
                local_register.PropertyName = prop_name;
                ObjectPropertyNameRegister.Add(local_register);
            }
            else
                register.Items.Add(local_register);

            var prop_names = prop_name.Split(new char[] { '.' });


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



            //switch (notified_object.GetType().Name)
            //{

            //    case nameof(WorksSection):
            //        {

            //            WorksSection w_section = (WorksSection)obj;
            //            if (!this.WorksSections.Contains(w_section))
            //                this.WorksSections.Add(w_section);
            //            break;
            //        }

            //    case nameof(MSGWork):
            //        {

            //            MSGWork msg_work = (MSGWork)obj;
            //            if (!this.MSGWorks.Contains(msg_work))
            //                this.MSGWorks.Add(msg_work);

            //            WorksSection w_section = this.WorksSections.Where(ws => ws.Number.StartsWith(msg_work.Number.Remove(msg_work.Number.LastIndexOf(".")))).FirstOrDefault();
            //            if (w_section != null)
            //            {
            //                w_section.MSGWorks.Add(msg_work);
            //            }
            //            break;
            //        }
            //    case nameof(NeedsOfWorker):
            //        {
            //            NeedsOfWorker needs_of_workers = (NeedsOfWorker)obj;

            //            MSGWork msg_work = this.MSGWorks.Where(w => w.Number.StartsWith(needs_of_workers.Number.Remove(needs_of_workers.Number.LastIndexOf(".")))).FirstOrDefault();
            //            if (msg_work != null)
            //            {
            //                msg_work.WorkersComposition.Add(needs_of_workers);
            //                needs_of_workers.Owner = msg_work;
            //                foreach (WorkScheduleChunk chunk in msg_work.WorkSchedules)
            //                {
            //                    for (DateTime date = chunk.StartTime; date <= chunk.EndTime; date = date.AddDays(1))
            //                    {
            //                        NeedsOfWorkersDay needsOfWorkersDay = new NeedsOfWorkersDay();
            //                        needsOfWorkersDay.Date = date;
            //                        needsOfWorkersDay.Quantity = needs_of_workers.Quantity;
            //                        needs_of_workers.NeedsOfWorkersReportCard.Add(needsOfWorkersDay);
            //                    }
            //                }
            //            }

            //            NeedsOfWorker global_needs_of_worker = this.WorkersComposition.FirstOrDefault(nw => nw.Name == needs_of_workers.Name);
            //            if (global_needs_of_worker == null)
            //            {
            //                global_needs_of_worker = new NeedsOfWorker();
            //                global_needs_of_worker.Number = needs_of_workers.Number;
            //                global_needs_of_worker.Name = needs_of_workers.Name;
            //                foreach (NeedsOfWorkersDay needsOfWorkersDay in needs_of_workers.NeedsOfWorkersReportCard)
            //                    global_needs_of_worker.NeedsOfWorkersReportCard.Add(needsOfWorkersDay);
            //                this.WorkersComposition.Add(global_needs_of_worker);
            //            }
            //            else
            //            {
            //                foreach (NeedsOfWorkersDay needsOfWorkersDay in needs_of_workers.NeedsOfWorkersReportCard)
            //                {
            //                    var nw_day = global_needs_of_worker.NeedsOfWorkersReportCard.FirstOrDefault(nwd => nwd.Date == needsOfWorkersDay.Date);
            //                    if (nw_day != null)
            //                    {
            //                        nw_day.Quantity += needsOfWorkersDay.Quantity;
            //                    }
            //                    else
            //                    {
            //                        NeedsOfWorkersDay new_nw_day = new NeedsOfWorkersDay(needsOfWorkersDay);
            //                        global_needs_of_worker.NeedsOfWorkersReportCard.Add(new_nw_day);
            //                    }
            //                }

            //            }

            //            break;
            //        }
            //    case nameof(VOVRWork):
            //        {
            //            VOVRWork vovr_work = (VOVRWork)obj;
            //            if (!this.VOVRWorks.Contains(vovr_work))
            //                this.VOVRWorks.Add(vovr_work);

            //            MSGWork msg_work = this.MSGWorks.Where(w => w.Number.StartsWith(vovr_work.Number.Remove(vovr_work.Number.LastIndexOf(".")))).FirstOrDefault();
            //            if (msg_work != null)
            //            {
            //                msg_work.VOVRWorks.Add(vovr_work);
            //            }

            //            break;
            //        }
            //    case nameof(KSWork):
            //        {
            //            KSWork ks_work = (KSWork)obj;
            //            if (!this.KSWorks.Contains(ks_work))
            //                this.KSWorks.Add(ks_work);

            //            VOVRWork vovr_work = VOVRWorks.Where(w => w.Number.StartsWith(ks_work.Number.Remove(ks_work.Number.LastIndexOf(".")))).FirstOrDefault();
            //            if (vovr_work != null)
            //                vovr_work.KSWorks.Add(ks_work);

            //            break;
            //        }

            //    case nameof(WorkReportCard):
            //        {
            //            WorkReportCard report_card = (WorkReportCard)obj;
            //            if (!this.WorkReportCards.Contains(report_card))
            //                this.WorkReportCards.Add(report_card);

            //            KSWork ks_work = KSWorks.Where(w => w.Number == report_card.Number).FirstOrDefault();
            //            if (ks_work != null && report_card.Count > 0)
            //                ks_work.ReportCard = report_card;
            //            else if (ks_work == null)
            //            {
            //                ks_work = new KSWork();
            //                ks_work.Number = report_card.Number;
            //                ks_work.ReportCard = report_card;
            //                report_card.Owner = ks_work;
            //            }

            //            break;
            //        }
            //    case nameof(WorkerConsumption):
            //        {

            //            WorkerConsumption w_consumption = (WorkerConsumption)obj;
            //            if (!this.WorkerConsumptions.Contains(w_consumption))
            //                this.WorkerConsumptions.Add(w_consumption);
            //            break;
            //        }
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
        private void OnPropertyChange(object sender, PropertyChangedEventArgs e)
        {
            if (sender is IExcelBindableBase bindable_object)
            {
                //  var ralated_abjects = this.RegistedObjects.Where(reg => reg.Key.RalateObjects.ContainsKey(e.PropertyName));
                var ralated_records = this.RegistedObjects
                    .Where(rr => rr.Entity.Id == bindable_object.Id)
                    .Where(r => r.PropertyName == e.PropertyName);

                foreach (RelateRecord related_rec in ralated_records)
                {
                    var parent_rrecord = this.GetFirstParentRelateRecord(related_rec);
                    ObservableCollection<Tuple<RelateRecord, string>> all_children_records = new ObservableCollection<Tuple<RelateRecord, string>>(); ;
                    this.GetChildrenRelateRecords(parent_rrecord, all_children_records);
                    var children_for_read_props = all_children_records.Where(ch => ch.Item2 == parent_rrecord.ExellPropAddress.ProprertyName);
                    foreach (Tuple<RelateRecord, string> rr_tuple in children_for_read_props)
                    {
                        RelateRecord child_rr = rr_tuple.Item1;
                        var val = child_rr.Entity.GetType().GetProperty(child_rr.PropertyName).GetValue(child_rr.Entity).ToString();
                        parent_rrecord.ExellPropAddress.Cell.Value = val;
                        parent_rrecord.ExellPropAddress.Cell.Interior.Color = XlRgbColor.rgbCoral;
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

                        DateTime start_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_START_DATE_COL].Value.ToString());
                        DateTime end_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_END_DATE_COL].Value.ToString());
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
                        this.Register(msg_work, "Number", rowIndex, MSG_NUMBER_COL, this.RegisterSheet);
                        this.Register(msg_work, "Name", rowIndex, MSG_NAME_COL, this.RegisterSheet);
                        this.Register(msg_work, "ProjectQuantity", rowIndex, MSG_QUANTITY_COL, this.RegisterSheet);
                        this.Register(msg_work, "Quantity", rowIndex, MSG_QUANTITY_FACT_COL, this.RegisterSheet);
                        this.Register(msg_work, "Laboriousness", rowIndex, MSG_LABOURNESS_COL, this.RegisterSheet);
                        this.MSGWorks.Add(msg_work);
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
                            msg_needs_of_workers.Owner = msg_work;
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
                        vovr_work.CellAddressesMap.Add("Number", new ExellPropAddress(rowIndex, VOVR_NUMBER_COL, this.RegisterSheet));

                        vovr_work.Name = registerSheet.Cells[rowIndex, VOVR_NAME_COL].Value.ToString();
                        vovr_work.CellAddressesMap.Add("Name", new ExellPropAddress(rowIndex, VOVR_NAME_COL, this.RegisterSheet));
                        vovr_work.CellAddressesMap.Add("UnitOfMeasurement", new ExellPropAddress(rowIndex, VOVR_MEASURE_COL, this.RegisterSheet));

                        vovr_work.CellAddressesMap.Add("ProjectQuantity", new ExellPropAddress(rowIndex, VOVR_QUANTITY_COL, this.RegisterSheet));
                        vovr_work.CellAddressesMap.Add("Quantity", new ExellPropAddress(rowIndex, VOVR_QUANTITY_FACT_COL, this.RegisterSheet));
                        vovr_work.CellAddressesMap.Add("Laboriousness", new ExellPropAddress(rowIndex, VOVR_LABOURNESS_COL, this.RegisterSheet));

                        if (registerSheet.Cells[rowIndex, VOVR_MEASURE_COL].Value != null)
                        {
                            vovr_work.UnitOfMeasurement = new UnitOfMeasurement(registerSheet.Cells[rowIndex, VOVR_MEASURE_COL].Value);
                            vovr_work.UnitOfMeasurement.CellAddressesMap.Add("Name", new ExellPropAddress(rowIndex, VOVR_MEASURE_COL, this.RegisterSheet));

                            registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_MEASURE_COL], registerSheet.Cells[rowIndex, VOVR_MEASURE_COL]].Interior.Color
                                = XlRgbColor.rgbWhite;
                        }
                        else
                            registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_MEASURE_COL], registerSheet.Cells[rowIndex, VOVR_MEASURE_COL]].Interior.Color
                                = XlRgbColor.rgbRed;

                        if (registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL].Value != null)
                        {
                            vovr_work.ProjectQuantity = Decimal.Parse(registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL].Value.ToString());
                            registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL], registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL]].Interior.Color
                                = XlRgbColor.rgbWhite;

                        }
                        else
                            registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL], registerSheet.Cells[rowIndex, VOVR_QUANTITY_COL]].Interior.Color
                                = XlRgbColor.rgbRed;

                        if (registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL].Value != null)
                        {
                            vovr_work.Laboriousness = Decimal.Parse(registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL].Value.ToString());
                            registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL], registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL]].Interior.Color
                                = XlRgbColor.rgbWhite;
                        }
                        else
                            registerSheet.Range[registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL], registerSheet.Cells[rowIndex, VOVR_LABOURNESS_COL]].Interior.Color
                                = XlRgbColor.rgbRed;

                        // this.Register(vovr_work);

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
                        ks_work.CellAddressesMap.Add("Number", new ExellPropAddress(rowIndex, KS_NUMBER_COL, this.RegisterSheet));

                        ks_work.Code = registerSheet.Cells[rowIndex, KS_CODE_COL].Value.ToString();
                        ks_work.CellAddressesMap.Add("Code", new ExellPropAddress(rowIndex, KS_CODE_COL, this.RegisterSheet));

                        ks_work.Name = registerSheet.Cells[rowIndex, KS_NAME_COL].Value;
                        ks_work.CellAddressesMap.Add("Name", new ExellPropAddress(rowIndex, KS_NAME_COL, this.RegisterSheet));
                        ks_work.CellAddressesMap.Add("UnitOfMeasurement", new ExellPropAddress(rowIndex, KS_MEASURE_COL, this.RegisterSheet));

                        ks_work.CellAddressesMap.Add("ProjectQuantity", new ExellPropAddress(rowIndex, KS_QUANTITY_COL, this.RegisterSheet));
                        ks_work.CellAddressesMap.Add("Quantity", new ExellPropAddress(rowIndex, KS_QUANTITY_FACT_COL, this.RegisterSheet));
                        ks_work.CellAddressesMap.Add("Laboriousness", new ExellPropAddress(rowIndex, KS_LABOURNESS_COL, this.RegisterSheet));

                        if (registerSheet.Cells[rowIndex, KS_MEASURE_COL].Value != null)
                        {
                            ks_work.UnitOfMeasurement = new UnitOfMeasurement(registerSheet.Cells[rowIndex, KS_MEASURE_COL].Value);
                            ks_work.UnitOfMeasurement.CellAddressesMap.Add("Name", new ExellPropAddress(rowIndex, KS_MEASURE_COL, this.RegisterSheet));
                            registerSheet.Range[registerSheet.Cells[rowIndex, KS_MEASURE_COL], registerSheet.Cells[rowIndex, KS_MEASURE_COL]].Interior.Color
                                = XlRgbColor.rgbWhite;
                        }
                        else
                            registerSheet.Range[registerSheet.Cells[rowIndex, KS_MEASURE_COL], registerSheet.Cells[rowIndex, KS_MEASURE_COL]].Interior.Color
                                = XlRgbColor.rgbRed;

                        if (registerSheet.Cells[rowIndex, KS_QUANTITY_COL].Value != null)
                        {
                            ks_work.ProjectQuantity = Decimal.Parse(registerSheet.Cells[rowIndex, KS_QUANTITY_COL].Value.ToString());
                            registerSheet.Range[registerSheet.Cells[rowIndex, KS_QUANTITY_COL], registerSheet.Cells[rowIndex, KS_QUANTITY_COL]].Interior.Color
                                = XlRgbColor.rgbWhite;
                        }
                        else
                            registerSheet.Range[registerSheet.Cells[rowIndex, KS_QUANTITY_COL], registerSheet.Cells[rowIndex, KS_QUANTITY_COL]].Interior.Color
                                = XlRgbColor.rgbRed;


                        if (registerSheet.Cells[rowIndex, KS_LABOURNESS_COL].Value != null)
                        {
                            ks_work.Laboriousness = Decimal.Parse(registerSheet.Cells[rowIndex, KS_LABOURNESS_COL].Value.ToString());
                            registerSheet.Range[registerSheet.Cells[rowIndex, KS_LABOURNESS_COL], registerSheet.Cells[rowIndex, KS_LABOURNESS_COL]].Interior.Color
                                = XlRgbColor.rgbWhite;
                        }
                        else
                            registerSheet.Range[registerSheet.Cells[rowIndex, KS_LABOURNESS_COL], registerSheet.Cells[rowIndex, KS_LABOURNESS_COL]].Interior.Color
                                = XlRgbColor.rgbRed;

                        ks_work.CellAddressesMap.Add("PreviousComplatedQuantity", new ExellPropAddress(rowIndex, KS_PC_QUANTITY_COL, this.RegisterSheet));

                        if (registerSheet.Cells[rowIndex, KS_PC_QUANTITY_COL].Value != null)
                            ks_work.PreviousComplatedQuantity = Decimal.Parse(registerSheet.Cells[rowIndex, KS_PC_QUANTITY_COL].Value.ToString());

                        //  this.Register(ks_work);
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
                            WorkReportCard report_card = new WorkReportCard();
                            DateTime end_date = DateTime.Parse(registerSheet.Cells[WORKS_END_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());
                            report_card.Number = registerSheet.Cells[rowIndex, WRC_NUMBER_COL].Value.ToString();
                            report_card.CellAddressesMap.Add("Number", new ExellPropAddress(rowIndex, WRC_NUMBER_COL, this.RegisterSheet));
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
                                    // workDay.CellAddressesMap.Add("Date", new ExellPropAddress(WRC_DATE_ROW, WRC_DATE_COL + date_index));
                                    workDay.Quantity = quantity;
                                    workDay.CellAddressesMap.Add("Quantity", new ExellPropAddress(rowIndex, WRC_DATE_COL + date_index, this.RegisterSheet));
                                    // this.Register(workDay);
                                    report_card.Add(workDay);
                                }
                                // this.Register(report_card);
                                date_index++;
                            }

                            KSWork ks_work = this.KSWorks.FirstOrDefault(w => w.Number == report_card.Number);
                            if (ks_work != null)
                                ks_work.ReportCard = report_card;
                            // this.Register(report_card);
                        }
                        rowIndex++;
                    }

            }
            public void LoadWorkerConsumptions()
            {
                Excel.Worksheet consumtionsSheet = this.WorkerConsumptionsSheet;
                int rowIndex = W_CONSUMPTIONS_FIRST_ROW_INDEX;
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
                        worker_consumption.CellAddressesMap.Add("Number", new ExellPropAddress(rowIndex, W_CONSUMPTIONS_NUMBER_COL, this.WorkerConsumptionsSheet));

                        worker_consumption.Name = consumtionsSheet.Cells[rowIndex, W_CONSUMPTIONS_NAME_COL].Value;
                        worker_consumption.CellAddressesMap.Add("Name", new ExellPropAddress(rowIndex, W_CONSUMPTIONS_NAME_COL, this.WorkerConsumptionsSheet));
                        worker_consumption.WorkersConsumptionReportCard.Clear();

                        if (!worker_consumption.CellAddressesMap.ContainsKey("WorkersConsumptionReportCard"))
                            worker_consumption.CellAddressesMap
                                          .Add("WorkersConsumptionReportCard", new ExellPropAddress(rowIndex, W_CONSUMPTIONS_NUMBER_COL, this.WorkerConsumptionsSheet));

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
                                    w_consumption_Day.CellAddressesMap.Add("Quantity", new ExellPropAddress(rowIndex, W_CONSUMPTIONS_FIRST_DATE_COL + date_index, this.WorkerConsumptionsSheet));
                                    worker_consumption.WorkersConsumptionReportCard.Add(w_consumption_Day);
                                    //         this.Register(w_consumption_Day);
                                }

                                //  this.Register(worker_consumption.WorkersConsumptionReportCard);
                                date_index++;
                            }

                        // this.Register(worker_consumption);
                    }
                    rowIndex++;
                }

            }

            /// <summary>
            /// Функция перезагружает все объекты из всех Worksheet в соотвествующие модели. 
            /// </summary>
            public void RealoadAllSheetsInModel()
            {
                foreach (MSGExellModel model in Children)
                    model.RealoadAllSheetsInModel();
                this.ContractCode = this.CommonSheet.Cells[CONTRACT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL].Value.ToString();
                this.ContructionObjectCode = this.CommonSheet.Cells[CONSTRUCTION_OBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL].Value.ToString();
                this.ConstructionSubObjectCode = this.CommonSheet.Cells[CONSTRUCTION_SUBOBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL].Value.ToString();

                //this.CellAddressesMap.Add("ContractCode", new ExellPropAddress<int, int, Worksheet>(CONTRACT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL, this.CommonSheet));
                //this.CellAddressesMap.Add("ContructionObjectCode", new ExellPropAddress<int, int, Worksheet>(CONSTRUCTION_OBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL, this.CommonSheet));
                //this.CellAddressesMap.Add("ConstructionSubObjectCode", new ExellPropAddress<int, int, Worksheet>(CONSTRUCTION_SUBOBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL, this.CommonSheet));

                //this.WorksStartDate=   DateTime.Parse(this.RegisterSheet.Cells[WORKS_START_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());
                //this.CellAddressesMap.Add("WorksStartDate", new ExellPropAddress<int, int, Worksheet>(WORKS_START_DATE_ROW, WORKS_END_DATE_COL, this.RegisterSheet));

                this.LoadWorksSections();
                //this.WorkersComposition.Clear();
                this.LoadMSGWorks();
                // this.LoadMSGWorkerCompositions();

                //this.MSGWorks[0].Name = "dsdsd";
                //this.LoadVOVRWorks();
                //this.LoadKSWorks();
                //this.LoadWorksReportCards();

                //this.LoadWorkerConsumptions();
            }

            /// <summary>
            /// Функиця пересчета трудоемкостей всех типов работ исходя из проставленных в трудоемкостей
            /// в работах типа КС-2
            /// </summary>
            public void CalcLabourness()
            {

                foreach (MSGWork msg_work in this.MSGWorks)
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
                                }
                                vovr_work.Laboriousness = common_ks_laboueness / vovr_work.ProjectQuantity;
                            }
                            common_vovr_laboueness += vovr_work.ProjectQuantity * vovr_work.Laboriousness;
                        }
                        msg_work.Laboriousness = common_vovr_laboueness / msg_work.ProjectQuantity;
                    }
                }
            }
            /// <summary>
            /// Функцич подсчета объемов выполненных работ 
            /// </summary>
            public void CalcQuantity()
            {

                foreach (MSGWork msg_work in this.MSGWorks)
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
                            ks_work.Quantity = 0;

                            if (this.Owner != null && ks_work.ReportCard != null)
                            {
                                ks_work.Quantity = ks_work.ReportCard.Quantity + ks_work.PreviousComplatedQuantity;
                            }
                            else
                            {
                                ks_work.PreviousComplatedQuantity = 0;
                                if (ks_work.ReportCard == null)
                                {
                                    // ks_work.CellAddressesMap.Add("", new ExellPropAddress<int, int, Worksheet>());
                                    ks_work.ReportCard = new WorkReportCard();
                                    ks_work.ReportCard.CellAddressesMap.Add("Number",
                                        new ExellPropAddress(ks_work.CellAddressesMap["Number"].Row, WRC_NUMBER_COL, this.RegisterSheet));
                                    //this.RegisterSheet.Cells[ks_work.CellAddressesMap["Number"].Item1,
                                    //    WRC_NUMBER_COL] = ks_work.Number;
                                    ks_work.ReportCard.Number = ks_work.Number;
                                    //ks_work.ReportCard.CellAddressesMap.Add("ReportCard",
                                    //  new ExellPropAddress<int, int, Excel.Worksheet>(ks_work.CellAddressesMap["Number"].Item1, WRC_NUMBER_COL, this.RegisterSheet));
                                    //  this.Register(ks_work.ReportCard);
                                }
                                else
                                    ks_work.ReportCard.Clear();

                                foreach (MSGExellModel model in this.Children)
                                {

                                    KSWork child_ks_work = model.KSWorks.FirstOrDefault(w => w.Number == ks_work.Number);


                                    if (child_ks_work != null && child_ks_work.ReportCard != null)
                                    {

                                        foreach (WorkDay child_w_day in child_ks_work.ReportCard)
                                        {
                                            WorkDay curent_w_day = ks_work.ReportCard.FirstOrDefault(wd => wd.Date == child_w_day.Date);
                                            if (curent_w_day != null)
                                            {
                                                curent_w_day.Quantity += child_w_day.Quantity;
                                                curent_w_day.LaborСosts = curent_w_day.Quantity * ks_work.Laboriousness;
                                            }
                                            else
                                            {
                                                curent_w_day = new WorkDay();
                                                // this.Register(curent_w_day);
                                                curent_w_day.Date = child_w_day.Date;

                                                curent_w_day.Quantity = child_w_day.Quantity;
                                                curent_w_day.LaborСosts = child_w_day.Quantity * ks_work.Laboriousness;
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

                                                    int curent_wrc_row = ks_work.ReportCard.CellAddressesMap["Number"].Row;

                                                    curent_w_day.CellAddressesMap.Add(map_item.Key,
                                                        new ExellPropAddress(curent_wrc_row, WRC_DATE_COL + date_index, this.RegisterSheet));
                                                    curent_w_day.Quantity = curent_w_day.Quantity;
                                                    //  this.Register(curent_w_day);
                                                }
                                                ks_work.ReportCard.Add(curent_w_day);
                                            }
                                        }
                                        ks_work.PreviousComplatedQuantity += child_ks_work.PreviousComplatedQuantity;
                                    }


                                }
                                //ks_work.Quantity = ks_work.ReportCard.Quantity;
                                ks_work.Quantity = ks_work.ReportCard.Quantity + ks_work.PreviousComplatedQuantity;
                            }

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
                    var msg_work_all_ksWorks = this.KSWorks.Where(w => w.Number.StartsWith(msg_work.Number + "."));
                    foreach (KSWork ks_work in msg_work_all_ksWorks)
                    {
                        if (ks_work.ReportCard != null)
                        {
                            foreach (WorkDay ks_w_day in ks_work.ReportCard)
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
                //  this.UpdateWorksheetCommonPart();
                this.RealoadAllSheetsInModel();
                this.CalcLabourness();
                this.CalcQuantity();
                this.LoadWorksReportCards();
                this.CalcWorkerConsumptions();
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
            public void UpdateWorksheetCommonPart()
            {
                if (this.Owner != null)
                {
                    this.WorksSections = (ExcelNotifyChangedCollection<WorksSection>)this.Owner.WorksSections.Clone();
                    //this.WorksSections.CellAddressesMap.SetWorksheet(this.RegisterSheet);
                    //    ClearWorksheetCommonPart();
                    foreach (WorksSection w_section in this.WorksSections)
                    {
                        //w_section.MSGWorks.CellAddressesMap.SetWorksheet(this.RegisterSheet);
                        //foreach (MSGWork msg_work in w_section.MSGWorks)
                        //{
                        //    msg_work.WorkSchedules.CellAddressesMap.SetWorksheet(this.RegisterSheet);
                        //    msg_work.UnitOfMeasurement.CellAddressesMap.SetWorksheet(this.RegisterSheet);
                        //}

                        this.UpdateExellBindableObject(w_section);
                    }
                    this.ResetCalculatesFields();
                }
                else
                    ;// this.ClearWorksheetDaysPart();
            }
            public void GetWorksFromOwner()
            {
                if (this.Owner != null)
                {

                }
            }
            /// <summary>
            /// Функция обновляет документальное представление объетка (рукурсивно проходит по всем объектам 
            /// реализующим интерфейс IExcelBindableBase). 
            /// </summary>
            /// <param name="obj">Связанный с докуметом Worksheet объект рализующий IExcelBindableBase </param>
            private void UpdateExellBindableObject(IExcelBindableBase obj)
            {
                //foreach (KeyValuePair<RalateRecord, Tuple<string, ExellPropAddress>> register_item in this.RegistedObjects)
                //{
                //    if (register_item.Value.Item2.Worksheet.Name.Contains("Ведомость"))
                //        register_item.Value.Item2.Worksheet = RegisterSheet;
                //    if (register_item.Value.Item2.Worksheet.Name.Contains("Люди"))
                //        register_item.Value.Item2.Worksheet = this.WorkerConsumptionsSheet;
                //}

                //var register_items = this.RegistedObjects.Where(ri => ri.Key.Entity.Id == obj.Id);
                //foreach (KeyValuePair<RalateRecord, Tuple<string, ExellPropAddress>> kvp in register_items)
                //{
                //    kvp.Value.Item2.Cell.Value = "ff";
                //}
                var prop_infoes = obj.GetType().GetProperties().Where(pr => pr.GetIndexParameters().Length == 0);
                if (obj is MSGWork)
                    ;
                foreach (PropertyInfo property_info in prop_infoes)
                {
                    var prop_val = property_info.GetValue(obj);

                    if (prop_val is IExcelBindableBase excel_bindable_val)
                    {

                        this.UpdateExellBindableObject(excel_bindable_val);
                    }
                }

            }
            //private void UpdateExellBindableObject(IExcelBindableBase obj)
            //{
            //    var prop_infoes = obj.GetType().GetProperties().Where(pr => pr.GetIndexParameters().Length == 0);
            //    if (obj is MSGWork)
            //        ;
            //    foreach (PropertyInfo property_info in prop_infoes)
            //    {
            //        var prop_val = property_info.GetValue(obj);
            //        if (prop_val is UnitOfMeasurement)
            //            ;
            //        if (obj.CellAddressesMap.ContainsKey(property_info.Name) && obj.CellAddressesMap[property_info.Name].Column <= WRC_NUMBER_COL)
            //        {
            //            var obj_map_item = obj.CellAddressesMap[property_info.Name];
            //            if (property_info.PropertyType.FullName.Contains("System."))
            //            {
            //                if (prop_val is DateTime date_val)
            //                    //obj_map_item.Worksheet.Cells[obj_map_item.Row,obj_map_item.Column] = date_val.ToString("d");
            //                    obj_map_item.Cell.Value = date_val.ToString("d");
            //                else
            //                    obj_map_item.Cell.Value = prop_val.ToString();

            //            }
            //            else if (prop_val is IExcelBindableBase exel_bindable_val)
            //            {
            //                this.UpdateExellBindableObject(exel_bindable_val);
            //            }
            //            else if (prop_val is INameable nameable_val)
            //            {
            //                obj_map_item.Cell.Value = nameable_val.Name;
            //            }
            //        }
            //        if (prop_val is IList list_prop_val)
            //        {

            //            foreach (object element in list_prop_val)
            //                if (element is IExcelBindableBase excel_bindable_obj)
            //                    this.UpdateExellBindableObject(excel_bindable_obj);

            //        }

            //    }
            //}
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
            public void ClearWorksheetCommonPart()
            {
                MSGExellModel mode;
                if (this.Owner != null)
                    mode = this.Owner;

                if (this.Owner != null && this.Owner.KSWorks.Count > 0)
                {
                    Excel.Range common_area_range = this.RegisterSheet.Range[this.RegisterSheet.Cells[FIRST_ROW_INDEX, WSEC_NUMBER_COL],
                        this.RegisterSheet.Cells[this.Owner.KSWorks[this.Owner.KSWorks.Count - 1].CellAddressesMap["Laboriousness"].Row, WRC_NUMBER_COL - 1]];
                    common_area_range.ClearContents();
                    common_area_range.Interior.ColorIndex = 0;
                }
                else
                {
                    Excel.Range common_area_range = this.RegisterSheet.Range[this.RegisterSheet.Cells[FIRST_ROW_INDEX, WSEC_NUMBER_COL], this.RegisterSheet.Cells[
                                           this.KSWorks[this.KSWorks.Count - 1].CellAddressesMap["Laboriousness"].Row, WRC_NUMBER_COL - 1]];
                    common_area_range.Interior.ColorIndex = 0;
                }


            }
        }
    }
