using ExellAddInsLib.MSG.Section;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExellAddInsLib.MSG
{
    public class MSGExellModel
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
        public const int MSG_WORKERS_NUMBER_COL = MSG_NUMBER_COL + 8;


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
        public ObservableCollection<WorksSection> WorksSections { get; private set; } = new ObservableCollection<WorksSection>();
        /// <summary>
        /// Коллекция с работами типа МСГ модели
        /// </summary>
        public ObservableCollection<MSGWork> MSGWorks { get; private set; } = new ObservableCollection<MSGWork>();
        /// <summary>
        /// Коллекция с работами типа ВОВР модели
        /// </summary>
        public ObservableCollection<VOVRWork> VOVRWorks { get; private set; } = new ObservableCollection<VOVRWork>();
        /// <summary>
        /// Коллекция с работами типа КС-2 модели
        /// </summary>
        public ObservableCollection<KSWork> KSWorks { get; private set; } = new ObservableCollection<KSWork>();
        /// <summary>
        /// Коллекция с единицами измерения модели
        /// </summary>
        public ObservableCollection<UnitOfMeasurement> UnitOfMeasurements { get; set; } = new ObservableCollection<UnitOfMeasurement>();

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
        /// <summary>
        /// Функция для регистрации объекта реализующего интрефейс INotifyPropertyChanged 
        /// для обработки событий изменения полей объета и соотвествующего изменения связанной с 
        /// с этим полем ячейки в документе Worksheet
        /// </summary>
        /// <param name="work"></param>
        public void Register(object obj)
        {
            if (obj is INotifyPropertyChanged notified_object)
                notified_object.PropertyChanged += OnPropertyChange;
            switch (obj.GetType().Name)
            {

                case nameof(WorksSection):
                    {

                        WorksSection w_section = (WorksSection)obj;
                        if (!this.WorksSections.Contains(w_section))
                            this.WorksSections.Add(w_section);
                        break;
                    }

                case nameof(MSGWork):
                    {

                        MSGWork msg_work = (MSGWork)obj;
                        if (!this.MSGWorks.Contains(msg_work))
                            this.MSGWorks.Add(msg_work);

                        WorksSection w_section = this.WorksSections.Where(ws => ws.Number.StartsWith(msg_work.Number.Remove(msg_work.Number.LastIndexOf(".")))).FirstOrDefault();
                        if (w_section != null)
                        {
                            w_section.MSGWorks.Add(msg_work);
                        }
                        break;
                    }
                case nameof(VOVRWork):
                    {
                        VOVRWork vovr_work = (VOVRWork)obj;
                        if (!this.VOVRWorks.Contains(vovr_work))
                            this.VOVRWorks.Add(vovr_work);

                        MSGWork msg_work = this.MSGWorks.Where(w => w.Number.StartsWith(vovr_work.Number.Remove(vovr_work.Number.LastIndexOf(".")))).FirstOrDefault();
                        if (msg_work != null)
                        {
                            msg_work.VOVRWorks.Add(vovr_work);
                        }

                        break;
                    }
                case nameof(KSWork):
                    {
                        KSWork ks_work = (KSWork)obj;
                        if (!this.KSWorks.Contains(ks_work))
                            this.KSWorks.Add(ks_work);

                        VOVRWork vovr_work = VOVRWorks.Where(w => w.Number.StartsWith(ks_work.Number.Remove(ks_work.Number.LastIndexOf(".")))).FirstOrDefault();
                        if (vovr_work != null)
                            vovr_work.KSWorks.Add(ks_work);

                        break;
                    }

                case nameof(WorkReportCard):
                    {
                        WorkReportCard report_card = (WorkReportCard)obj;

                        KSWork ks_work = KSWorks.Where(w => w.Number == report_card.Number).FirstOrDefault();
                        if (ks_work != null && report_card.Count > 0)
                            ks_work.ReportCard = report_card;
                        else if (ks_work == null)
                        {
                            ks_work = new KSWork();
                            ks_work.Number = report_card.Number;
                            ks_work.ReportCard = report_card;
                        }

                        break;
                    }

            }
        }
        private void OnPropertyChange(object sender, PropertyChangedEventArgs e)
        {
            if (sender is IExcelBindableBase bindable_object)
            {
                if (bindable_object.CellAddressesMap.ContainsKey(e.PropertyName))
                {
                    RegisterSheet.Cells[bindable_object.CellAddressesMap[e.PropertyName].Item1,
                   bindable_object.CellAddressesMap[e.PropertyName].Item2] = sender.GetType().GetProperty(e.PropertyName).GetValue(sender).ToString();
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
            this.MSGWorks.Clear();

            while (null_str_count < 100)
            {
                if (registerSheet.Cells[rowIndex, WSEC_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    WorksSection w_section = new WorksSection();

                    w_section.Number = registerSheet.Cells[rowIndex, WSEC_NUMBER_COL].Value.ToString();

                    w_section.CellAddressesMap.Add("Number", Tuple.Create(rowIndex, WSEC_NUMBER_COL));

                    w_section.Name = registerSheet.Cells[rowIndex, WSEC_NAME_COL].Value;
                    w_section.CellAddressesMap.Add("Name", Tuple.Create(rowIndex, WSEC_NAME_COL));

                    this.Register(w_section);
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

                    msg_work.CellAddressesMap.Add("Number", Tuple.Create(rowIndex, MSG_NUMBER_COL));

                    msg_work.Name = registerSheet.Cells[rowIndex, MSG_NAME_COL].Value;
                    msg_work.CellAddressesMap.Add("Name", Tuple.Create(rowIndex, MSG_NAME_COL));
                    msg_work.CellAddressesMap.Add("UnitOfMeasurement", Tuple.Create(rowIndex, MSG_MEASURE_COL));
                    msg_work.CellAddressesMap.Add("ProjectQuantity", Tuple.Create(rowIndex, MSG_QUANTITY_COL));
                    msg_work.CellAddressesMap.Add("Quantity", Tuple.Create(rowIndex, MSG_QUANTITY_FACT_COL));
                    msg_work.CellAddressesMap.Add("Laboriousness", Tuple.Create(rowIndex, MSG_LABOURNESS_COL));

                    if (registerSheet.Cells[rowIndex, MSG_MEASURE_COL].Value != null)
                    {
                        string un_name = registerSheet.Cells[rowIndex, MSG_MEASURE_COL].Value.ToString();
                        UnitOfMeasurement unitOfMeasurement = UnitOfMeasurements.FirstOrDefault(um => um.Name == un_name);
                        if (unitOfMeasurement != null)
                        {
                            msg_work.UnitOfMeasurement = unitOfMeasurement;
                            registerSheet.Range[registerSheet.Cells[rowIndex, MSG_MEASURE_COL], registerSheet.Cells[rowIndex, MSG_MEASURE_COL]].Interior.Color
                                = XlRgbColor.rgbWhite;
                        }
                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_MEASURE_COL], registerSheet.Cells[rowIndex, MSG_MEASURE_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, MSG_QUANTITY_COL].Value != null)
                    {
                        msg_work.ProjectQuantity = Decimal.Parse(registerSheet.Cells[rowIndex, MSG_QUANTITY_COL].Value.ToString());
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_QUANTITY_COL], registerSheet.Cells[rowIndex, MSG_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;

                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_QUANTITY_COL], registerSheet.Cells[rowIndex, MSG_QUANTITY_COL]].Interior.Color
                            = XlRgbColor.rgbRed;

                    if (registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value != null)
                    {
                        var fdf = registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value.ToString();
                        decimal res;
                        Decimal.TryParse(registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value.ToString(), out res);
                        msg_work.Laboriousness = res;//Decimal.Parse(registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL].Value.ToString());
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL], registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbWhite;
                    }
                    else
                        registerSheet.Range[registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL], registerSheet.Cells[rowIndex, MSG_LABOURNESS_COL]].Interior.Color
                            = XlRgbColor.rgbRed;



                    DateTime start_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_START_DATE_COL].Value.ToString());
                    DateTime end_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_END_DATE_COL].Value.ToString());
                    WorkScheduleChunk work_sh_chunk = new WorkScheduleChunk(start_time, end_time);
                    work_sh_chunk.CellAddressesMap.Add("StartTime", Tuple.Create(rowIndex, MSG_START_DATE_COL));
                    work_sh_chunk.CellAddressesMap.Add("EndTime", Tuple.Create(rowIndex, MSG_END_DATE_COL));
                    msg_work.WorkSchedules.Add(work_sh_chunk);
                    this.Register(work_sh_chunk);

                    while (registerSheet.Cells[rowIndex + 1, MSG_NUMBER_COL].Value == null
                                 && registerSheet.Cells[rowIndex + 1, MSG_START_DATE_COL].Value != null)
                    {
                        rowIndex++;
                        start_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_START_DATE_COL].Value.ToString());
                        end_time = DateTime.Parse(registerSheet.Cells[rowIndex, MSG_END_DATE_COL].Value.ToString());
                        var fd = registerSheet.Cells[rowIndex, MSG_WORKERS_NUMBER_COL].Value;
                        int workers_number = 0;
                        if (registerSheet.Cells[rowIndex, MSG_WORKERS_NUMBER_COL].Value != null)
                        {
                            int.TryParse(registerSheet.Cells[rowIndex, MSG_WORKERS_NUMBER_COL].Value.ToString(), out workers_number);
                            registerSheet.Range[registerSheet.Cells[rowIndex, MSG_WORKERS_NUMBER_COL], registerSheet.Cells[rowIndex, MSG_WORKERS_NUMBER_COL]].Interior.Color
                                                             = XlRgbColor.rgbWhite;
                        }
                        else
                            registerSheet.Range[registerSheet.Cells[rowIndex, MSG_WORKERS_NUMBER_COL], registerSheet.Cells[rowIndex, MSG_WORKERS_NUMBER_COL]].Interior.Color
                                                              = XlRgbColor.rgbRed;

                        WorkScheduleChunk extra_work_sh_chunk = new WorkScheduleChunk(start_time, end_time);
                        extra_work_sh_chunk.WorkesNumber = workers_number;
                        extra_work_sh_chunk.CellAddressesMap.Add("StartTime", Tuple.Create(rowIndex, MSG_START_DATE_COL));
                        extra_work_sh_chunk.CellAddressesMap.Add("EndTime", Tuple.Create(rowIndex, MSG_END_DATE_COL));
                        extra_work_sh_chunk.CellAddressesMap.Add("WorkerNumber", Tuple.Create(rowIndex, MSG_WORKERS_NUMBER_COL));


                        msg_work.WorkSchedules.Add(extra_work_sh_chunk);
                        this.Register(extra_work_sh_chunk);
                    }
                    this.Register(msg_work);
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
            null_str_count = 0;

            while (null_str_count < 100)
            {
                if (registerSheet.Cells[rowIndex, VOVR_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    VOVRWork vovr_work = new VOVRWork();

                    vovr_work.Number = registerSheet.Cells[rowIndex, VOVR_NUMBER_COL].Value.ToString();
                    vovr_work.CellAddressesMap.Add("Number", Tuple.Create(rowIndex, VOVR_NUMBER_COL));

                    vovr_work.Name = registerSheet.Cells[rowIndex, VOVR_NAME_COL].Value.ToString();
                    vovr_work.CellAddressesMap.Add("Name", Tuple.Create(rowIndex, VOVR_NAME_COL));
                    vovr_work.CellAddressesMap.Add("UnitOfMeasurement", Tuple.Create(rowIndex, VOVR_MEASURE_COL));
                    vovr_work.CellAddressesMap.Add("ProjectQuantity", Tuple.Create(rowIndex, VOVR_QUANTITY_COL));
                    vovr_work.CellAddressesMap.Add("Quantity", Tuple.Create(rowIndex, VOVR_QUANTITY_FACT_COL));
                    vovr_work.CellAddressesMap.Add("Laboriousness", Tuple.Create(rowIndex, VOVR_LABOURNESS_COL));

                    if (registerSheet.Cells[rowIndex, VOVR_MEASURE_COL].Value != null)
                    {
                        vovr_work.UnitOfMeasurement = new UnitOfMeasurement(registerSheet.Cells[rowIndex, VOVR_MEASURE_COL].Value);
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

                    this.Register(vovr_work);

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
            null_str_count = 0;
            while (null_str_count < 100)
            {
                if (registerSheet.Cells[rowIndex, KS_NUMBER_COL].Value == null) null_str_count++;
                else
                {
                    null_str_count = 0;
                    KSWork ks_work = new KSWork();

                    ks_work.Number = registerSheet.Cells[rowIndex, KS_NUMBER_COL].Value.ToString();
                    ks_work.CellAddressesMap.Add("Number", Tuple.Create(rowIndex, KS_NUMBER_COL));

                    ks_work.Code = registerSheet.Cells[rowIndex, KS_CODE_COL].Value.ToString();
                    ks_work.CellAddressesMap.Add("Code", Tuple.Create(rowIndex, KS_CODE_COL));

                    ks_work.Name = registerSheet.Cells[rowIndex, KS_NAME_COL].Value;
                    ks_work.CellAddressesMap.Add("Name", Tuple.Create(rowIndex, KS_NAME_COL));

                    ks_work.CellAddressesMap.Add("UnitOfMeasurement", Tuple.Create(rowIndex, KS_MEASURE_COL));
                    ks_work.CellAddressesMap.Add("ProjectQuantity", Tuple.Create(rowIndex, KS_QUANTITY_COL));
                    ks_work.CellAddressesMap.Add("Quantity", Tuple.Create(rowIndex, KS_QUANTITY_FACT_COL));
                    ks_work.CellAddressesMap.Add("Laboriousness", Tuple.Create(rowIndex, KS_LABOURNESS_COL));
                 
                    if (registerSheet.Cells[rowIndex, KS_MEASURE_COL].Value != null)
                    {
                        ks_work.UnitOfMeasurement = new UnitOfMeasurement(registerSheet.Cells[rowIndex, KS_MEASURE_COL].Value);
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
                  
                    ks_work.CellAddressesMap.Add("PreviousComplatedQuantity", Tuple.Create(rowIndex, KS_PC_QUANTITY_COL));

                    if (registerSheet.Cells[rowIndex, KS_PC_QUANTITY_COL].Value != null)
                        ks_work.PreviousComplatedQuantity = Decimal.Parse(registerSheet.Cells[rowIndex, KS_PC_QUANTITY_COL].Value.ToString());

                    this.Register(ks_work);
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
                        report_card.CellAddressesMap.Add("Number", Tuple.Create(rowIndex, WRC_NUMBER_COL));
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
                                // workDay.CellAddressesMap.Add("Date", Tuple.Create(WRC_DATE_ROW, WRC_DATE_COL + date_index));
                                workDay.Quantity = quantity;
                                workDay.CellAddressesMap.Add("Quantity", Tuple.Create(rowIndex, WRC_DATE_COL + date_index));
                                this.Register(workDay);
                                report_card.Add(workDay);
                            }
                            this.Register(report_card);
                            date_index++;
                        }

                        KSWork ks_work = this.KSWorks.FirstOrDefault(w => w.Number == report_card.Number);
                        if (ks_work != null)
                            ks_work.ReportCard = report_card;
                        this.Register(report_card);
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
            this.WorksSections.Clear();
            this.MSGWorks.Clear();
            this.VOVRWorks.Clear();
            this.KSWorks.Clear();
            this.ContractCode = this.CommonSheet.Cells[CONTRACT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL].Value.ToString();
            this.ContructionObjectCode = this.CommonSheet.Cells[CONSTRUCTION_OBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL].Value.ToString();
            this.ConstructionSubObjectCode = this.CommonSheet.Cells[CONSTRUCTION_SUBOBJECT_CODE_ROW, COMMON_PARAMETRS_VALUE_COL].Value.ToString();

            this.LoadWorksSections();
            this.LoadMSGWorks();
            this.LoadVOVRWorks();
            this.LoadKSWorks();
            this.LoadWorksReportCards();
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
                            ks_work.Quantity = ks_work.ReportCard.Quantity+ ks_work.PreviousComplatedQuantity;
                        }
                        else
                        {
                            ks_work.PreviousComplatedQuantity = 0;
                            if (ks_work.ReportCard == null)
                            {
                                ks_work.ReportCard = new WorkReportCard();
                                ks_work.ReportCard.Number = ks_work.Number;
                                this.RegisterSheet.Cells[ks_work.CellAddressesMap["Number"].Item1,
                                    WRC_NUMBER_COL] = ks_work.Number;
                                ks_work.ReportCard.CellAddressesMap.Add("ReportCard",
                                    new Tuple<int, int>(ks_work.CellAddressesMap["Number"].Item1, WRC_NUMBER_COL));
                                this.Register(ks_work.ReportCard);
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
                                            this.Register(curent_w_day);
                                            curent_w_day.Date = child_w_day.Date;
                                            curent_w_day.Quantity = child_w_day.Quantity;
                                            curent_w_day.LaborСosts = child_w_day.Quantity * ks_work.Laboriousness;
                                            DateTime end_date = DateTime.Parse(this.RegisterSheet.Cells[WORKS_END_DATE_ROW, WORKS_END_DATE_COL].Value.ToString());

                                            foreach (KeyValuePair<string, Tuple<int, int>> map_item in child_w_day.CellAddressesMap)
                                            {
                                                int date_index = 0;
                                                while (this.RegisterSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value != null &&
                                                  DateTime.Parse(this.RegisterSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value.ToString()) < end_date)
                                                {
                                                    if (DateTime.Parse(this.RegisterSheet.Cells[WRC_DATE_ROW, WRC_DATE_COL + date_index].Value.ToString()) == curent_w_day.Date)
                                                        break;
                                                    date_index++;
                                                }
                                                int curent_wrc_row = ks_work.ReportCard.CellAddressesMap["ReportCard"].Item1;

                                                curent_w_day.CellAddressesMap.Add(map_item.Key, new Tuple<int, int>(curent_wrc_row, WRC_DATE_COL + date_index));

                                                this.RegisterSheet.Cells[curent_wrc_row, WRC_DATE_COL + date_index] =
                                                    curent_w_day.Quantity.ToString();
                                                this.Register(curent_w_day);
                                            }
                                            ks_work.ReportCard.Add(curent_w_day);
                                        }
                                    }
                                }
                                ks_work.PreviousComplatedQuantity += child_ks_work.PreviousComplatedQuantity;
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
                    common_vovr_previos_complate_labour_quantity+= vovr_work.PreviousComplatedQuantity * vovr_work.Laboriousness;
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
        /// <summary>
        /// Вычисление всех вычисляемых величин внутри модели и всех его дочерних моделей.
        /// </summary>
        public void CalcAll()
        {
            this.UpdateWorksheetCommonPart();
            this.RealoadAllSheetsInModel();
            this.CalcLabourness();
            this.CalcQuantity();
            this.LoadWorksReportCards();
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

                ClearWorksheetCommonPart();
                foreach (WorksSection w_section in this.Owner.WorksSections)
                {
                    this.UpdateExellBindableObject(w_section);
                }
                this.ResetCalculatesFields();
            }
            else
                this.ClearWorksheetDaysPart();
        }
        /// <summary>
        /// Функция обновляет документальное представление объетка (рукурсивно проходит по всем объектам 
        /// реализующим интерфейс IExcelBindableBase). 
        /// </summary>
        /// <param name="obj">Связанный с докуметом Worksheet объект рализующий IExcelBindableBase </param>
        private void UpdateExellBindableObject(IExcelBindableBase obj)
        {
            var prop_infoes = obj.GetType().GetProperties().Where(pr => pr.GetIndexParameters().Length == 0);
            foreach (PropertyInfo property_info in prop_infoes)
            {
                var prop_val = property_info.GetValue(obj);

                if (obj.CellAddressesMap.ContainsKey(property_info.Name) && obj.CellAddressesMap[property_info.Name].Item2 <= WRC_NUMBER_COL)
                {
                    if (property_info.PropertyType.FullName.Contains("System."))
                    {
                        if (prop_val is DateTime date_val)
                            this.RegisterSheet.Cells[obj.CellAddressesMap[property_info.Name].Item1,
                                                obj.CellAddressesMap[property_info.Name].Item2]
                                                = date_val.ToString("d");
                        else
                            this.RegisterSheet.Cells[obj.CellAddressesMap[property_info.Name].Item1,
                                               obj.CellAddressesMap[property_info.Name].Item2]
                                               = prop_val.ToString();
                    }
                    else if (prop_val is IExcelBindableBase exel_bindable_val)
                    {
                        this.UpdateExellBindableObject(exel_bindable_val);
                    }
                    else if (prop_val is INameable nameable_val)
                    {
                        this.RegisterSheet.Cells[obj.CellAddressesMap[property_info.Name].Item1,
                                                obj.CellAddressesMap[property_info.Name].Item2]
                                                = nameable_val.Name;
                    }
                }
                if (prop_val is IList list_prop_val)
                {

                    foreach (object element in list_prop_val)
                        if (element is IExcelBindableBase excel_bindable_obj)
                            this.UpdateExellBindableObject(excel_bindable_obj);

                }

            }
        }
        /// <summary>
        /// Фунция очищает календарную часть ведомости (очищает все записи выполненных работ)
        /// </summary>
        public void ClearWorksheetDaysPart()
        {
            Excel.Range last_cell = this.RegisterSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            Excel.Range common_area_range = this.RegisterSheet.Range[this.RegisterSheet.Cells[FIRST_ROW_INDEX, WRC_DATE_COL], this.RegisterSheet.Cells[10000, 10000]];
            common_area_range.ClearContents();
        }
        /// <summary>
        /// Функия очищает левую часть вдомости с МСГ, ВОВР и КС-2.
        /// </summary>
        public void ClearWorksheetCommonPart()
        {
            Excel.Range last_cell = this.RegisterSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            Excel.Range common_area_range = this.RegisterSheet.Range[this.RegisterSheet.Cells[FIRST_ROW_INDEX, WSEC_NUMBER_COL], this.RegisterSheet.Cells[10000, WRC_NUMBER_COL - 1]];
            common_area_range.ClearContents();
        }
    }
}
