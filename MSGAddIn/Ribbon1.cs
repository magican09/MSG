using ExellAddInsLib.MSG;
using ExellAddInsLib.MSG.Section;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace MSGAddIn
{
    public partial class Ribbon1
    {


        private const int POST_NUMBER_COL = 1;
        private const int POST_NAME_COL = 2;

        private const int EMPLOYER_NUMBER_COL = 1;
        private const int EMPLOYER_NAME_COL = 2;
        private const int EMPLOYER_POSTNAME_COL = 3;

        private const int UM_NUMBER_COL = 1;
        private const int UM_NAME_COL = 2;

        bool first_start_flag = true;

        MSGExellModel CurrentMSGExellModel;
        MSGExellModel CommonMSGExellModel;
        ObservableCollection<MSGExellModel> MSGExellModels = new ObservableCollection<MSGExellModel>();

        ObservableCollection<Employer> Employers { get; set; } = new ObservableCollection<Employer>();
        ExcelNotifyChangedCollection<UnitOfMeasurement> UnitOfMeasurements = new ExcelNotifyChangedCollection<UnitOfMeasurement>();

        Excel._Workbook CurrentWorkbook;
        Excel._Workbook MSGTemplateWorkbook;

        Excel.Worksheet EmployersWorksheet;
        Excel.Worksheet PostsWorksheet;
        Excel.Worksheet UnitMeasurementsWorksheet;
        Excel.Worksheet CommonWorksheet;
        Excel.Worksheet CommonMSGWorksheet;
        Excel.Worksheet CommonWorkConsumptionsWorksheet;
        Excel.Worksheet TemplateMSGWorksheet;

        ObservableCollection<Excel.Worksheet> EmployerMSGWorksheets = new ObservableCollection<Worksheet>();
        ObservableCollection<Excel.Worksheet> EmployerWorkConsumptionsWorksheets = new ObservableCollection<Worksheet>();

        Employer SelectedEmloeyer;
        private bool InMSGWorkbook = false;
        private void OnActiveWorkbookChanged(Workbook last_wbk, Workbook new_wbk)
        {
            CommonMSGWorksheet = new_wbk.Worksheets.OfType<Excel.Worksheet>().FirstOrDefault(w => w.Name == "Ведомость_общая");
            CommonWorkConsumptionsWorksheet = new_wbk.Worksheets.OfType<Excel.Worksheet>().FirstOrDefault(w => w.Name == "Люди_общая");
            CommonWorksheet = new_wbk.Worksheets.OfType<Excel.Worksheet>().FirstOrDefault(w => w.Name == "Начальная");
            UnitMeasurementsWorksheet = new_wbk.Worksheets.OfType<Excel.Worksheet>().FirstOrDefault(w => w.Name == "Ед_изм");
            PostsWorksheet = new_wbk.Worksheets.OfType<Excel.Worksheet>().FirstOrDefault(w => w.Name == "Должности");
            EmployersWorksheet = new_wbk.Worksheets.OfType<Excel.Worksheet>().FirstOrDefault(w => w.Name == "Ответственные");
            if (CommonMSGWorksheet != null && CommonWorksheet != null && UnitMeasurementsWorksheet != null && PostsWorksheet != null && EmployersWorksheet != null)
            {
                //    LoadMSG_File();
                InMSGWorkbook = true;
                groupFileLaod.Visible = true;
                groupMSGCommon.Visible = true;
                grpInChargePersons.Visible = true;

                if (CurrentMSGExellModel != null && CurrentMSGExellModel.ContractCode ==
                    CommonWorksheet.Cells[MSGExellModel.CONTRACT_CODE_ROW, MSGExellModel.COMMON_PARAMETRS_VALUE_COL].Value.ToString())
                {
                    this.ShowWorksheet(CurrentMSGExellModel.RegisterSheet);
                    this.AjastBtnsState();
                }

            }
            else
            {
                InMSGWorkbook = false;
                groupFileLaod.Visible = false;
                groupMSGCommon.Visible = false;
                grpInChargePersons.Visible = false;
                this.SetBtnsState(false);
            }
        }
        private void OnActiveWorksheetChanged(Excel.Worksheet last_wsh, Excel.Worksheet new_wsh)
        {
            if (InMSGWorkbook)
            {
                this.ReloadEmployersList();
                this.ReloadMeasurementsList();

            }

            // this.ReloadAllModels();



        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.OnActiveWorksheetChanged += OnActiveWorksheetChanged;
            Globals.ThisAddIn.OnActiveWorkbookChanged += OnActiveWorkbookChanged;

        }
        private void SetBtnsState(bool state)
        {
            btnCalcLabournes.Enabled = state;
            btnCalcAll.Enabled = state;
            btnReloadWorksheets.Enabled = state;
            btnChangeCommonMSG.Enabled = state;
            btnLoadTeplateFile.Enabled = state;
            buttonCalc.Enabled = state;
        }
        private void AjastBtnsState()
        {
            this.SetBtnsState(true);
        }

        private void btnLoadMSGFile_Click(object sender, RibbonControlEventArgs e)
        {
            LoadMSG_File();
            CurrentMSGExellModel.SetStyleFormats();
        }
        private void LoadMSG_File()
        {
            CurrentWorkbook = Globals.ThisAddIn.CurrentActivWorkbook;
            EmployersWorksheet = CurrentWorkbook.Worksheets["Ответственные"];
            PostsWorksheet = CurrentWorkbook.Worksheets["Должности"];
            UnitMeasurementsWorksheet = CurrentWorkbook.Worksheets["Ед_изм"];
            CommonWorksheet = CurrentWorkbook.Worksheets["Начальная"];
            CommonMSGWorksheet = CurrentWorkbook.Worksheets["Ведомость_общая"];
            CommonWorkConsumptionsWorksheet = CurrentWorkbook.Worksheets["Люди_общая"];

            EmployerMSGWorksheets = new ObservableCollection<Excel.Worksheet>();

            this.ReloadEmployersList();
            this.ReloadMeasurementsList();
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                if (worksheet.Name.Contains("_"))
                {
                    string emoloyer_namber_str = worksheet.Name.Substring(worksheet.Name.LastIndexOf("_") + 1, worksheet.Name.Length - worksheet.Name.LastIndexOf("_") - 1);
                    if (worksheet.Name.Contains("Ведомость"))
                        EmployerMSGWorksheets.Add(worksheet);
                    else if (worksheet.Name.Contains("Люди"))
                        EmployerWorkConsumptionsWorksheets.Add(worksheet);
                }
            }
           
            this.ReloadAllModels();
            CurrentMSGExellModel = CommonMSGExellModel;

            this.ShowWorksheet(CommonMSGWorksheet);
            this.SetAllWorksheetsVisibleState(XlSheetVisibility.xlSheetHidden);
            labelConractCode.Label = $"Шифр:{CurrentMSGExellModel.ContractCode.Substring(0, 15)}\n" +
                                    $"Объект:{CurrentMSGExellModel.ContructionObjectCode.Substring(0, 15)}\n " +
                                    $"Подобъект:{CurrentMSGExellModel.ConstructionSubObjectCode.Substring(0, 15)}";
            first_start_flag = false;

            this.SetBtnsState(true);
        }

        private void btnChangeCommonMSG_Click(object sender, RibbonControlEventArgs e)
        {
            if (CommonMSGExellModel == null)
                this.ReloadAllModels();
            CurrentMSGExellModel = CommonMSGExellModel;
            this.SetAllWorksheetsVisibleState(XlSheetVisibility.xlSheetHidden);
            this.ShowWorksheet(CommonMSGWorksheet);
            this.ShowWorksheet(CommonWorkConsumptionsWorksheet);
            CommonMSGWorksheet.Activate();
            btnCalcLabournes.Enabled = true;
            btnCalcAll.Enabled = true;
            btnReloadWorksheets.Enabled = true;
            labelCurrentEmployerName.Label = $"ОБЩИЕ ДАННЫЕ";
        }
        private void btnCalcLabournes_Click(object sender, RibbonControlEventArgs e)
        {
            CurrentMSGExellModel.CalcLabourness();
        }
      
        private void btnCalcAll_Click(object sender, RibbonControlEventArgs e)
        {
            CurrentMSGExellModel.CalcAll();
            // CurrentMSGExellModel.FormatWorkSheetStyle();
        }
        private void buttonCalc_Click(object sender, RibbonControlEventArgs e)
        {
            CurrentMSGExellModel.CalcQuantity();
            CurrentMSGExellModel.SetStyleFormats();
           // CurrentMSGExellModel.SetFormulas();
        }

        private void btnShowAlllHidenWorksheets_Click(object sender, RibbonControlEventArgs e)
        {
            this.SetAllWorksheetsVisibleState(XlSheetVisibility.xlSheetVisible);

        }
        private void btnChangeUOM_Click(object sender, RibbonControlEventArgs e)
        {
            this.SetAllWorksheetsVisibleState(XlSheetVisibility.xlSheetHidden);
            this.ShowWorksheet(UnitMeasurementsWorksheet);
            labelCurrentEmployerName.Label = $"ЕД. ИЗМ.";
        }


        private void comboBoxEmployerName_TextChanged(object sender, RibbonControlEventArgs e)
        {
            string sected_empl_name = comboBoxEmployerName.Text;
            SelectedEmloeyer = Employers.FirstOrDefault(em => em.Name == sected_empl_name);
            bntChangeEmployerMSG.Enabled = SelectedEmloeyer != null;
        }
        private void bntChangeEmployerMSG_Click(object sender, RibbonControlEventArgs e)
        {
            MSGExellModel empl_model = MSGExellModels.FirstOrDefault(m => m.Employer.Name == SelectedEmloeyer.Name);
            if (empl_model == null) //Если оаботник новый и на него нет еще модель и листы в книге - создаем их
            {
                Excel.Worksheet new_employer_worksheet = CurrentWorkbook.Worksheets.Add(CommonMSGWorksheet, Type.Missing, Type.Missing, Type.Missing);
                string new_worksheet_name = CommonMSGWorksheet.Name.Substring(0, CommonMSGWorksheet.Name.IndexOf('_') + 1) + SelectedEmloeyer.Number.ToString();
                new_employer_worksheet.Name = new_worksheet_name;

                Range last_source = CommonMSGWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);

                Excel.Range source = CommonMSGWorksheet.Range[CommonMSGWorksheet.Rows[1],
                    CommonMSGWorksheet.Rows[MSGExellModel.FIRST_ROW_INDEX - 1]];
                source.Copy();
                Range last_dest = new_employer_worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                Excel.Range dest = new_employer_worksheet.Range[new_employer_worksheet.Cells[1, 1], last_dest];
                dest.PasteSpecial(XlPasteType.xlPasteAll);

                EmployerMSGWorksheets.Add(new_employer_worksheet);

                Excel.Worksheet employer_worker_consumption_worksheet = CurrentWorkbook.Worksheets.Add(CommonWorkConsumptionsWorksheet, Type.Missing, Type.Missing, Type.Missing);

                string work_consumptions_worksheet_name = CommonWorkConsumptionsWorksheet.Name.Substring(0, CommonWorkConsumptionsWorksheet.Name.IndexOf('_') + 1) + SelectedEmloeyer.Number.ToString();
                employer_worker_consumption_worksheet.Name = work_consumptions_worksheet_name;

                last_source = CommonWorkConsumptionsWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                source = CommonWorkConsumptionsWorksheet.Range[CommonWorkConsumptionsWorksheet.Cells[1, 1], last_source];
                source.Copy();
                last_dest = employer_worker_consumption_worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                dest = employer_worker_consumption_worksheet.Range[employer_worker_consumption_worksheet.Cells[1, 1], last_dest];
                dest.PasteSpecial(XlPasteType.xlPasteAll);
                EmployerWorkConsumptionsWorksheets.Add(employer_worker_consumption_worksheet);

                this.ReloadAllModels();

                empl_model = MSGExellModels.FirstOrDefault(m => m.Employer.Name == SelectedEmloeyer.Name);
                //  empl_model.ClearWorksheetDaysPart();
            }
            CurrentMSGExellModel = empl_model;


            this.SetAllWorksheetsVisibleState(XlSheetVisibility.xlSheetHidden);
            this.ShowWorksheet(empl_model.RegisterSheet);
            this.ShowWorksheet(empl_model.WorkerConsumptionsSheet);
            empl_model.RegisterSheet.Activate();
            labelCurrentEmployerName.Label = $"ОТВЕСТВЕННЫЙ: {empl_model.Employer.Name}";
     //       CurrentMSGExellModel.ResetCalculatesFields();
        }
        private void bntChangeEmployerWorkersConsumption_Click(object sender, RibbonControlEventArgs e)
        {
            //MSGExellModel empl_model = MSGExellModels.FirstOrDefault(m => m.Employer.Name == SelectedEmloeyer.Name);
            //if (empl_model != null)
            //{
            //    Excel.Worksheet employer_worker_consumption_worksheet = CurrentWorkbook.Worksheets.Add(CommonMSGWorksheet, Type.Missing, Type.Missing, Type.Missing);

            //}


        }
        private void btnChangeEmployers_Click(object sender, RibbonControlEventArgs e)
        {
            this.SetAllWorksheetsVisibleState(XlSheetVisibility.xlSheetHidden);
            this.ShowWorksheet(EmployersWorksheet);
            labelCurrentEmployerName.Label = $"ОТВЕТСТВЕННЫЕ";
        }
        private void btnChangePosts_Click(object sender, RibbonControlEventArgs e)
        {
            this.SetAllWorksheetsVisibleState(XlSheetVisibility.xlSheetHidden);
            this.ShowWorksheet(PostsWorksheet);
            labelCurrentEmployerName.Label = $"ДОЛЖНОСТИ";
        }

        /// <summary>
        /// Функция делает видимими или не видимыми соотвествующие листы Workbook
        /// </summary>
        /// <param name="visibility"></param>
        private void SetAllWorksheetsVisibleState(Excel.XlSheetVisibility visibility)
        {
            CommonMSGWorksheet.Visible = visibility;
            CommonWorkConsumptionsWorksheet.Visible = visibility;
            EmployersWorksheet.Visible = visibility;
            PostsWorksheet.Visible = visibility;
            UnitMeasurementsWorksheet.Visible = visibility;

            foreach (Excel.Worksheet worksheet in EmployerMSGWorksheets)
                worksheet.Visible = visibility;

            foreach (Excel.Worksheet worksheet in EmployerWorkConsumptionsWorksheets)
                worksheet.Visible = visibility;
        }
        /// <summary>
        /// Функция делает видимой и активной соотвествующий лист в Workbook
        /// </summary>
        /// <param name="worksheet"></param>
        private void ShowWorksheet(Excel.Worksheet worksheet)
        {
            worksheet.Visible = XlSheetVisibility.xlSheetVisible;
            worksheet.Activate();

        }

        private void ReloadEmployersList()
        {
            Employers.Clear();
            ObservableCollection<Post> PostsList = new ObservableCollection<Post>();
            int row_index = 2;
            while (PostsWorksheet.Cells[row_index, POST_NUMBER_COL].Value != null)
            {
                //int number = int.Parse(PostsWorksheet.Cells[row_index, POST_NUMBER_COL].Value.ToString());
                string number = PostsWorksheet.Cells[row_index, POST_NUMBER_COL].Value.ToString();
                string name = PostsWorksheet.Cells[row_index, POST_NAME_COL].Value.ToString();
                PostsList.Add(new Post(number, name));
                row_index++;
            }
            row_index = 2;
            while (EmployersWorksheet.Cells[row_index, EMPLOYER_NUMBER_COL].Value != null)
            {
                string number = EmployersWorksheet.Cells[row_index, EMPLOYER_NUMBER_COL].Value.ToString();
                string name = EmployersWorksheet.Cells[row_index, EMPLOYER_NAME_COL].Value.ToString();
                string post_name = EmployersWorksheet.Cells[row_index, EMPLOYER_POSTNAME_COL].Value.ToString();
                Employers.Add(new Employer(number, name, PostsList.FirstOrDefault(p => p.Name == post_name)));
                row_index++;
            }
            comboBoxEmployerName.Items.Clear();
            foreach (Employer employer in Employers)
            {
                RibbonDropDownItem ddItem1 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                ddItem1.Label = employer.Name;
                comboBoxEmployerName.Items.Add(ddItem1);
            }
        }
        private void ReloadMeasurementsList()
        {
            UnitOfMeasurements.Clear();
            int row_index = 2;
            while (UnitMeasurementsWorksheet.Cells[row_index, POST_NUMBER_COL].Value != null)
            {
                int number = int.Parse(UnitMeasurementsWorksheet.Cells[row_index, UM_NUMBER_COL].Value.ToString());
                string name = UnitMeasurementsWorksheet.Cells[row_index, UM_NAME_COL].Value.ToString();
                UnitOfMeasurements.Add(new UnitOfMeasurement(number, name));
                row_index++;
            }

        }
        private void ReloadAllModels()
        {
            MSGExellModels.Clear();
            CommonMSGExellModel = new MSGExellModel();
            CommonMSGExellModel.RegisterSheet = CommonMSGWorksheet;
            CommonMSGExellModel.WorkerConsumptionsSheet = CommonWorkConsumptionsWorksheet;
            CommonMSGExellModel.CommonSheet = CommonWorksheet;
            CommonMSGExellModel.UnitOfMeasurements = UnitOfMeasurements;
            
            foreach (Excel.Worksheet worksheet in EmployerMSGWorksheets)
            {
                string emoloyer_namber_str = worksheet.Name.Substring(worksheet.Name.LastIndexOf("_") + 1, worksheet.Name.Length - worksheet.Name.LastIndexOf("_") - 1);
                Employer employer = Employers.Where(em => em.Number == emoloyer_namber_str).FirstOrDefault();
                if (employer != null && worksheet.Name.Contains("Ведомость"))
                {
                    MSGExellModel model = new MSGExellModel();
                    model.RegisterSheet = worksheet;
                    model.CommonSheet = CommonWorksheet;
                    model.UnitOfMeasurements = UnitOfMeasurements;
                    model.Employer = employer;
                    model.Owner = CommonMSGExellModel;
                    CommonMSGExellModel.Children.Add(model);
                    MSGExellModels.Add(model);
                }
            }
            foreach (Excel.Worksheet worksheet in EmployerWorkConsumptionsWorksheets)
            {
                string emoloyer_namber_str = worksheet.Name.Substring(worksheet.Name.LastIndexOf("_") + 1, worksheet.Name.Length - worksheet.Name.LastIndexOf("_") - 1);
                Employer employer = Employers.Where(em => em.Number == emoloyer_namber_str).FirstOrDefault();
                var model = MSGExellModels.FirstOrDefault(m => m.Employer.Number == emoloyer_namber_str);
                if (model != null && worksheet.Name.Contains("Люди"))
                    model.WorkerConsumptionsSheet = worksheet;
            }
            CommonMSGExellModel.ReloadAllSheetsInModel();
            foreach (MSGExellModel model in this.MSGExellModels)
            {
                model.CopyOwnerObjectModels();
               // model.LoadWorksReportCards();
            }
           
        }

        private void btnReloadWorksheets_Click(object sender, RibbonControlEventArgs e)
        {

            CurrentMSGExellModel.Update();

        }

        private void btnLoadTeplateFile_Click(object sender, RibbonControlEventArgs e)
        {
            //  string solutionpath = Directory.GetParent(Globals.ThisAddIn.Application.Path).Parent.Parent.Parent.FullName; 
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse Text Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xlsx",
                Filter = "xlsx files (*.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };
            string temlate_file_name;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                temlate_file_name = openFileDialog1.FileName;
                CurrentMSGExellModel.Update();
                CurrentMSGExellModel.CalcAll();
                MSGTemplateWorkbook = Globals.ThisAddIn.Application.Workbooks.Open(temlate_file_name);
                MSGTemplateWorkbook.Activate();
                if (CommonMSGExellModel != null)
                    this.FillMSG_OUT_File();
                btnFillTemlate.Enabled = true;
            }
        }

        private void btnFillTemlate_Click(object sender, RibbonControlEventArgs e)
        {

            if (CommonMSGExellModel != null)
                this.FillMSG_OUT_File();
        }
        #region Генерация МСГ в формате ЗМУО
        const int TMP_NOW_DATE_ROW = 1;
        const int TMP_NOW_DATE_COL = 1;

        const int TMP_CONTRACT_CODE_ROW = 3;
        const int TMP_COMMON_PARAMETRS_VALUE_COL = 2;


        const int TMP_CONSTRUCTION_OBJECT_CODE_ROW = 4;

        const int TMP_WORK_SELECTION_FIRST_ROW = 5;

        const int TMP_WORK_FIRST_INDEX_ROW = 6;

        const int TMP_WORK_NUMBER_COL = 1;
        const int TMP_WORK_NAME_COL = 2;
        const int TMP_WORK_PROJECT_QUANTITY_COL = 4;
        const int TMP_U_MRASURE_COL = 5;
        const int TMP_WORK_START_DATE_COL = 13;
        const int TMP_WORK_END_DATE_COL = 14;
        const int TMP_WORK_DAYS_NUMBER_COL = 15;

        const int TMP_PREVIOUS_WORK_QUANTITY_COL = 17;

        const int TMP_WORKDAY_DATE_ROW_COL = 2;
        const int TMP_WORKDAY_DATE_FIRST_COL = 19;

        const int WORKDAY_DATE_ROW = 2;
        const int WORKDAY_DATE_FIRST_COL = 18;

        const int NEEDS_NOW_DATE_ROW = 9;
        const int NEEDS_NOW_DATE_COL = 5;

        const int NEEDS_CONTRACT_CODE_ROW = 11;
        const int NEEDS_CONSTRUCTION_OBJECT_CODE_ROW = 10;

        const int NEEDS_WORKDAY_DATE_ROW = 9;
        const int NEEDS_WORKDAY_DATE_FIRST_COL = 10;
        const int NEEDS_WORKERS_FIRST_ROW = 12;
        const int NEEDS_WORKERS_NAME_COL = 6;

        private void FillMSG_OUT_File()
        {


            int row_index = TMP_WORK_SELECTION_FIRST_ROW;
            const int PLAN_PERIOD_MANTHS_NUMBER = 1;

            int work_needs_iterator = 0;

            Excel.Worksheet MSGOutWorksheet = MSGTemplateWorkbook.Worksheets["МСГ"];
            Excel.Worksheet MSGNeedsOutWorksheet = MSGTemplateWorkbook.Worksheets["Людские, технические ресурсы"];

            Excel.Worksheet MSGTemplateWorksheet = MSGTemplateWorkbook.Worksheets["МСГ_Шаблон"];
            Excel.Worksheet MSGNeedsTemplateWorksheet = MSGTemplateWorkbook.Worksheets["Людские_тех_ресурсы_Шаблон"];
            //MSGOutWorksheet.Visible = XlSheetVisibility.xlSheetHidden;
            //MSGNeedsOutWorksheet.Visible = XlSheetVisibility.xlSheetHidden;
            //MSGTemplateWorksheet.Visible = XlSheetVisibility.xlSheetHidden;
            //   MSGNeedsTemplateWorksheet.Visible = XlSheetVisibility.xlSheetHidden; 
            DateTime current_day_date = DateTime.Now;

            MSGOutWorksheet.Cells[TMP_NOW_DATE_ROW, TMP_NOW_DATE_COL] = current_day_date.ToString("d");
            MSGOutWorksheet.Cells[TMP_CONTRACT_CODE_ROW, TMP_COMMON_PARAMETRS_VALUE_COL] = CommonMSGExellModel.ContractCode;
            MSGOutWorksheet.Cells[TMP_CONSTRUCTION_OBJECT_CODE_ROW, TMP_COMMON_PARAMETRS_VALUE_COL] = CommonMSGExellModel.ContructionObjectCode;

            MSGNeedsTemplateWorksheet.Cells[NEEDS_NOW_DATE_ROW, NEEDS_NOW_DATE_COL] = current_day_date.ToString("d");
            MSGNeedsTemplateWorksheet.Cells[NEEDS_CONTRACT_CODE_ROW, NEEDS_NOW_DATE_COL] = CommonMSGExellModel.ContructionObjectCode;
            MSGNeedsTemplateWorksheet.Cells[NEEDS_CONSTRUCTION_OBJECT_CODE_ROW, NEEDS_NOW_DATE_COL] = CommonMSGExellModel.ContractCode;

            if (checkBoxRerightDatePart.Checked)
                this.FillMSG_OUT_File_Headers();

            int date_col_index = 0;
            int in_worksheet_number = 18;
            int worked_days_number = (CommonMSGExellModel.WorksEndDate - CommonMSGExellModel.WorksStartDate).Days;
            int last_day_col_index = Convert.ToInt32(Math.Round(WORKDAY_DATE_FIRST_COL + worked_days_number * (1 + 1 / 7.0) + 5));



            Excel.Range tmp_works_selection_range = MSGTemplateWorksheet.UsedRange.Rows[TMP_WORK_SELECTION_FIRST_ROW];
            Excel.Range tmp_work_rows_range = MSGOutWorksheet.Range[MSGOutWorksheet.UsedRange.Rows[TMP_WORK_FIRST_INDEX_ROW], MSGOutWorksheet.UsedRange.Rows[TMP_WORK_FIRST_INDEX_ROW + 1]];
            Excel.Range tmp_dest = MSGTemplateWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW, 1];
            tmp_work_rows_range.Copy();
            tmp_dest.PasteSpecial(XlPasteType.xlPasteAll);
            tmp_work_rows_range = MSGTemplateWorksheet.Range[MSGTemplateWorksheet.UsedRange.Rows[TMP_WORK_FIRST_INDEX_ROW], MSGTemplateWorksheet.UsedRange.Rows[TMP_WORK_FIRST_INDEX_ROW + 1]];

            #region Заполнение формы данными из модели...
            const int LAST_ROW_MAX_COUNT = 100;
            int last_not_null_row_index = TMP_WORK_SELECTION_FIRST_ROW;
            int work_local_index_iterator = TMP_WORK_SELECTION_FIRST_ROW;
            int saved_iterator = TMP_WORK_SELECTION_FIRST_ROW;
            foreach (WorksSection w_section in CommonMSGExellModel.WorksSections)
            {
                int section_local_index_iterator = TMP_WORK_SELECTION_FIRST_ROW;
                int section_null_cell_counter = 0;

                while (section_null_cell_counter <= LAST_ROW_MAX_COUNT)
                {
                    if (MSGOutWorksheet.Cells[section_local_index_iterator, TMP_WORK_NUMBER_COL].Value == null)
                        section_null_cell_counter++;
                    else
                    {
                        section_null_cell_counter = 0;
                        string w_section_name = MSGOutWorksheet.Cells[section_local_index_iterator, TMP_WORK_NUMBER_COL].Value.ToString();
                        if (w_section_name == w_section.Name)
                        {
                            saved_iterator = section_local_index_iterator;
                            break;
                        }


                    }
                    section_local_index_iterator++;
                }
                if (section_null_cell_counter >= LAST_ROW_MAX_COUNT)
                    section_local_index_iterator = saved_iterator;

                tmp_works_selection_range.Copy();
                Excel.Range sect_row_dest = MSGOutWorksheet.Cells[section_local_index_iterator, 1];
                sect_row_dest.PasteSpecial(XlPasteType.xlPasteAll);
                MSGOutWorksheet.Cells[section_local_index_iterator, 1] = w_section.Name;

                //    row_index++;
                saved_iterator = section_local_index_iterator + 1;
                foreach (MSGWork msg_work in w_section.MSGWorks)
                {
                    ///Копируем и вставляем строку для работы в МСГ
                    work_local_index_iterator = section_local_index_iterator + 1;
                    int null_cell_counter = 0;


                    while (null_cell_counter <= LAST_ROW_MAX_COUNT)
                    {
                        if (MSGOutWorksheet.Cells[work_local_index_iterator, TMP_WORK_NUMBER_COL].Value == null)
                            null_cell_counter++;
                        else
                        {
                            null_cell_counter = 0;
                            string msg_work_number = "";
                            if (MSGOutWorksheet.Cells[work_local_index_iterator, TMP_WORK_NUMBER_COL].Value != null)
                                msg_work_number = MSGOutWorksheet.Cells[work_local_index_iterator, TMP_WORK_NUMBER_COL].Value.ToString();



                            if (CommonMSGExellModel.WorksSections.FirstOrDefault(wc => wc.Name == msg_work_number) != null
                                                 && msg_work_number != w_section.Name)
                            {
                                saved_iterator = work_local_index_iterator;
                                tmp_work_rows_range.Copy();
                                Excel.Range dest = MSGOutWorksheet.Rows[saved_iterator];
                                // dest.PasteSpecial(XlPasteType.xlPasteAll);
                                dest.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Type.Missing);

                                break;
                            }
                            //var f = MSGOutWorksheet.Cells[work_local_index_iterator, TMP_WORK_NUMBER_COL].Value;
                            //var f2 = MSGOutWorksheet.Cells[work_local_index_iterator + 1, TMP_WORK_NUMBER_COL].Value;

                            //if (MSGOutWorksheet.Cells[work_local_index_iterator, TMP_WORK_NUMBER_COL].Value == null &&
                            //     MSGOutWorksheet.Cells[work_local_index_iterator + 1, TMP_WORK_NUMBER_COL].Value == null )
                            //{

                            //    row_index = work_local_index_iterator;
                            //    tmp_work_rows_range.Copy();
                            //    Excel.Range dest = MSGOutWorksheet.Rows[row_index];
                            //    dest.PasteSpecial(XlPasteType.xlPasteAll);
                            //    break;
                            //}
                            if (msg_work.Number == msg_work_number)
                            {
                                saved_iterator = work_local_index_iterator;
                                break;
                            }
                            if (msg_work.Number != msg_work_number && msg_work_number != "")
                            {
                                saved_iterator = work_local_index_iterator + 2;
                                //      section_local_index_iterator = work_local_index_iterator;

                                //  break;
                            }
                            //if (MSGOutWorksheet.Cells[local_index_iterator+2, TMP_WORK_NUMBER_COL].Value == null &&
                            //     MSGOutWorksheet.Cells[local_index_iterator + 3, TMP_WORK_NUMBER_COL].Value == null)
                            //{
                            //    row_index +=2;
                            //    tmp_work_rows_range.Copy();
                            //    Excel.Range dest = MSGOutWorksheet.Rows[row_index];
                            //    dest.PasteSpecial(XlPasteType.xlPasteAll);
                            //    break;
                            //}

                        }
                        work_local_index_iterator++;
                    }
                    row_index = saved_iterator;
                    if (null_cell_counter >= LAST_ROW_MAX_COUNT)
                    {
                        //  row_index = saved_iterator;
                        tmp_work_rows_range.Copy();
                        Excel.Range dest = MSGOutWorksheet.Rows[row_index];
                        dest.PasteSpecial(XlPasteType.xlPasteAll);
                        saved_iterator += 2;
                    }

                    //tmp_work_rows_range.Copy();
                    //Excel.Range dest = MSGOutWorksheet.Rows[row_index];
                    //dest.Insert(Excel.XlInsertShiftDirection.xlShiftDown,Type.Missing);

                    ///Заполняем основыне данные работы                
                    MSGOutWorksheet.Cells[row_index, TMP_WORK_NUMBER_COL] = msg_work.Number;
                    MSGOutWorksheet.Cells[row_index, TMP_WORK_NAME_COL] = msg_work.Name;
                    MSGOutWorksheet.Cells[row_index, TMP_WORK_PROJECT_QUANTITY_COL] = msg_work.ProjectQuantity;
                    MSGOutWorksheet.Cells[row_index, TMP_U_MRASURE_COL] = msg_work.UnitOfMeasurement.Name;

                    MSGOutWorksheet.Cells[row_index, TMP_WORK_START_DATE_COL] = msg_work.WorkSchedules.StartDate;
                    MSGOutWorksheet.Cells[row_index, TMP_WORK_END_DATE_COL] = msg_work.WorkSchedules.EndDate;

                    MSGOutWorksheet.Cells[row_index + 1, TMP_PREVIOUS_WORK_QUANTITY_COL] = msg_work.PreviousComplatedQuantity;
                    // MSGOutWorksheet.Cells[row_index, TMP_WORK_DAYS_NUMBER_COL] = (msg_work.WorkSchedules.EndDate - msg_work.WorkSchedules.StartDate)?.Days;
                    ///Заполняем плановые объемы в календарной части
                    foreach (WorkScheduleChunk schedule_chunk in msg_work.WorkSchedules)
                    {
                        int date_index = 0;
                        while (MSGOutWorksheet.Cells[TMP_WORKDAY_DATE_ROW_COL, TMP_WORKDAY_DATE_FIRST_COL + date_index].Value != null && date_index <= last_day_col_index)
                        {
                            DateTime date;
                            DateTime.TryParse(MSGOutWorksheet.Cells[TMP_WORKDAY_DATE_ROW_COL, TMP_WORKDAY_DATE_FIRST_COL + date_index].Value.ToString(), out date);

                            if (date >= schedule_chunk.StartTime && date <= schedule_chunk.EndTime
                                && (date.DayOfWeek != DayOfWeek.Sunday || !checkBoxSandayVocationrStatus.Checked))
                            {

                                MSGOutWorksheet.Cells[row_index, TMP_WORKDAY_DATE_FIRST_COL + date_index] =
                                    msg_work.ProjectQuantity / msg_work.GetShedulesAllDaysNumber(checkBoxSandayVocationrStatus.Checked);

                            }
                            date_index++;
                        }
                    }
                    ///Заполняем  фактически выполненные объемы в календарной части
                    if (msg_work.ReportCard != null)
                        foreach (WorkDay msg_work_day in msg_work.ReportCard)
                        {
                            int date_index = 0;
                            while (MSGOutWorksheet.Cells[TMP_WORKDAY_DATE_ROW_COL, TMP_WORKDAY_DATE_FIRST_COL + date_index].Value != null && date_index <= last_day_col_index)
                            {
                                DateTime date;
                                DateTime.TryParse(MSGOutWorksheet.Cells[TMP_WORKDAY_DATE_ROW_COL, TMP_WORKDAY_DATE_FIRST_COL + date_index].Value.ToString(), out date);


                                if (date == msg_work_day.Date)
                                {

                                    MSGOutWorksheet.Cells[row_index + 1, TMP_WORKDAY_DATE_FIRST_COL + date_index] = msg_work_day.Quantity;
                                }
                                date_index++;
                            }
                        }
                    //row_index += 2;
                    work_local_index_iterator += 2;
                }
            }
            #endregion

            #region Заполняем потребности 

            work_needs_iterator = 0;

            while (MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator, NEEDS_WORKERS_NAME_COL].Value != "Общее количество")
            {
                int work_needs_date_col_index = 1;
                DateTime current_date;
                var ss = MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW, NEEDS_WORKDAY_DATE_FIRST_COL + work_needs_date_col_index].Value;
                DateTime.TryParse(MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW, NEEDS_WORKDAY_DATE_FIRST_COL + work_needs_date_col_index].Value.ToString("d"), out current_date);
                string worker_post_name = MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator, NEEDS_WORKERS_NAME_COL].Value;
                var current_needs_of_worker = CommonMSGExellModel.WorkersComposition.FirstOrDefault(nw => nw.Name == worker_post_name);
                var current_worker_consumption = CommonMSGExellModel.WorkerConsumptions.FirstOrDefault(wc => wc.Name == worker_post_name);

                while (work_needs_date_col_index < last_day_col_index && current_needs_of_worker != null)
                {
                    NeedsOfWorkersDay needsOfWorkersDay = current_needs_of_worker.NeedsOfWorkersReportCard.FirstOrDefault(nwd => nwd.Date == current_date);
                    if (needsOfWorkersDay != null)
                    {
                        MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator,
                            NEEDS_WORKDAY_DATE_FIRST_COL + work_needs_date_col_index] = needsOfWorkersDay.Quantity;
                    }

                    WorkerConsumptionDay worker_consumption_day = current_worker_consumption.WorkersConsumptionReportCard.FirstOrDefault(wcd => wcd.Date == current_date);
                    if (worker_consumption_day != null)
                    {
                        MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator + 1,
                            NEEDS_WORKDAY_DATE_FIRST_COL + work_needs_date_col_index] = worker_consumption_day.Quantity;
                    }

                    work_needs_date_col_index++;
                    if (MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW, NEEDS_WORKDAY_DATE_FIRST_COL + work_needs_date_col_index].Value == null) break;
                    DateTime.TryParse(MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW, NEEDS_WORKDAY_DATE_FIRST_COL + work_needs_date_col_index].Value.ToString(), out current_date);

                }


                work_needs_iterator++;
            }
            MSGOutWorksheet.Visible = XlSheetVisibility.xlSheetVisible;
            MSGNeedsOutWorksheet.Visible = XlSheetVisibility.xlSheetVisible;
            MSGTemplateWorksheet.Visible = XlSheetVisibility.xlSheetVisible;
            MSGNeedsTemplateWorksheet.Visible = XlSheetVisibility.xlSheetVisible;

            #endregion
            MSGTemplateWorkbook.SaveAs(@"D:\1234.xlsx");
            MSGTemplateWorkbook.Close();
        }
        private void FillMSG_OUT_File_Headers()
        {
            #region Создание календарной части - дней, недель...
            Excel.Worksheet MSGOutWorksheet = MSGTemplateWorkbook.Worksheets["МСГ"];
            Excel.Worksheet MSGNeedsOutWorksheet = MSGTemplateWorkbook.Worksheets["Людские, технические ресурсы"];

            Excel.Worksheet MSGTemplateWorksheet = MSGTemplateWorkbook.Worksheets["МСГ_Шаблон"];
            Excel.Worksheet MSGNeedsTemplateWorksheet = MSGTemplateWorkbook.Worksheets["Людские_тех_ресурсы_Шаблон"];

            Excel.Range tmp_week_col_range = MSGTemplateWorksheet.UsedRange.Columns[TMP_WORKDAY_DATE_FIRST_COL - 1];
            Excel.Range tmp_date_col_range = MSGTemplateWorksheet.UsedRange.Columns[TMP_WORKDAY_DATE_FIRST_COL];

            Excel.Range tmp_needs_week_col_range = MSGNeedsTemplateWorksheet.UsedRange.Columns[NEEDS_WORKDAY_DATE_FIRST_COL + 1]; //Столбец недели в шаблоне ресурсов
            Excel.Range tmp_needs_date_col_range = MSGNeedsTemplateWorksheet.UsedRange.Columns[NEEDS_WORKDAY_DATE_FIRST_COL];//Столбец дня в шаблоне ресурсов


            int work_needs_iterator = 0;
            int date_col_index = 0;
            int in_worksheet_number = 0;

            int first_week_day_col;
            int last_week_day_col;

            int week_signatura_first_col = 0;
            int week_signatura_last_col = 0;

            string last_week_name_signatura = "";
            ///Цикл создающий календарные дни и недели до конца периода..
            if (checkBoxRerightDatePart.Checked == true)
                for (DateTime date = CommonMSGExellModel.WorksStartDate; date <= CommonMSGExellModel.WorksEndDate; date = date.AddDays(1))
                {
                    ///Если текущий день является восскесеньем или является первым днем все ведомости -
                    /// втавляем и заполняем недельный столбец в календарь
                    if (date.DayOfWeek == DayOfWeek.Monday || date == CommonMSGExellModel.WorksStartDate)
                    {
                        #region Календраня часть МСГ недельный столбец
                        ///Вставляем из шаблона МСГ недельный столбец в каледаре ( копируем из шаблона)
                        tmp_week_col_range.Copy();
                        Excel.Range week_day_dest = MSGOutWorksheet.UsedRange.Columns[WORKDAY_DATE_FIRST_COL + date_col_index];
                        week_day_dest.PasteSpecial(XlPasteType.xlPasteAll);

                        ///Заполняем из шаблона МСГ недельный столбец в каледаре 
                        string week_name_signatura =
                          $"неделя\r\n {date.ToString("dd")} - {this.GetLastNotVocationDate(date).AddDays(1).ToString("dd")}";
                        MSGOutWorksheet.Cells[WORKDAY_DATE_ROW, WORKDAY_DATE_FIRST_COL + date_col_index] = week_name_signatura;
                        MSGOutWorksheet.Cells[WORKDAY_DATE_ROW + 1, WORKDAY_DATE_FIRST_COL + date_col_index] = in_worksheet_number++;


                        first_week_day_col = date_col_index + 1;
                        last_week_day_col = first_week_day_col + (this.GetLastNotVocationDate(date).AddDays(1) - date).Days;

                        Excel.Range project_week_pr_q = MSGOutWorksheet.Range[
                          MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW, WORKDAY_DATE_FIRST_COL + date_col_index],
                          MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW, WORKDAY_DATE_FIRST_COL + date_col_index]];

                        Excel.Range project_week_q = MSGOutWorksheet.Range[
                            MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW + 1, WORKDAY_DATE_FIRST_COL + date_col_index],
                            MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW + 1, WORKDAY_DATE_FIRST_COL + date_col_index]];

                        Excel.Range project_week_pr_q_first_day = MSGOutWorksheet.Range[
                             MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW, WORKDAY_DATE_FIRST_COL + first_week_day_col],
                             MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW, WORKDAY_DATE_FIRST_COL + first_week_day_col]];

                        Excel.Range project_week_pr_q_last_day = MSGOutWorksheet.Range[
                            MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW, WORKDAY_DATE_FIRST_COL + last_week_day_col],
                            MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW, WORKDAY_DATE_FIRST_COL + last_week_day_col]];

                        Excel.Range project_week_q_first_day = MSGOutWorksheet.Range[
                            MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW + 1, WORKDAY_DATE_FIRST_COL + first_week_day_col],
                            MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW + 1, WORKDAY_DATE_FIRST_COL + first_week_day_col]];

                        Excel.Range project_week_q_last_day = MSGOutWorksheet.Range[
                            MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW + 1, WORKDAY_DATE_FIRST_COL + last_week_day_col],
                            MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW + 1, WORKDAY_DATE_FIRST_COL + last_week_day_col]];
                        ///Вставляем форму для подсчета суммы планового количества работан на  неделю и фактического объема...
                        project_week_q.Formula = $"=SUM({Func.RangeAddress(project_week_q_first_day)}:{Func.RangeAddress(project_week_q_last_day)})"; ;
                        project_week_pr_q.Formula = $"=SUM({Func.RangeAddress(project_week_pr_q_first_day)}:{Func.RangeAddress(project_week_pr_q_last_day)})"; ;
                        #endregion

                        #region Календарная часть потребности ресурсов недельный столбец
                        ///Вставляем из шаблона ресурсов  недельный столбец в каледаре ( копируем из шаблона)
                        tmp_needs_week_col_range.Copy();
                        week_day_dest = MSGNeedsOutWorksheet.UsedRange.Columns[NEEDS_WORKDAY_DATE_FIRST_COL + date_col_index];
                        week_day_dest.PasteSpecial(XlPasteType.xlPasteAll);
                        ///Заполняем из шаблона ресурсов  недельный столбец в каледаре ( копируем из шаблона)
                        MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW, NEEDS_WORKDAY_DATE_FIRST_COL + date_col_index] = "Всего";
                        if (week_signatura_last_col > 0)
                        {
                            Excel.Range week_name_range_first_cell = MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW - 1, NEEDS_WORKDAY_DATE_FIRST_COL + week_signatura_first_col];
                            Excel.Range week_name_range_last_cell = MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW - 1, NEEDS_WORKDAY_DATE_FIRST_COL + week_signatura_last_col];
                            Excel.Range week_name_range = MSGNeedsOutWorksheet.get_Range(week_name_range_first_cell, week_name_range_last_cell);
                            week_name_range.Merge();
                            week_name_range_first_cell.Value = last_week_name_signatura.Replace("\r\n", " ");
                        }

                        work_needs_iterator = 0;
                        while (MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator, NEEDS_WORKERS_NAME_COL].Value != "Общее количество")
                        {
                            Excel.Range needs_first_day = MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator, NEEDS_WORKDAY_DATE_FIRST_COL + first_week_day_col];
                            Excel.Range needs_last_day = MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator, NEEDS_WORKDAY_DATE_FIRST_COL + last_week_day_col];


                            MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator, NEEDS_WORKDAY_DATE_FIRST_COL + date_col_index] =
                               $"=SUM({Func.RangeAddress(needs_first_day)}:{Func.RangeAddress(needs_last_day)})"; ;

                            work_needs_iterator++;
                        }



                        #endregion
                        week_signatura_first_col = first_week_day_col - 1;
                        week_signatura_last_col = last_week_day_col;
                        last_week_name_signatura = week_name_signatura;
                        first_week_day_col = 0;
                        last_week_day_col = 0;


                        date_col_index++;
                    }
                    #region Каледарная часть МСГ рабочие дни
                    tmp_date_col_range.Copy();
                    Excel.Range dest = MSGOutWorksheet.UsedRange.Columns[WORKDAY_DATE_FIRST_COL + date_col_index];
                    dest.PasteSpecial(XlPasteType.xlPasteAll);

                    MSGOutWorksheet.Cells[WORKDAY_DATE_ROW, WORKDAY_DATE_FIRST_COL + date_col_index] = date;
                    MSGOutWorksheet.Cells[WORKDAY_DATE_ROW + 1, WORKDAY_DATE_FIRST_COL + date_col_index] = in_worksheet_number++;
                    if (date.DayOfWeek == DayOfWeek.Sunday)
                        MSGOutWorksheet.Cells[WORKDAY_DATE_ROW, WORKDAY_DATE_FIRST_COL + date_col_index].Interior.Color
                                    = XlRgbColor.rgbOrangeRed;

                    #endregion

                    #region Календарная часть потребности ресурсов столбец ежедневных потребностей 

                    tmp_needs_date_col_range.Copy();
                    Excel.Range needs_day_dest = MSGNeedsOutWorksheet.UsedRange.Columns[NEEDS_WORKDAY_DATE_FIRST_COL + date_col_index];
                    needs_day_dest.PasteSpecial(XlPasteType.xlPasteAll);

                    MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW, NEEDS_WORKDAY_DATE_FIRST_COL + date_col_index].Value = date;
                    #endregion

                    date_col_index++;
                }
            #endregion
        }
        #endregion
        private void FillMSG_NEEDS_File()
        {
            const int TMP_CONTRACT_CODE_ROW = 3;

        }
        private DateTime GetLastNotVocationDate(DateTime date)
        {
            DateTime out_date = date;
            while (out_date.AddDays(1).DayOfWeek != DayOfWeek.Sunday)
                out_date = out_date.AddDays(1);
            return out_date;
        }
        

        private void checkBoxRerightDatePart_Click(object sender, RibbonControlEventArgs e)
        {

        }
        private IExcelBindableBase CopyedObject;
        private string commands_group_label = "";
        private void buttonCopy_Click(object sender, RibbonControlEventArgs e)
        {

            var selection = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            var sected_object = CurrentMSGExellModel.GetObjectBySelection(selection, typeof(WorksSection));
            switch (sected_object?.GetType().Name)
            {
                case nameof(WorksSection):
                    {
                        CopyedObject = (WorksSection)sected_object.Clone();
                        buttonPaste.Enabled = true;
                        commands_group_label = $"Разадел {CopyedObject.Name} скопирован.\n Выбрерите ячеку с номеров нового раздела.";

                        break;
                    }
                default:
                    {
                        buttonPaste.Enabled = true;
                        break;
                    }
            }
            groupCommands.Label = commands_group_label;
            //Excel.Dialog dlg = Globals.ThisAddIn.Application.Dialogs[Excel.XlBuiltInDialog.xlDialogFont];
            //dlg.Show();


        }

        private void buttonPaste_Click(object sender, RibbonControlEventArgs e)
        {
            var selection = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            string cell_val = selection.Value.ToString();
            switch (CopyedObject?.GetType().Name)
            {
                case nameof(WorksSection):
                    {
                        if (cell_val != null)
                        {
                            WorksSection section = CopyedObject as WorksSection;
                            CurrentMSGExellModel.WorksSections.Add(section);
                            section.SetNumber(cell_val);
                            CurrentMSGExellModel.UpdateSectionWorksheetCommonPart(section, selection.Row);
                            CurrentMSGExellModel.RegisterObjectInObjectPropertyNameRegister(section);
                            CurrentMSGExellModel.SetStyleFormats();
                            commands_group_label = "";
                        }
                        break;
                    }
                default:
                    {
                        break;
                    }
            }



        }

       
    }
}
