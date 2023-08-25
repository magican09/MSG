using ExellAddInsLib.MSG;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace MSGAddIn
{
    public partial class Ribbon1
    {


        private const int POST_NUMBER_COL = 1;
        private const int POST_NAME_COL = 2;

        private const int MACHINE_NUMBER_COL = 1;
        private const int MACHINE_NAME_COL = 2;

        private const int EMPLOYER_NUMBER_COL = 1;
        private const int EMPLOYER_NAME_COL = 2;
        private const int EMPLOYER_POSTNAME_COL = 3;

        private const int UM_NUMBER_COL = 1;
        private const int UM_NAME_COL = 2;


        MSGExellModel CurrentMSGExellModel;
        MSGExellModel CommonMSGExellModel;
        ObservableCollection<MSGExellModel> MSGExellModels = new ObservableCollection<MSGExellModel>();

        ObservableCollection<Employer> Employers { get; set; } = new ObservableCollection<Employer>();
        ObservableCollection<Machine> Machines { get; set; } = new ObservableCollection<Machine>();
        ExcelNotifyChangedCollection<UnitOfMeasurement> UnitOfMeasurements = new ExcelNotifyChangedCollection<UnitOfMeasurement>();

        Excel._Workbook CurrentWorkbook;
        Excel._Workbook MSGTemplateWorkbook;

        Excel.Worksheet EmployersWorksheet;
        Excel.Worksheet MachinesWorksheet;

        Excel.Worksheet PostsWorksheet;
        Excel.Worksheet UnitMeasurementsWorksheet;
        Excel.Worksheet CommonWorksheet;
        Excel.Worksheet CommonMSGWorksheet;
        Excel.Worksheet CommonWorkConsumptionsWorksheet;
        Excel.Worksheet CommonMachineConsumptionsWorksheet;
        Excel.Worksheet GuidWorksheet;

        public Guid Workbook_Guid { get; set; }

        ObservableCollection<Excel.Worksheet> EmployerMSGWorksheets = new ObservableCollection<Worksheet>();
        ObservableCollection<Excel.Worksheet> EmployerWorkConsumptionsWorksheets = new ObservableCollection<Worksheet>();
        ObservableCollection<Excel.Worksheet> MachineMSGWorksheets = new ObservableCollection<Worksheet>();
        ObservableCollection<Excel.Worksheet> MachineConsumptionsWorksheets = new ObservableCollection<Worksheet>();

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
            GuidWorksheet = new_wbk.Worksheets.OfType<Excel.Worksheet>().FirstOrDefault(w => w.Name == $"Guid_{Workbook_Guid.ToString().Split('-')[0]}");

            MachinesWorksheet = new_wbk.Worksheets.OfType<Excel.Worksheet>().FirstOrDefault(w => w.Name == "Машины_механизмы");
            if (CommonMSGWorksheet != null && CommonWorksheet != null
                && UnitMeasurementsWorksheet != null
                && PostsWorksheet != null
                && EmployersWorksheet != null && GuidWorksheet == null && Workbook_Guid == Guid.Empty)
            {
                this.SetApplicationGroupsVisibility(false);
                groupFileLaod.Visible = true;
                return;
            }

            if (CommonMSGWorksheet != null && CommonWorksheet != null
               && UnitMeasurementsWorksheet != null
               && PostsWorksheet != null
               && EmployersWorksheet != null && GuidWorksheet != null)
                this.SetApplicationGroupsVisibility(true);
            else
                this.SetApplicationGroupsVisibility(false);
        }

        private void SetApplicationGroupsVisibility(bool visibility)
        {
            groupFileLaod.Visible = visibility;
            InMSGWorkbook = visibility;
            groupMSGCommon.Visible = visibility;
            grpInChargePersons.Visible = visibility;
            groupCommands.Visible = visibility;
            groupMSG_OUT.Visible = visibility;
            groupInfo.Visible = visibility;
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
        private void OnBeforeCloseWorkbookChanged(Workbook Wb, ref bool Cancel)
        {
            if (CommonMSGWorksheet != null && CommonWorksheet != null
                && UnitMeasurementsWorksheet != null
                && PostsWorksheet != null
                && EmployersWorksheet != null && GuidWorksheet != null)
            {
                EmployersWorksheet = null;
                MachinesWorksheet = null;
                PostsWorksheet = null;
                UnitMeasurementsWorksheet = null;
                CommonWorksheet = null;
                CommonMSGWorksheet = null;
                CommonWorkConsumptionsWorksheet = null;
                CommonMachineConsumptionsWorksheet = null;

                EmployerWorkConsumptionsWorksheets.Clear();
                EmployerMSGWorksheets.Clear();
                MachineMSGWorksheets.Clear();
                MachineConsumptionsWorksheets.Clear();

                this.Employers.Clear();
                this.Machines.Clear();
                this.UnitOfMeasurements.Clear();
                ///Загрузка всех  листов


                this.MSGExellModels.Clear();
                CommonMSGExellModel = null;

                CurrentMSGExellModel = null;

                labelConractCode.Label = $"";
                this.SetApplicationGroupsVisibility(false);
                groupFileLaod.Visible = true;
                Workbook_Guid = Guid.Empty;
            }

        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.OnActiveWorksheetChanged += OnActiveWorksheetChanged;
            Globals.ThisAddIn.OnActiveWorkbookChanged += OnActiveWorkbookChanged;
            Globals.ThisAddIn.OnBeforeCloseWorkbookChanged += OnBeforeCloseWorkbookChanged;

        }



        private void SetBtnsState(bool state)
        {
            btnUpdateAll.Enabled = state;
            btnLoadInModel.Enabled = state;
            btnLoadFromModel.Enabled = state;
            btnChangeCommonMSG.Enabled = state;

            btnCalcLabournes.Enabled = state;
            //   btnCreateTemplateFile.Enabled = state;
            buttonCalc.Enabled = state;

            menuSection.Enabled = state;
            menuMSG.Enabled = state;
            menuVOVR.Enabled = state;
            menuKS.Enabled = state;
            menuRC.Enabled = state;
            buttonPaste.Enabled = state;
            btnChangeEmployers.Enabled = state;
            btnChangePosts.Enabled = state;
            btnChangeUOM.Enabled = state;
            btnMachines.Enabled = state;
            chckBoxHashEnable.Enabled = state;
            btnLoadInModelLocal.Enabled = state & chckBoxHashEnable.Checked;
            btnCreateMSGForEmployers.Enabled = state;
        }
        private void AjastBtnsState()
        {
            this.SetBtnsState(true);
        }
        private void AddExcellVBAFunctions()
        {
            VBComponent vba_module = null;
            try
            {
                vba_module = CurrentWorkbook.VBProject.VBComponents.Item("Functions");
            }
            catch
            {
                if (vba_module == null)
                    vba_module = CurrentWorkbook.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                vba_module.Name = "Functions";
            }


            var codeModule = vba_module.CodeModule;
            var lineNum = codeModule.CountOfLines + 1;
            var macroName = "EasyHash";
            var lines = "";
            if (codeModule.CountOfLines != 0)
                lines = codeModule.Lines[1, codeModule.CountOfLines];
            if (!lines.Contains(macroName))
            {
                var codeText = $"Function {macroName}(ByRef Str$) As Long\r\n" +
                             "Dim i As Integer, Hash As Long\r\n" +
                             "For i = 1 To Len(Str)\r\n" +
                             "        Hash = i + 1664525 * AscB(Mid(Str, i, 1)) + 1013904223\r\n" +
                             "        EasyHash = ((Hash Xor Abs(1365 / i)) And 65535) + EasyHash\r\n" +
                             "Next\r\n" +
                             "End Function";
                codeModule.InsertLines(lineNum, codeText);
            }

        }

        private void btnLoadMSGFile_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                CurrentWorkbook = Globals.ThisAddIn.CurrentActivWorkbook;
                EmployersWorksheet = CurrentWorkbook.Worksheets["Ответственные"];
                MachinesWorksheet = CurrentWorkbook.Worksheets["Машины_механизмы"];
                PostsWorksheet = CurrentWorkbook.Worksheets["Должности"];
                UnitMeasurementsWorksheet = CurrentWorkbook.Worksheets["Ед_изм"];
                CommonWorksheet = CurrentWorkbook.Worksheets["Начальная"];
                CommonMSGWorksheet = CurrentWorkbook.Worksheets["Ведомость_общая"];
                CommonWorkConsumptionsWorksheet = CurrentWorkbook.Worksheets["Люди_общая"];
                CommonMachineConsumptionsWorksheet = CurrentWorkbook.Worksheets["Техника_общая"];

             //   this.AddExcellVBAFunctions();

                EmployerMSGWorksheets = new ObservableCollection<Excel.Worksheet>();
                MachineMSGWorksheets = new ObservableCollection<Excel.Worksheet>();

                this.ReloadEmployersList();
                this.ReloadMachinesList();
                this.ReloadMeasurementsList();
                ///Загрузка всех  листов
                foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
                {
                    if (worksheet.Name.Contains("_"))
                    {
                        string emoloyer_namber_str = worksheet.Name.Substring(worksheet.Name.LastIndexOf("_") + 1, worksheet.Name.Length - worksheet.Name.LastIndexOf("_") - 1);
                        if (worksheet.Name.Contains("Ведомость_"))
                            EmployerMSGWorksheets.Add(worksheet);
                        else if (worksheet.Name.Contains("Люди_"))
                            EmployerWorkConsumptionsWorksheets.Add(worksheet);
                        else if (worksheet.Name.Contains("Техника_"))
                            MachineConsumptionsWorksheets.Add(worksheet);
                        else if (worksheet.Name.Contains("Guid_"))
                        {
                            GuidWorksheet = worksheet;
                            Workbook_Guid = Guid.NewGuid();
                            GuidWorksheet.Name = $"Guid_{Workbook_Guid.ToString().Split('-')[0]}";
                            GuidWorksheet.Visible = XlSheetVisibility.xlSheetHidden;
                            this.SetApplicationGroupsVisibility(true);
                        }
                    }
                }



                this.ReloadAllModels();


                CurrentMSGExellModel = CommonMSGExellModel;
                //labelConractCode.Label = $"Шифр:{CurrentMSGExellModel.ContractCode}\n" +
                //                        $"Объект:{CurrentMSGExellModel.ContructionObjectCode}\n " +
                //                        $"Подобъект:{CurrentMSGExellModel.ConstructionSubObjectCode}";
                labelConractCode.Label = $"Шифр:{CurrentMSGExellModel.ContractCode}\n" +
                                        $"Объект:{CurrentMSGExellModel.ContructionObjectCode}";

                this.SetBtnsState(true);

                //    CurrentMSGExellModel.SetFormulas(); 
                CurrentMSGExellModel.SetStyleFormats();
            }
               catch (Exception exp)
            {
                       MessageBox.Show($"Ошибка при зазугрузка данных. Ошибка: {exp.Message}");
            }

        }

        private void btnChangeCommonMSG_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (CommonMSGExellModel == null)
                    this.ReloadAllModels();
                CurrentMSGExellModel = CommonMSGExellModel;
                this.SetAllWorksheetsVisibleState(XlSheetVisibility.xlSheetHidden);
                this.ShowWorksheet(CommonMSGWorksheet);
                this.ShowWorksheet(CommonWorkConsumptionsWorksheet);
                this.ShowWorksheet(CommonMachineConsumptionsWorksheet);
                CommonMSGWorksheet.Activate();

                btnCalcLabournes.Enabled = true;

                btnLoadFromModel.Enabled = true;
                menuSection.Enabled = true;
                btnCreateTemplateFile.Enabled = true;

                groupCommands.Visible = true;
                labelCurrentEmployerName.Label = $"ОБЩИЕ ДАННЫЕ";
            }

            catch (Exception exp)
            {
                MessageBox.Show($"Ошибка при попытке показать общую ведомость. Ошибка: {exp.Message}");
            }

        }
        private void btnCalcLabournes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                CurrentMSGExellModel.CalcLabourness();
            }
            catch (Exception exp)
            {
                MessageBox.Show($"Ошибка при попытке персчета трудоемкостей. Ошибка: {exp.Message}");
            }
        }


        private void buttonCalc_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                CurrentMSGExellModel.SetFormulas();
                CurrentMSGExellModel.CalcAll();
                CurrentMSGExellModel.SetStyleFormats();
            }
            catch (Exception exp)
            {
                MessageBox.Show($"Ошибка при попытке персчета всех полей. Ошибка: {exp.Message}");
            }
            CurrentMSGExellModel.CalcAll();
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
            try
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
                    ////////////////////
                    Excel.Worksheet employer_worker_consumption_worksheet = CurrentWorkbook.Worksheets.Add(CommonWorkConsumptionsWorksheet, Type.Missing, Type.Missing, Type.Missing);

                    string work_consumptions_worksheet_name = CommonWorkConsumptionsWorksheet.Name.Substring(0, CommonWorkConsumptionsWorksheet.Name.IndexOf('_') + 1) + SelectedEmloeyer.Number.ToString();
                    employer_worker_consumption_worksheet.Name = work_consumptions_worksheet_name;

                    last_source = CommonWorkConsumptionsWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                    source = CommonWorkConsumptionsWorksheet.Range[CommonWorkConsumptionsWorksheet.Cells[1, 1], last_source];
                    source.Copy();

                    last_dest = employer_worker_consumption_worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                    dest = employer_worker_consumption_worksheet.Range[employer_worker_consumption_worksheet.Cells[1, 1], last_dest];
                    dest.PasteSpecial(XlPasteType.xlPasteAll);
                    Excel.Range comsup_day_range = employer_worker_consumption_worksheet.Range[
                        employer_worker_consumption_worksheet.Cells[MSGExellModel.W_CONSUMPTIONS_DATE_RAW + 1, MSGExellModel.W_CONSUMPTIONS_FIRST_DATE_COL],
                         employer_worker_consumption_worksheet.Cells[MSGExellModel.W_CONSUMPTIONS_DATE_RAW + 20, 5000]];
                    comsup_day_range.ClearContents();
                    EmployerWorkConsumptionsWorksheets.Add(employer_worker_consumption_worksheet);
                    var start_date_cell = CommonMSGWorksheet.Cells[MSGExellModel.WORKS_START_DATE_ROW, MSGExellModel.WORKS_START_DATE_COL];
                    string date_formula = $"={CommonMSGWorksheet.Name}!{Func.RangeAddress(start_date_cell)}";

                    Excel.Range cons_start_date_cell = employer_worker_consumption_worksheet.Cells[MSGExellModel.W_CONSUMPTIONS_DATE_RAW, MSGExellModel.W_CONSUMPTIONS_FIRST_DATE_COL];
                    cons_start_date_cell.Formula = date_formula;
                    new_employer_worksheet.Cells[MSGExellModel.WORKS_START_DATE_ROW, MSGExellModel.WORKS_START_DATE_COL]
                        .Formula = date_formula;

                    ///////////////////
                    Excel.Worksheet employer_machine_consumption_worksheet = CurrentWorkbook.Worksheets.Add(CommonMachineConsumptionsWorksheet, Type.Missing, Type.Missing, Type.Missing);

                    string machine_consumptions_worksheet_name = CommonMachineConsumptionsWorksheet.Name.Substring(0, CommonMachineConsumptionsWorksheet.Name.IndexOf('_') + 1) + SelectedEmloeyer.Number.ToString();
                    employer_machine_consumption_worksheet.Name = machine_consumptions_worksheet_name;

                    last_source = CommonMachineConsumptionsWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                    source = CommonMachineConsumptionsWorksheet.Range[CommonMachineConsumptionsWorksheet.Cells[1, 1], last_source];
                    source.Copy();

                    last_dest = employer_machine_consumption_worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                    dest = employer_machine_consumption_worksheet.Range[employer_machine_consumption_worksheet.Cells[1, 1], last_dest];
                    dest.PasteSpecial(XlPasteType.xlPasteAll);
                    Excel.Range machine_comsup_day_range = employer_machine_consumption_worksheet.Range[
                        employer_machine_consumption_worksheet.Cells[MSGExellModel.MCH_CONSUMPTIONS_DATE_RAW + 1, MSGExellModel.MCH_CONSUMPTIONS_FIRST_DATE_COL],
                         employer_machine_consumption_worksheet.Cells[MSGExellModel.MCH_CONSUMPTIONS_DATE_RAW + 20, 5000]];
                    machine_comsup_day_range.ClearContents();
                    MachineConsumptionsWorksheets.Add(employer_machine_consumption_worksheet);

                    Excel.Range machine_cons_start_date_cell = employer_machine_consumption_worksheet.Cells[MSGExellModel.MCH_CONSUMPTIONS_DATE_RAW, MSGExellModel.MCH_CONSUMPTIONS_FIRST_DATE_COL];
                    cons_start_date_cell.Formula = date_formula;

                    ///////////////////

                    this.ReloadAllModels();
                    empl_model = MSGExellModels.FirstOrDefault(m => m.Employer.Name == SelectedEmloeyer.Name);
                    //  empl_model.ClearWorksheetDaysPart();
                }
                CurrentMSGExellModel = empl_model;
                //   CurrentMSGExellModel.ReloadSheetModel();
                //   CurrentMSGExellModel.SetStyleFormats();


                this.SetAllWorksheetsVisibleState(XlSheetVisibility.xlSheetHidden);
                this.ShowWorksheet(empl_model.RegisterSheet);
                this.ShowWorksheet(empl_model.WorkerConsumptionsSheet);
                this.ShowWorksheet(empl_model.MachineConsumptionsSheet);
                empl_model.RegisterSheet.Activate();

                labelCurrentEmployerName.Label = $"ОТВЕСТВЕННЫЙ: {empl_model.Employer.Name}";
                menuSection.Enabled = false;
                btnCreateTemplateFile.Enabled = false;

                groupCommands.Visible = false;
                //       CurrentMSGExellModel.ResetCalculatesFields();
            }
            catch (Exception exp)
            {
                MessageBox.Show($"Ошибка при попытке перехода к данным отвественного. Ошибка: {exp.Message}");
            }
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
        private void btnMachines_Click(object sender, RibbonControlEventArgs e)
        {
            this.SetAllWorksheetsVisibleState(XlSheetVisibility.xlSheetHidden);
            this.ShowWorksheet(MachinesWorksheet);
            labelCurrentEmployerName.Label = $"МАШИНЫ И МЕХАНИЗМЫ";
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
            MachinesWorksheet.Visible = visibility;

            foreach (Excel.Worksheet worksheet in EmployerMSGWorksheets)
                worksheet.Visible = visibility;

            foreach (Excel.Worksheet worksheet in EmployerWorkConsumptionsWorksheets)
                worksheet.Visible = visibility;

            foreach (Excel.Worksheet worksheet in MachineConsumptionsWorksheets)
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
        private void ReloadMachinesList()
        {
            Machines.Clear();
            //   ObservableCollection<Machine> MachinesList = new ObservableCollection<Machine>();
            int row_index = 2;
            while (MachinesWorksheet.Cells[row_index, MACHINE_NUMBER_COL].Value != null)
            {
                string number = MachinesWorksheet.Cells[row_index, MACHINE_NUMBER_COL].Value.ToString();
                string name = MachinesWorksheet.Cells[row_index, MACHINE_NAME_COL].Value.ToString();
                Machines.Add(new Machine(number, name));
                row_index++;
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
            CommonMSGExellModel.MachineConsumptionsSheet = CommonMachineConsumptionsWorksheet;
            CommonMSGExellModel.CommonSheet = CommonWorksheet;
            CommonMSGExellModel.UnitOfMeasurements = UnitOfMeasurements;
            CommonMSGExellModel.IsHasEnabled = chckBoxHashEnable.Checked;

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

            foreach (Excel.Worksheet worksheet in MachineConsumptionsWorksheets)
            {
                string emoloyer_namber_str = worksheet.Name.Substring(worksheet.Name.LastIndexOf("_") + 1, worksheet.Name.Length - worksheet.Name.LastIndexOf("_") - 1);
                Employer employer = Employers.Where(em => em.Number == emoloyer_namber_str).FirstOrDefault();
                var model = MSGExellModels.FirstOrDefault(m => m.Employer.Number == emoloyer_namber_str);
                if (model != null && worksheet.Name.Contains("Техника"))
                    model.MachineConsumptionsSheet = worksheet;
            }

            CommonMSGExellModel.ReloadSheetModel();


        }
        private void btnLoadInModel_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                CurrentMSGExellModel.ReloadSheetModel();
                //CurrentMSGExellModel.SetFormulas();
                CurrentMSGExellModel.SetStyleFormats();
            }
            catch (Exception exp)
            {
                MessageBox.Show($"Ошибка при загрузке данных из документа в модель. Ошибка:{exp.Message}");
            }
        }

        private void btnLoadFromModel_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (CurrentMSGExellModel.Owner != null)
                {
                    CurrentMSGExellModel.ClearAllSections();
                    CurrentMSGExellModel.CopyOwnerObjectModels();
                    CurrentMSGExellModel.ReloadSheetModel();

                }
                //  else
                {
                    CurrentMSGExellModel.UpdateExcelRepresetation();
                    CurrentMSGExellModel.SetFormulas();
                    CurrentMSGExellModel.SetStyleFormats();
                }

            }
            catch (Exception exp)
            {
                MessageBox.Show($"Ошибка при выгрузке данных из модели в документ. Ошибка:{exp.Message}");
            }
        }
        private void btnUpdateAll_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (CurrentMSGExellModel.Owner == null)
                    foreach (MSGExellModel model in CurrentMSGExellModel.Children)
                        model.UpdateExcelRepresetation();

                CurrentMSGExellModel.ReloadSheetModel();
                CurrentMSGExellModel.UpdateExcelRepresetation();
                CurrentMSGExellModel.SetFormulas();
                CurrentMSGExellModel.SetStyleFormats();
            }
            catch (Exception exp)
            {
                MessageBox.Show($"Ошибка при полном обновлении всех моделей. Ошибка:{exp.Message}");
            }
        }
        #region Выгрузка данных в файл МСГ шаблона

        private void btnLoadTeplateFile_Click(object sender, RibbonControlEventArgs e)
        {
            //  string solutionpath = Directory.GetParent(Globals.ThisAddIn.Application.Path).Parent.Parent.Parent.FullName; 
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog
                {
                    InitialDirectory = @"С:\",
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
                    CurrentMSGExellModel.ReloadSheetModel();
                    CurrentMSGExellModel.CalcAll();
                    MSGTemplateWorkbook = Globals.ThisAddIn.Application.Workbooks.Open(temlate_file_name);
                    MSGTemplateWorkbook.Activate();
                    if (CommonMSGExellModel != null)
                        this.FillMSG_OUT_File(CurrentMSGExellModel, (w) => { return true; });

                    MSGTemplateWorkbook.SaveAs($"{MSGTemplateWorkbook.Path}\\{CurrentMSGExellModel.ContractCode}.xlsx");
                    MSGTemplateWorkbook.Close();
                }
            }
            catch (Exception exp)
            {

                MessageBox.Show($"Ошибка при выводе данных в шаблом графика МСГ. Ошибка:{exp.Message}");
            }
        }

        private void btnFillTemlate_Click(object sender, RibbonControlEventArgs e)
        {

            //     if (CommonMSGExellModel != null)
            //      this.FillMSG_OUT_File(CommonMSGExellModel);
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

        const int NEEDS_MACHINE_FIRST_ROW = 36;


        private void FillMSG_OUT_File(MSGExellModel curren_model, Func<MSGWork, bool> selection_predicate)
        {


            int row_index = TMP_WORK_SELECTION_FIRST_ROW;
            const int PLAN_PERIOD_MANTHS_NUMBER = 1;

            int work_needs_iterator = 0;

            Excel.Worksheet MSGOutWorksheet = MSGTemplateWorkbook.Worksheets["МСГ"];
            Excel.Worksheet MSGNeedsOutWorksheet = MSGTemplateWorkbook.Worksheets["Людские, технические ресурсы"];
            Excel.Worksheet MSGTemplateWorksheet = MSGTemplateWorkbook.Worksheets["МСГ_Шаблон"];
            Excel.Worksheet MSGNeedsTemplateWorksheet = MSGTemplateWorkbook.Worksheets["Людские_тех_ресурсы_Шаблон"];

            // MSGOutWorksheet.Visible = XlSheetVisibility.xlSheetHidden;
            //   MSGNeedsOutWorksheet.Visible = XlSheetVisibility.xlSheetHidden;
            MSGTemplateWorksheet.Visible = XlSheetVisibility.xlSheetHidden;
            MSGNeedsTemplateWorksheet.Visible = XlSheetVisibility.xlSheetHidden;
            DateTime current_day_date = DateTime.Now;

            MSGOutWorksheet.Cells[TMP_NOW_DATE_ROW, TMP_NOW_DATE_COL] = current_day_date.ToString("d");
            MSGOutWorksheet.Cells[TMP_CONTRACT_CODE_ROW, TMP_COMMON_PARAMETRS_VALUE_COL] = curren_model.ContractCode;
            MSGOutWorksheet.Cells[TMP_CONSTRUCTION_OBJECT_CODE_ROW, TMP_COMMON_PARAMETRS_VALUE_COL] = curren_model.ConstructionSubObjectCode;

            MSGNeedsTemplateWorksheet.Cells[NEEDS_NOW_DATE_ROW, NEEDS_NOW_DATE_COL].Value = current_day_date.ToString("d");
            MSGNeedsTemplateWorksheet.Cells[NEEDS_CONTRACT_CODE_ROW, NEEDS_NOW_DATE_COL] = curren_model.ContructionObjectCode;
            MSGNeedsTemplateWorksheet.Cells[NEEDS_CONSTRUCTION_OBJECT_CODE_ROW, NEEDS_NOW_DATE_COL] = curren_model.ConstructionSubObjectCode;

            if (checkBoxRerightDatePart.Checked)
                this.FillMSG_OUT_File_Headers(curren_model);

            int date_col_index = 0;
            int in_worksheet_number = 18;
            int worked_days_number = (curren_model.WorksEndDate - curren_model.WorksStartDate).Days;
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
            Dictionary<DateTime, decimal> lobournes_coefficents = new Dictionary<DateTime, decimal>();

            if (curren_model.Owner != null)
            {
                for (DateTime date = curren_model.Owner.WorksStartDate; date <= curren_model.WorksEndDate; date = date.AddDays(1))
                {
                    decimal common_workers_number = 0;
                    decimal current_workers_number = 0;
                    foreach (var w_cons in curren_model.Owner.WorkerConsumptions.Where(wc => wc.Name != "ИТР"))
                    {
                        var consumtion_report_card = w_cons.WorkersConsumptionReportCard.Where(rc => rc.Date == date);
                        foreach (var c in consumtion_report_card)
                            common_workers_number += c.Quantity;
                    }
                    foreach (var w_cons in curren_model.WorkerConsumptions.Where(wc => wc.Name != "ИТР"))
                    {
                        var consumtion_report_card = w_cons.WorkersConsumptionReportCard.Where(rc => rc.Date == date);
                        foreach (var c in consumtion_report_card)
                            current_workers_number += c.Quantity;
                    }

                    decimal w_coefficent = 0;
                    if (common_workers_number != 0)
                        w_coefficent = current_workers_number / common_workers_number;

                    if (!lobournes_coefficents.ContainsKey(date))
                        lobournes_coefficents.Add(date, w_coefficent);
                }
            }



            foreach (WorksSection w_section in curren_model.WorksSections)
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
                        string w_section_number = w_section_name.Split(' ')[0];
                        if (w_section_number == w_section.Number)
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
                MSGOutWorksheet.Cells[section_local_index_iterator, 1] = $"{w_section.Number} {w_section.Name}";

                saved_iterator = section_local_index_iterator + 1;
                var works = w_section.MSGWorks.Where(w => selection_predicate(w));

                foreach (MSGWork msg_work in w_section.MSGWorks.Where(w => selection_predicate(w)))
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



                            if (curren_model.WorksSections.FirstOrDefault(wc => wc.Name == msg_work_number) != null
                                                 && msg_work_number != w_section.Name)
                            {
                                saved_iterator = work_local_index_iterator;
                                tmp_work_rows_range.Copy();
                                Excel.Range dest = MSGOutWorksheet.Rows[saved_iterator];
                                dest.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Type.Missing);

                                break;
                            }
                            if (msg_work.Number == msg_work_number)
                            {
                                saved_iterator = work_local_index_iterator;
                                break;
                            }
                            if (msg_work.Number != msg_work_number && msg_work_number != "")
                            {
                                saved_iterator = work_local_index_iterator + 2;
                            }
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

                    //decimal project_quantity = 0;
                    //if (curren_model.Owner == null)
                    //    project_quantity = msg_work.ProjectQuantity;
                    //else
                    //{
                    //    var owner_model = curren_model.Owner;

                    //}
                    MSGOutWorksheet.Cells[row_index, TMP_WORK_PROJECT_QUANTITY_COL] = msg_work.ProjectQuantity;
                    MSGOutWorksheet.Cells[row_index, TMP_U_MRASURE_COL] = msg_work.UnitOfMeasurement.Name;

                    MSGOutWorksheet.Cells[row_index, TMP_WORK_START_DATE_COL] = msg_work.WorkSchedules.StartDate;
                    MSGOutWorksheet.Cells[row_index, TMP_WORK_END_DATE_COL] = msg_work.WorkSchedules.EndDate;

                    MSGOutWorksheet.Cells[row_index + 1, TMP_PREVIOUS_WORK_QUANTITY_COL] = msg_work.PreviousComplatedQuantity;
                    // MSGOutWorksheet.Cells[row_index, TMP_WORK_DAYS_NUMBER_COL] = (msg_work.WorkSchedules.EndDate - msg_work.WorkSchedules.StartDate)?.Days;
                    ///Заполняем плановые объемы в календарной части
                    MSGExellModel owner_model = curren_model.Owner;

                    int? workable_days_num = msg_work.WorkSchedules.GetShedulesAllDaysNumber();
                    foreach (WorkScheduleChunk schedule_chunk in msg_work.WorkSchedules)
                    {
                        int date_index = 0;
                        while (MSGOutWorksheet.Cells[TMP_WORKDAY_DATE_ROW_COL, TMP_WORKDAY_DATE_FIRST_COL + date_index].Value != null && date_index < last_day_col_index)
                        {
                            DateTime date;
                            DateTime.TryParse(MSGOutWorksheet.Cells[TMP_WORKDAY_DATE_ROW_COL, TMP_WORKDAY_DATE_FIRST_COL + date_index].Value.ToString(), out date);

                            if (date >= schedule_chunk.StartTime && date <= schedule_chunk.EndTime
                                && (date.DayOfWeek != DayOfWeek.Sunday || schedule_chunk.IsSundayVacationDay == "Нет"))
                            {
                                decimal? project_quantity = 0;
                                if (curren_model.Owner != null)
                                {
                                    project_quantity = (msg_work.ProjectQuantity - msg_work.PreviousComplatedQuantity) / workable_days_num;
                                    project_quantity = project_quantity * lobournes_coefficents[date];
                                }
                                else
                                    project_quantity = (msg_work.ProjectQuantity - msg_work.PreviousComplatedQuantity) / workable_days_num;

                                MSGOutWorksheet.Cells[row_index, TMP_WORKDAY_DATE_FIRST_COL + date_index] = project_quantity;
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
                var current_needs_of_worker = curren_model.WorkersComposition.FirstOrDefault(nw => nw.Name == worker_post_name);
                var current_worker_consumption = curren_model.WorkerConsumptions.FirstOrDefault(wc => wc.Name == worker_post_name);

                while (work_needs_date_col_index < last_day_col_index)
                {
                    if (current_needs_of_worker != null)
                    {
                        NeedsOfWorkersDay needsOfWorkersDay = current_needs_of_worker.NeedsOfWorkersReportCard.FirstOrDefault(nwd => nwd.Date == current_date);
                        if (needsOfWorkersDay != null)
                        {
                            if (curren_model.Owner != null)
                            {
                                MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator,
                                NEEDS_WORKDAY_DATE_FIRST_COL + work_needs_date_col_index] = needsOfWorkersDay.Quantity * lobournes_coefficents[current_date];
                            }
                            else
                                MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator,
                                NEEDS_WORKDAY_DATE_FIRST_COL + work_needs_date_col_index] = needsOfWorkersDay.Quantity;
                        }
                    }
                    if (current_worker_consumption != null)
                    {
                        WorkerConsumptionDay worker_consumption_day = current_worker_consumption.WorkersConsumptionReportCard.FirstOrDefault(wcd => wcd.Date == current_date);
                        if (worker_consumption_day != null)
                        {
                            if (curren_model.Owner != null)
                            {
                                MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator + 1,
                                NEEDS_WORKDAY_DATE_FIRST_COL + work_needs_date_col_index] = worker_consumption_day.Quantity * lobournes_coefficents[current_date]; ;
                            }
                            else
                                MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator + 1,
                               NEEDS_WORKDAY_DATE_FIRST_COL + work_needs_date_col_index] = worker_consumption_day.Quantity;
                        }
                    }
                    work_needs_date_col_index++;
                    if (MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW, NEEDS_WORKDAY_DATE_FIRST_COL + work_needs_date_col_index].Value == null) break;
                    DateTime.TryParse(MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW, NEEDS_WORKDAY_DATE_FIRST_COL + work_needs_date_col_index].Value.ToString(), out current_date);
                }

                work_needs_iterator++;
            }

            int machine_needs_iterator = 0;
            while (MSGNeedsOutWorksheet.Cells[NEEDS_MACHINE_FIRST_ROW + machine_needs_iterator, NEEDS_WORKERS_NAME_COL].Value != "Итого")
            {
                int machine_needs_date_col_index = 1;
                DateTime current_date;
                var ss = MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW, NEEDS_WORKDAY_DATE_FIRST_COL + machine_needs_date_col_index].Value;
                DateTime.TryParse(MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW, NEEDS_WORKDAY_DATE_FIRST_COL + machine_needs_date_col_index].Value.ToString("d"), out current_date);
                string worker_post_name = MSGNeedsOutWorksheet.Cells[NEEDS_MACHINE_FIRST_ROW + machine_needs_iterator, NEEDS_WORKERS_NAME_COL].Value;
                var current_needs_of_worker = curren_model.MachinesComposition.FirstOrDefault(nw => nw.Name == worker_post_name);
                var current_machine_consumption = curren_model.MachineConsumptions.FirstOrDefault(wc => wc.Name == worker_post_name);

                while (machine_needs_date_col_index < last_day_col_index)
                {
                    if (current_needs_of_worker != null)
                    {
                        NeedsOfMachineDay needsOfMachinesDay = current_needs_of_worker.NeedsOfMachinesReportCard.FirstOrDefault(nwd => nwd.Date == current_date);
                        if (needsOfMachinesDay != null)
                            MSGNeedsOutWorksheet.Cells[NEEDS_MACHINE_FIRST_ROW + machine_needs_iterator,
                                NEEDS_WORKDAY_DATE_FIRST_COL + machine_needs_date_col_index] = needsOfMachinesDay.Quantity;
                    }
                    if (current_machine_consumption != null)
                    {
                        MachineConsumptionDay machine_consumption_day = current_machine_consumption.MachinesConsumptionReportCard.FirstOrDefault(wcd => wcd.Date == current_date);
                        if (machine_consumption_day != null)
                            MSGNeedsOutWorksheet.Cells[NEEDS_MACHINE_FIRST_ROW + machine_needs_iterator + 1,
                                NEEDS_WORKDAY_DATE_FIRST_COL + machine_needs_date_col_index] = machine_consumption_day.Quantity;
                    }
                    machine_needs_date_col_index++;
                    if (MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW, NEEDS_WORKDAY_DATE_FIRST_COL + machine_needs_date_col_index].Value == null) break;
                    DateTime.TryParse(MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW, NEEDS_WORKDAY_DATE_FIRST_COL + machine_needs_date_col_index].Value.ToString(), out current_date);
                }
                var dd = MSGNeedsOutWorksheet.Cells[NEEDS_MACHINE_FIRST_ROW + machine_needs_iterator, NEEDS_WORKERS_NAME_COL].Value;
                machine_needs_iterator++;
            }

            MSGOutWorksheet.Visible = XlSheetVisibility.xlSheetVisible;
            MSGNeedsOutWorksheet.Visible = XlSheetVisibility.xlSheetVisible;
            // MSGTemplateWorksheet.Visible = XlSheetVisibility.xlSheetVisible;
            // MSGNeedsTemplateWorksheet.Visible = XlSheetVisibility.xlSheetVisible;

            #endregion


        }

        private void FillMSG_OUT_File_Headers(MSGExellModel curren_model)
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
                for (DateTime date = curren_model.WorksStartDate; date <= curren_model.WorksEndDate; date = date.AddDays(1))
                {

                    ///Если текущий день является восскесеньем или является первым днем все ведомости -
                    /// втавляем и заполняем недельный столбец в календарь
                    if (date.DayOfWeek == DayOfWeek.Monday || date == curren_model.WorksStartDate)
                    {
                        #region Календраня часть МСГ недельный столбец

                        if (week_signatura_last_col > 0)
                        {
                            //       if (week_signatura_last_col > 1) week_signatura_last_col++;
                            Excel.Range week_name_range_first_cell = MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW - 1, NEEDS_WORKDAY_DATE_FIRST_COL + week_signatura_first_col];
                            Excel.Range week_name_range_last_cell = MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW - 1, NEEDS_WORKDAY_DATE_FIRST_COL + week_signatura_last_col];
                            Excel.Range week_name_range = MSGNeedsOutWorksheet.get_Range(week_name_range_first_cell, week_name_range_last_cell);
                            week_name_range.Merge();
                            week_name_range_first_cell.Value = last_week_name_signatura.Replace("\r\n", " ");
                        }


                        ///Вставляем из шаблона МСГ недельный столбец в каледаре ( копируем из шаблона)
                        tmp_week_col_range.Copy();
                        Excel.Range week_day_dest = MSGOutWorksheet.UsedRange.Columns[WORKDAY_DATE_FIRST_COL + date_col_index];
                        week_day_dest.PasteSpecial(XlPasteType.xlPasteAll);

                        ///Заполняем из шаблона МСГ недельный столбец в каледаре 
                        string week_name_signatura =
                          $"неделя\r\n {date.ToString("dd")} - {this.GetLastWeekdayDate(date).ToString("dd")}";
                        MSGOutWorksheet.Cells[WORKDAY_DATE_ROW, WORKDAY_DATE_FIRST_COL + date_col_index] = week_name_signatura;
                        MSGOutWorksheet.Cells[WORKDAY_DATE_ROW + 1, WORKDAY_DATE_FIRST_COL + date_col_index] = in_worksheet_number++;


                        first_week_day_col = date_col_index + 1;
                        last_week_day_col = first_week_day_col + (this.GetLastWeekdayDate(date) - date).Days;

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

                        work_needs_iterator = 0;
                        while (MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator, NEEDS_WORKERS_NAME_COL].Value != "Общее количество")
                        {

                            Excel.Range needs_first_day = MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator, NEEDS_WORKDAY_DATE_FIRST_COL + first_week_day_col];
                            Excel.Range needs_last_day = MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator, NEEDS_WORKDAY_DATE_FIRST_COL + last_week_day_col];
                            MSGNeedsOutWorksheet.Cells[NEEDS_WORKERS_FIRST_ROW + work_needs_iterator, NEEDS_WORKDAY_DATE_FIRST_COL + date_col_index] =
                               $"=SUM({Func.RangeAddress(needs_first_day)}:{Func.RangeAddress(needs_last_day)})"; ;
                            work_needs_iterator++;
                        }
                        work_needs_iterator = 0;
                        while (MSGNeedsOutWorksheet.Cells[NEEDS_MACHINE_FIRST_ROW + work_needs_iterator, NEEDS_WORKERS_NAME_COL].Value != "Итого")
                        {

                            Excel.Range needs_first_day = MSGNeedsOutWorksheet.Cells[NEEDS_MACHINE_FIRST_ROW + work_needs_iterator, NEEDS_WORKDAY_DATE_FIRST_COL + first_week_day_col];
                            Excel.Range needs_last_day = MSGNeedsOutWorksheet.Cells[NEEDS_MACHINE_FIRST_ROW + work_needs_iterator, NEEDS_WORKDAY_DATE_FIRST_COL + last_week_day_col];
                            MSGNeedsOutWorksheet.Cells[NEEDS_MACHINE_FIRST_ROW + work_needs_iterator, NEEDS_WORKDAY_DATE_FIRST_COL + date_col_index] =
                               $"=SUM({Func.RangeAddress(needs_first_day)}:{Func.RangeAddress(needs_last_day)})"; ;
                            work_needs_iterator++;
                        }

                        #endregion


                        Excel.Range week_range = MSGOutWorksheet.Range[MSGOutWorksheet.Columns[WORKDAY_DATE_FIRST_COL + first_week_day_col],
                           MSGOutWorksheet.Columns[WORKDAY_DATE_FIRST_COL + last_week_day_col]];

                        Excel.Range needs_week_range = MSGNeedsOutWorksheet.Range[MSGNeedsOutWorksheet.Columns[NEEDS_WORKDAY_DATE_FIRST_COL + first_week_day_col],
                            MSGNeedsOutWorksheet.Columns[NEEDS_WORKDAY_DATE_FIRST_COL + last_week_day_col]];
                        try
                        {
                            week_range.Group();
                            needs_week_range.Group();
                        }
                        catch (Exception exp)
                        {
                            throw new Exception($"Ошибка при заполении заголовков потребностей в шаблоне МСГ графика.{week_range.ToString()},{needs_week_range.ToString()}.Ошибка:{exp.Message}");

                        }

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
                    if (date.DayOfWeek == DayOfWeek.Sunday)
                        MSGNeedsOutWorksheet.Cells[NEEDS_WORKDAY_DATE_ROW, NEEDS_WORKDAY_DATE_FIRST_COL + date_col_index].Interior.Color = XlRgbColor.rgbOrangeRed;
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
        private DateTime GetLastWeekdayDate(DateTime date)
        {
            DateTime out_date = date;
            if (out_date.DayOfWeek == DayOfWeek.Sunday) return date;
            while (out_date.AddDays(1).DayOfWeek != DayOfWeek.Monday)
                out_date = out_date.AddDays(1);
            return out_date;
        }
        #endregion

        private void checkBoxRerightDatePart_Click(object sender, RibbonControlEventArgs e)
        {

        }

        #region Редактивроние данных
        private List<IExcelBindableBase> CopyedObjectsList = new List<IExcelBindableBase>();
        private string commands_group_label = "";

        /// <summary>
        /// Допивать в пустые МСГ все ВОВР, КС и табельные работы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnInitMSGContent_Click(object sender, RibbonControlEventArgs e)
        {
            var selection = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            var sected_object = CurrentMSGExellModel.GetObjectsBySelection(selection, typeof(MSGWork));

            foreach (MSGWork msg_work in sected_object)
            {
                if (msg_work.VOVRWorks.Count == 0)
                {
                    VOVRWork vovr_work = new VOVRWork();
                    vovr_work.Worksheet = CurrentMSGExellModel.RegisterSheet;
                    vovr_work.Number = $"{msg_work.Number}.1";
                    vovr_work.Name = msg_work.Name;
                    vovr_work.UnitOfMeasurement = msg_work.UnitOfMeasurement;
                    vovr_work.ProjectQuantity = msg_work.ProjectQuantity;
                    vovr_work.Laboriousness = msg_work.Laboriousness;
                    int rowIndex = msg_work.CellAddressesMap["Number"].Row;
                    CurrentMSGExellModel.Register(vovr_work, "Number", rowIndex, MSGExellModel.VOVR_NUMBER_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(vovr_work, "Name", rowIndex, MSGExellModel.VOVR_NAME_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(vovr_work, "ProjectQuantity", rowIndex, MSGExellModel.VOVR_QUANTITY_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(vovr_work, "Quantity", rowIndex, MSGExellModel.VOVR_QUANTITY_FACT_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(vovr_work, "Laboriousness", rowIndex, MSGExellModel.VOVR_LABOURNESS_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(vovr_work, "UnitOfMeasurement.Name", rowIndex, MSGExellModel.VOVR_MEASURE_COL, CurrentMSGExellModel.RegisterSheet);

                    KSWork ks_work = new KSWork();
                    ks_work.Worksheet = CurrentMSGExellModel.RegisterSheet;
                    ks_work.Number = $"{vovr_work.Number}.1";
                    ks_work.Code = "-";
                    ks_work.Name = msg_work.Name;
                    ks_work.UnitOfMeasurement = msg_work.UnitOfMeasurement;
                    ks_work.ProjectQuantity = msg_work.ProjectQuantity;
                    ks_work.Laboriousness = msg_work.Laboriousness;
                    CurrentMSGExellModel.Register(ks_work, "Number", rowIndex, MSGExellModel.KS_NUMBER_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(ks_work, "Code", rowIndex, MSGExellModel.KS_CODE_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(ks_work, "Name", rowIndex, MSGExellModel.KS_NAME_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(ks_work, "ProjectQuantity", rowIndex, MSGExellModel.KS_QUANTITY_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(ks_work, "Quantity", rowIndex, MSGExellModel.KS_QUANTITY_FACT_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(ks_work, "Laboriousness", rowIndex, MSGExellModel.KS_LABOURNESS_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(ks_work, "UnitOfMeasurement.Name", rowIndex, MSGExellModel.KS_MEASURE_COL, CurrentMSGExellModel.RegisterSheet);

                    RCWork rc_work = new RCWork();
                    rc_work.Worksheet = CurrentMSGExellModel.RegisterSheet;
                    rc_work.Number = $"{ks_work.Number}.1";
                    rc_work.Code = "-";
                    rc_work.Name = msg_work.Name;
                    rc_work.UnitOfMeasurement = msg_work.UnitOfMeasurement;
                    rc_work.ProjectQuantity = msg_work.ProjectQuantity;
                    rc_work.Laboriousness = msg_work.Laboriousness;
                    CurrentMSGExellModel.Register(rc_work, "Number", rowIndex, MSGExellModel.RC_NUMBER_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(rc_work, "Code", rowIndex, MSGExellModel.RC_CODE_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(rc_work, "Name", rowIndex, MSGExellModel.RC_NAME_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(rc_work, "ProjectQuantity", rowIndex, MSGExellModel.RC_QUANTITY_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(rc_work, "Quantity", rowIndex, MSGExellModel.RC_QUANTITY_FACT_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(rc_work, "LabournessCoefficient", rowIndex, MSGExellModel.RC_LABOURNESS_COEFFICIENT_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(rc_work, "Laboriousness", rowIndex, MSGExellModel.RC_LABOURNESS_COL, CurrentMSGExellModel.RegisterSheet);
                    CurrentMSGExellModel.Register(rc_work, "UnitOfMeasurement.Name", rowIndex, MSGExellModel.RC_MEASURE_COL, CurrentMSGExellModel.RegisterSheet);

                    msg_work.VOVRWorks.Add(vovr_work);
                    vovr_work.KSWorks.Add(ks_work);
                    ks_work.RCWorks.Add(rc_work);

                    msg_work.AdjustExcelRepresentionTree(rowIndex);
                    msg_work.UpdateExcelRepresetation();
                    msg_work.SetStyleFormats(MSGExellModel.W_SECTION_COLOR + 1);
                }
            }
        }

        /// <summary>
        /// Вставить из беферного списка объекты
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <exception cref="Exception"></exception>
        private void buttonPaste_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var selection = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                //var sected_object = CurrentMSGExellModel.GetObjectBySelection(selection, typeof(WorksSection));
                //    var sected_objects = CurrentMSGExellModel.GetObjectsBySelection(selection, typeof(IExcelBindableBase));
                if (CopyedObjectsList.Count > 0)
                    switch (CopyedObjectsList[0]?.GetType().Name)
                    {
                        case nameof(WorksSection):
                            {

                                try
                                {
                                    if (selection.Column != MSGExellModel.WSEC_NUMBER_COL) return;
                                    int cell_val;
                                    Int32.TryParse(selection.Value.ToString(), out cell_val);
                                    int _serction_row = selection.Row;
                                    if (cell_val == 0) return;
                                    foreach (WorksSection section in CopyedObjectsList)
                                    {
                                        CurrentMSGExellModel.WorksSections.Add(section);
                                        section.SetNumberItem(0, cell_val.ToString());

                                        cell_val++;
                                    }
                                    CurrentMSGExellModel.UpdateExcelRepresetation();

                                    foreach (WorksSection section in CopyedObjectsList)
                                        CurrentMSGExellModel.RegisterObjectInObjectPropertyNameRegister(section);

                                    CurrentMSGExellModel.UpdateExcelRepresetation();
                                    CurrentMSGExellModel.SetStyleFormats();
                                    commands_group_label = "";
                                }
                                catch (Exception exp)
                                {
                                    throw new Exception($"Ошибка вставки раздела. {CopyedObjectsList.ToString()}. Ошибка:{exp.Message}");
                                }

                                break;
                            }
                        case nameof(MSGWork):
                            {
                                if (selection.Column <= MSGExellModel.MSG_NUMBER_COL | selection.Column >= MSGExellModel.MSG_NEEDS_OF_MACHINE_QUANTITY_COL) return;
                                var selected_above_msg_works = CurrentMSGExellModel
                                        .GetObjectsBySelection(selection, (obj) => obj is MSGWork msg_obj

                                                                                   && obj.Owner != null);

                                MSGWork _current_work = selected_above_msg_works.FirstOrDefault(ob => ob.GetTopRow() == selection.Row) as MSGWork;
                                MSGWork _next_work = selected_above_msg_works.FirstOrDefault(ob => ob.GetTopRow() > selection.Row) as MSGWork;
                                MSGWork _previous_work = selected_above_msg_works.FirstOrDefault(ob => ob.GetTopRow() < selection.Row) as MSGWork;
                                MSGWork selected_work;

                                if (_current_work != null) selected_work = _current_work;
                                else if (_next_work != null) selected_work = _next_work;
                                else
                                    selected_work = _previous_work;

                                if (selected_work != null)
                                {
                                    WorksSection selected_section = selected_work.Owner as WorksSection;
                                    if (selected_section == null) return;
                                    int selected_work_index = selected_section.MSGWorks.IndexOf(selected_work);

                                    foreach (MSGWork msg_work in CopyedObjectsList.OrderBy(ob => Int32.Parse(ob.NumberSuffix)))
                                    {
                                        selected_section.MSGWorks.Insert(selected_work_index, msg_work);
                                        selected_work_index++;
                                    }
                                    CurrentMSGExellModel.UpdateExcelRepresetation();
                                    foreach (MSGWork msg_work in CopyedObjectsList)
                                        CurrentMSGExellModel.RegisterObjectInObjectPropertyNameRegister(msg_work);

                                    CurrentMSGExellModel.SetStyleFormats();
                                    commands_group_label = "";
                                }

                                break;
                            }
                        case nameof(VOVRWork):
                            {
                                if (selection.Column <= MSGExellModel.VOVR_NUMBER_COL | selection.Column >= MSGExellModel.VOVR_LABOURNESS_COL) return;
                                var selected_above_msg_works = CurrentMSGExellModel
                                        .GetObjectsBySelection(selection, (obj) => obj is VOVRWork msg_obj

                                                                                   && obj.Owner != null);

                                VOVRWork _current_work = selected_above_msg_works.FirstOrDefault(ob => ob.GetTopRow() == selection.Row) as VOVRWork;
                                VOVRWork _next_work = selected_above_msg_works.FirstOrDefault(ob => ob.GetTopRow() > selection.Row) as VOVRWork;
                                VOVRWork _previous_work = selected_above_msg_works.FirstOrDefault(ob => ob.GetTopRow() < selection.Row) as VOVRWork;
                                VOVRWork selected_work;

                                if (_current_work != null) selected_work = _current_work;
                                else if (_next_work != null && _previous_work != null && _next_work.Owner == _previous_work.Owner)
                                    selected_work = _next_work;
                                else
                                    selected_work = _previous_work;

                                if (selected_work != null)
                                {
                                    MSGWork selected_msg_work = selected_work.Owner as MSGWork;
                                    if (selected_msg_work == null) return;
                                    int selected_work_index = selected_msg_work.VOVRWorks.IndexOf(selected_work);

                                    foreach (VOVRWork _work in CopyedObjectsList.OrderBy(ob => Int32.Parse(ob.NumberSuffix)))
                                    {
                                        selected_msg_work.VOVRWorks.Insert(selected_work_index, _work);
                                        selected_work_index++;
                                    }
                                    CurrentMSGExellModel.UpdateExcelRepresetation();
                                    foreach (VOVRWork msg_work in CopyedObjectsList)
                                        CurrentMSGExellModel.RegisterObjectInObjectPropertyNameRegister(msg_work);

                                    CurrentMSGExellModel.SetStyleFormats();
                                    commands_group_label = "";
                                }

                                break;
                            }
                        case nameof(KSWork):
                            {
                                if (selection.Column <= MSGExellModel.KS_NUMBER_COL | selection.Column >= MSGExellModel.KS_LABOURNESS_COL) return;
                                var selected_above_msg_works = CurrentMSGExellModel
                                        .GetObjectsBySelection(selection, (obj) => obj is KSWork msg_obj

                                                                                   && obj.Owner != null);

                                KSWork _current_work = selected_above_msg_works.FirstOrDefault(ob => ob.GetTopRow() == selection.Row) as KSWork;
                                KSWork _next_work = selected_above_msg_works.FirstOrDefault(ob => ob.GetTopRow() > selection.Row) as KSWork;
                                KSWork _previous_work = selected_above_msg_works.FirstOrDefault(ob => ob.GetTopRow() < selection.Row) as KSWork;
                                KSWork selected_work;

                                if (_current_work != null) selected_work = _current_work;
                                else if (_next_work != null && _previous_work != null && _next_work.Owner == _previous_work.Owner)
                                    selected_work = _next_work;
                                else
                                    selected_work = _previous_work;

                                if (selected_work != null)
                                {
                                    VOVRWork selected_vovr_work = selected_work.Owner as VOVRWork;
                                    if (selected_vovr_work == null) return;
                                    int selected_work_index = selected_vovr_work.KSWorks.IndexOf(selected_work);

                                    foreach (KSWork _work in CopyedObjectsList.OrderBy(ob => Int32.Parse(ob.NumberSuffix)))
                                    {
                                        selected_vovr_work.KSWorks.Insert(selected_work_index, _work);
                                        selected_work_index++;
                                    }
                                    CurrentMSGExellModel.UpdateExcelRepresetation();
                                    foreach (KSWork msg_work in CopyedObjectsList)
                                        CurrentMSGExellModel.RegisterObjectInObjectPropertyNameRegister(msg_work);

                                    CurrentMSGExellModel.SetStyleFormats();
                                    commands_group_label = "";
                                }
                                break;
                            }

                        case nameof(RCWork):
                            {
                                if (selection.Column <= MSGExellModel.RC_NUMBER_COL | selection.Column >= MSGExellModel.RC_LABOURNESS_COL) return;
                                var selected_above_msg_works = CurrentMSGExellModel
                                        .GetObjectsBySelection(selection, (obj) => obj is RCWork msg_obj

                                                                                   && obj.Owner != null);

                                RCWork _current_work = selected_above_msg_works.FirstOrDefault(ob => ob.GetTopRow() == selection.Row) as RCWork;
                                RCWork _next_work = selected_above_msg_works.FirstOrDefault(ob => ob.GetTopRow() > selection.Row) as RCWork;
                                RCWork _previous_work = selected_above_msg_works.FirstOrDefault(ob => ob.GetTopRow() < selection.Row) as RCWork;
                                RCWork selected_work;

                                if (_current_work != null) selected_work = _current_work;
                                else if (_next_work != null && _previous_work != null && _next_work.Owner == _previous_work.Owner)
                                    selected_work = _next_work;
                                else
                                    selected_work = _previous_work;

                                if (selected_work != null)
                                {
                                    KSWork selected_KS_work = selected_work.Owner as KSWork;
                                    if (selected_KS_work == null) return;
                                    int selected_work_index = selected_KS_work.RCWorks.IndexOf(selected_work);

                                    foreach (RCWork _work in CopyedObjectsList.OrderBy(ob => Int32.Parse(ob.NumberSuffix)))
                                    {
                                        selected_KS_work.RCWorks.Insert(selected_work_index, _work);
                                        selected_work_index++;
                                    }
                                    CurrentMSGExellModel.UpdateExcelRepresetation();
                                    foreach (RCWork msg_work in CopyedObjectsList)
                                        CurrentMSGExellModel.RegisterObjectInObjectPropertyNameRegister(msg_work);

                                    CurrentMSGExellModel.SetStyleFormats();
                                    commands_group_label = "";
                                }
                                break;
                            }
                        case nameof(NeedsOfWorker):
                            {
                                if (selection.Column <= MSGExellModel.MSG_NUMBER_COL | selection.Column >= MSGExellModel.MSG_NEEDS_OF_MACHINE_QUANTITY_COL) return;

                                var sected_object = CurrentMSGExellModel.GetObjectsBySelection(selection, typeof(MSGWork));

                                foreach (MSGWork msg_work in sected_object)
                                {

                                    foreach (NeedsOfWorker n_w in CopyedObjectsList)
                                    {
                                        var msg_n_w = msg_work.WorkersComposition.FirstOrDefault(n => n.Name == n_w.Name);
                                        if (msg_n_w != null)
                                            msg_n_w.Quantity = n_w.Quantity;
                                        else
                                            msg_work.WorkersComposition.Add(n_w.Clone() as NeedsOfWorker);
                                    }
                                }

                                CurrentMSGExellModel.UpdateExcelRepresetation();
                                foreach (MSGWork msg_work in sected_object)
                                    CurrentMSGExellModel.RegisterObjectInObjectPropertyNameRegister(msg_work);
                                CurrentMSGExellModel.SetStyleFormats();
                                break;
                            }
                        case nameof(NeedsOfMachine):
                            {
                                if (selection.Column <= MSGExellModel.MSG_NUMBER_COL | selection.Column >= MSGExellModel.MSG_NEEDS_OF_MACHINE_QUANTITY_COL) return;

                                var sected_object = CurrentMSGExellModel.GetObjectsBySelection(selection, typeof(MSGWork));

                                foreach (MSGWork msg_work in sected_object)
                                {
                                    foreach (NeedsOfMachine n_w in CopyedObjectsList)
                                    {
                                        var msg_n_w = msg_work.MachinesComposition.FirstOrDefault(n => n.Name == n_w.Name);
                                        if (msg_n_w != null)
                                            msg_n_w.Quantity = n_w.Quantity;
                                        else
                                            msg_work.MachinesComposition.Add(n_w.Clone() as NeedsOfMachine);
                                    }
                                }
                                CurrentMSGExellModel.UpdateExcelRepresetation();
                                foreach (MSGWork msg_work in sected_object)
                                    CurrentMSGExellModel.RegisterObjectInObjectPropertyNameRegister(msg_work);

                                CurrentMSGExellModel.SetStyleFormats();
                                break;
                            }


                        default:
                            {
                                break;
                            }
                    }

                CommonMSGExellModel.SetHashFormulas();
                groupCommands.Label = commands_group_label;

            }
            catch (Exception exp)
            {
                MessageBox.Show($"Ошибка при попытке всавки. Ошибка: {exp.Message}");
            }

        }

        private void btnCopyWorkerComposition_Click(object sender, RibbonControlEventArgs e)
        {
            CopyedObjectsList.Clear();
            var selection = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            var sected_objects = CurrentMSGExellModel.GetObjectsBySelection(selection, typeof(NeedsOfWorker));
            if (sected_objects != null)
            {
                foreach (var obj in sected_objects)
                    CopyedObjectsList.Add((NeedsOfWorker)obj.Clone()); ;

                buttonPaste.Enabled = true;
                commands_group_label = $"Рабники скопированы.\n Выбрерите разде куда вставить МСГ";

            }
            groupCommands.Label = commands_group_label;
        }

        private void btnCopyMachineComposition_Click(object sender, RibbonControlEventArgs e)
        {
            CopyedObjectsList.Clear();
            var selection = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            var sected_objects = CurrentMSGExellModel.GetObjectsBySelection(selection, typeof(NeedsOfMachine));
            if (sected_objects != null)
            {
                foreach (var obj in sected_objects)
                    CopyedObjectsList.Add((NeedsOfMachine)obj.Clone()); ;

                buttonPaste.Enabled = true;
                commands_group_label = $"Техника скопирована.\n Выбрерите разде куда вставить МСГ";

            }
            groupCommands.Label = commands_group_label;
        }
        #endregion
        /// <summary>
        /// Копировать разделы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonCopyWorkSection_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                CopyedObjectsList.Clear();
                var selection = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                var sected_objects = CurrentMSGExellModel.GetObjectsBySelection(selection, typeof(WorksSection)).Where(ob => !CopyedObjectsList.Contains(ob));
                if (sected_objects != null)
                {
                    foreach (var obj in sected_objects)
                        CopyedObjectsList.Add((WorksSection)obj.Clone()); ;

                    if (CopyedObjectsList.Count > 0)
                    {
                        buttonPaste.Enabled = true;
                        commands_group_label = $"Разадел(ы) скопирован.\n Выбрерите ячеку с номеров нового раздела.";
                    }

                }
                //Excel.Dialog dlg = Globals.ThisAddIn.Application.Dialogs[Excel.XlBuiltInDialog.xlDialogFont];
                //dlg.Show();
                groupCommands.Label = commands_group_label;
            }
            catch (Exception exp)
            {
                MessageBox.Show($"Ошибка при попытке копирования Разделов. Ошибка: {exp.Message}");
            }
        }
        /// <summary>
        /// Копировать МСГ работы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCopyMSGWork_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                CopyedObjectsList.Clear();
                var selection = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                if (selection.Column <= MSGExellModel.MSG_NUMBER_COL | selection.Column >= MSGExellModel.MSG_NEEDS_OF_MACHINE_QUANTITY_COL) return;

                var sected_objects = CurrentMSGExellModel.GetObjectsBySelection(selection, typeof(MSGWork)).Where(ob => !CopyedObjectsList.Contains(ob));
                if (sected_objects != null)
                {
                    foreach (var obj in sected_objects)
                        CopyedObjectsList.Add((MSGWork)obj.Clone()); ;

                    buttonPaste.Enabled = true;
                    commands_group_label = $"МСГ скопированы.\n Выбрерите разде куда вставить МСГ";

                }
                groupCommands.Label = commands_group_label;
            }
            catch (Exception exp)
            {
                MessageBox.Show($"Ошибка при попытке копирования МСГ работ. Ошибка: {exp.Message}");
            }
        }
        private void btnCopyVOVRWork_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                CopyedObjectsList.Clear();
                var selection = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                if (selection.Column <= MSGExellModel.VOVR_NUMBER_COL | selection.Column >= MSGExellModel.VOVR_LABOURNESS_COL) return;
                var sected_objects = CurrentMSGExellModel.GetObjectsBySelection(selection, typeof(VOVRWork)).Where(ob => !CopyedObjectsList.Contains(ob));
                if (sected_objects != null)
                {
                    foreach (var obj in sected_objects)
                        CopyedObjectsList.Add((VOVRWork)obj.Clone()); ;

                    buttonPaste.Enabled = true;
                    commands_group_label = $"ВОВР скопированы.\n Выбрерите разде куда вставить МСГ";

                }
                groupCommands.Label = commands_group_label;
            }
            catch (Exception exp)
            {
                MessageBox.Show($"Ошибка при попытке копирования ВОВР работ. Ошибка: {exp.Message}");
            }
        }

        private void btnCopyKSWork_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                CopyedObjectsList.Clear();
                var selection = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                if (selection.Column <= MSGExellModel.KS_NUMBER_COL | selection.Column >= MSGExellModel.KS_LABOURNESS_COL) return;
                var sected_objects = CurrentMSGExellModel.GetObjectsBySelection(selection, typeof(KSWork)).Where(ob => !CopyedObjectsList.Contains(ob));
                if (sected_objects != null)
                {
                    foreach (var obj in sected_objects)
                        CopyedObjectsList.Add((KSWork)obj.Clone()); ;

                    buttonPaste.Enabled = true;
                    commands_group_label = $"RC-2 скопированы.\n Выбрерите разде куда вставить МСГ";

                }
                groupCommands.Label = commands_group_label;
            }
            catch (Exception exp)
            {
                MessageBox.Show($"Ошибка при попытке копирования КС работ. Ошибка: {exp.Message}");
            }
        }
        private void btnCopyRCWork_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                CopyedObjectsList.Clear();
                var selection = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                if (selection.Column <= MSGExellModel.RC_NUMBER_COL | selection.Column >= MSGExellModel.RC_LABOURNESS_COL) return;
                var sected_objects = CurrentMSGExellModel.GetObjectsBySelection(selection, typeof(RCWork)).Where(ob => !CopyedObjectsList.Contains(ob));
                if (sected_objects != null)
                {
                    foreach (var obj in sected_objects)
                        CopyedObjectsList.Add((RCWork)obj.Clone()); ;

                    buttonPaste.Enabled = true;
                    commands_group_label = $"ТУВР скопированы.\n Выбрерите разде куда вставить МСГ";

                }
                groupCommands.Label = commands_group_label;
            }
            catch (Exception exp)
            {
                MessageBox.Show($"Ошибка при попытке копирования ТУВР работ. Ошибка: {exp.Message}");
            }
        }

        private void btnLoadInModelLocal_Click(object sender, RibbonControlEventArgs e)
        {
            //if (CurrentMSGExellModel.IsHasEnabled)
            //    CurrentMSGExellModel.SetHashFormulas();
            if (CurrentMSGExellModel.Owner == null)
                CurrentMSGExellModel.ReloadSheetModelLocal();
        }

        private void chckBoxHashEnable_Click(object sender, RibbonControlEventArgs e)
        {
            CurrentMSGExellModel.IsHasEnabled = chckBoxHashEnable.Checked;
            btnLoadInModelLocal.Enabled = chckBoxHashEnable.Checked;


        }

        private void btnCreateMSGForEmployers_Click(object sender, RibbonControlEventArgs e)
        {
            //  try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog
                {
                    InitialDirectory = @"С:\",
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
                    // CurrentMSGExellModel.ReloadSheetModel();
                    //  CurrentMSGExellModel.CalcAll();
                    MSGExellModel current_model = CurrentMSGExellModel;

                    if (CurrentMSGExellModel.Owner != null)
                        current_model = CurrentMSGExellModel.Owner;

                    foreach (MSGExellModel model in current_model.Children)
                    {
                        MSGTemplateWorkbook = Globals.ThisAddIn.Application.Workbooks.Open(temlate_file_name);
                        MSGTemplateWorkbook.Activate();
                        model.ReloadSheetModel();
                        model.CalcAll();
                        this.FillMSG_OUT_File(model, (w) => w.Quantity != 0);
                        MSGTemplateWorkbook.SaveAs($"{MSGTemplateWorkbook.Path}\\{model.Employer.Name}_{model.ContractCode}.xlsx");
                        MSGTemplateWorkbook.Close();
                    }



                }
            }
            //   catch (Exception exp)
            {

                //     MessageBox.Show($"Ошибка при выводе данных в шаблом графика МСГ. Ошибка:{exp.Message}");
            }
        }
    }

}
