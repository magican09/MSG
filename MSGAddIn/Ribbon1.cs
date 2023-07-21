using ExellAddInsLib.MSG;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;
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
        ObservableCollection<UnitOfMeasurement> UnitOfMeasurements = new ObservableCollection<UnitOfMeasurement>();

        Excel._Workbook CurrentWorkbook;
        Excel._Workbook MSGTemplateWorkbook;

        Excel.Worksheet EmployersWorksheet;
        Excel.Worksheet PostsWorksheet;
        Excel.Worksheet UnitMeasurementsWorksheet;
        Excel.Worksheet CommonMSGWorksheet;
        Excel.Worksheet TemplateMSGWorksheet;

        ObservableCollection<Excel.Worksheet> EmployerMSGWorksheets;

        Employer SelectedEmloeyer;

        private void OnActiveWorksheetChanged(Excel.Worksheet last_wsh, Excel.Worksheet new_wsh)
        {
            if (first_start_flag)
            {
                CurrentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                EmployersWorksheet = CurrentWorkbook.Worksheets["Ответственные"];
                PostsWorksheet = CurrentWorkbook.Worksheets["Должности"];
                UnitMeasurementsWorksheet = CurrentWorkbook.Worksheets["Ед_изм"];
                CommonMSGWorksheet = CurrentWorkbook.Worksheets["Ведомость_общая"];
                TemplateMSGWorksheet = CurrentWorkbook.Worksheets["Ведомость_шаблон"];
                EmployerMSGWorksheets = new ObservableCollection<Excel.Worksheet>();
                foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
                {
                    if (worksheet.Name.Contains("_"))
                    {
                        string emoloyer_namber_str = worksheet.Name.Substring(worksheet.Name.LastIndexOf("_") + 1, worksheet.Name.Length - worksheet.Name.LastIndexOf("_") - 1);
                        int employer_number;
                        int.TryParse(emoloyer_namber_str, out employer_number);
                        if (employer_number != 0)
                        {
                            EmployerMSGWorksheets.Add(worksheet);
                        }
                    }
                }
                first_start_flag = false;
                this.SetAllWorksheetsVisibleState(XlSheetVisibility.xlSheetHidden);

            }
            this.ReloadEmployersList();
            this.ReloadMeasurementsList();
            // this.ReloadAllModels();



        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.OnActiveWorksheetChanged += OnActiveWorksheetChanged;

        }

        private void btnChangeCommonMSG_Click(object sender, RibbonControlEventArgs e)
        {
            if (CommonMSGExellModel == null)
                this.ReloadAllModels();
            CurrentMSGExellModel = CommonMSGExellModel;
            this.ShowWorksheet(CommonMSGWorksheet);
            btnCalcLabournes.Enabled = true;
            btnCalcQuantities.Enabled = true;
            btnReloadWorksheets.Enabled = true;
        }
        private void btnCalcLabournes_Click(object sender, RibbonControlEventArgs e)
        {
            CurrentMSGExellModel.CalcLabourness();
        }
        private void btnCalcQuantities_Click(object sender, RibbonControlEventArgs e)
        {
            //    CurrentMSGExellModel.CalcLabourness();
            CurrentMSGExellModel.UpdateWorksheetCommonPart();
            CurrentMSGExellModel.CalcQuantity();
        }

        private void btnShowAlllHidenWorksheets_Click(object sender, RibbonControlEventArgs e)
        {
            this.SetAllWorksheetsVisibleState(XlSheetVisibility.xlSheetVisible);

        }
        private void btnChangeUOM_Click(object sender, RibbonControlEventArgs e)
        {
            this.SetAllWorksheetsVisibleState(XlSheetVisibility.xlSheetHidden);
            UnitMeasurementsWorksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            UnitMeasurementsWorksheet.Activate();
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
                //   Excel.Worksheet new_employer_worksheet = CurrentWorkbook.Worksheets.Add(new_worksheet_name);
                new_employer_worksheet.Name = new_worksheet_name;
                //Excel.Range source = CommonMSGWorksheet.Range[CommonMSGWorksheet.Cells[0, 0], CommonMSGWorksheet.Cells[200, MSGExellModel.WRC_NUMBER_COL]]
                //    .Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                //Excel.Range dest = new_employer_worksheet.Range[new_employer_worksheet.Cells[0, 0]];
                //CommonMSGWorksheet.UsedRange.Copy();
                //new_employer_worksheet.UsedRange.PasteSpecial(
                //    XlPasteType.xlPasteAll,
                //    XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                //    Type.Missing, Type.Missing);

                //Excel.Range source = CommonMSGWorksheet.Range[CommonMSGWorksheet.Cells[1, 1], CommonMSGWorksheet.Cells[20000, MSGExellModel.WRC_NUMBER_COL]];
                //source.Copy(Type.Missing);
                //Excel.Range dest = new_employer_worksheet.Range["A1"];
                //dest.PasteSpecial(
                //    XlPasteType.xlPasteAll,
                //    XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                //    Type.Missing, Type.Missing);
                //     new_employer_worksheet.Visible = XlSheetVisibility.xlSheetVisible;
                //     new_employer_worksheet.Activate();
                Range last_source = CommonMSGWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                Excel.Range source = CommonMSGWorksheet.Range[CommonMSGWorksheet.Cells[1, 1], last_source];
                source.Copy();
                Range last_dest = new_employer_worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                Excel.Range dest = new_employer_worksheet.Range[new_employer_worksheet.Cells[1, 1], last_dest];
                dest.PasteSpecial(XlPasteType.xlPasteAll);

                EmployerMSGWorksheets.Add(new_employer_worksheet);
                this.ReloadAllModels();

                empl_model = MSGExellModels.FirstOrDefault(m => m.Employer.Name == SelectedEmloeyer.Name);
                empl_model.ClearWorksheetDaysPart();
            }
            CurrentMSGExellModel = empl_model;
            this.ShowWorksheet(empl_model.RegisterSheet);

            CurrentMSGExellModel.ResetCalculatesFields();
        }
        private void btnChangeEmployers_Click(object sender, RibbonControlEventArgs e)
        {
            this.ShowWorksheet(EmployersWorksheet);

        }
        private void btnChangePosts_Click(object sender, RibbonControlEventArgs e)
        {
            this.ShowWorksheet(PostsWorksheet);
        }

        private void SetAllWorksheetsVisibleState(Excel.XlSheetVisibility visibility)
        {
            CommonMSGWorksheet.Visible = visibility;
            TemplateMSGWorksheet.Visible = visibility;

            EmployersWorksheet.Visible = visibility;
            PostsWorksheet.Visible = visibility;
            UnitMeasurementsWorksheet.Visible = visibility;

            foreach (Excel.Worksheet worksheet in EmployerMSGWorksheets)
                worksheet.Visible = visibility;
        }
        private void ShowWorksheet(Excel.Worksheet worksheet)
        {
            this.SetAllWorksheetsVisibleState(XlSheetVisibility.xlSheetHidden);
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
                int number = int.Parse(PostsWorksheet.Cells[row_index, POST_NUMBER_COL].Value.ToString());
                string name = PostsWorksheet.Cells[row_index, POST_NAME_COL].Value.ToString();
                PostsList.Add(new Post(number, name));
                row_index++;
            }
            row_index = 2;
            while (EmployersWorksheet.Cells[row_index, EMPLOYER_NUMBER_COL].Value != null)
            {
                int number = int.Parse(EmployersWorksheet.Cells[row_index, EMPLOYER_NUMBER_COL].Value.ToString());
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
            CommonMSGExellModel.UnitOfMeasurements = UnitOfMeasurements;
            CommonMSGExellModel.RealoadAll();
            foreach (Excel.Worksheet worksheet in EmployerMSGWorksheets)
            {
                string emoloyer_namber_str = worksheet.Name.Substring(worksheet.Name.LastIndexOf("_") + 1, worksheet.Name.Length - worksheet.Name.LastIndexOf("_") - 1);
                int employer_number;
                int.TryParse(emoloyer_namber_str, out employer_number);
                Employer employer = Employers.Where(em => em.Number == employer_number).FirstOrDefault();
                if (employer != null)
                {
                    MSGExellModel model = new MSGExellModel();
                    model.RegisterSheet = worksheet;
                    model.UnitOfMeasurements = UnitOfMeasurements;

                    model.Employer = employer;
                    model.Owner = CommonMSGExellModel;
                    CommonMSGExellModel.Children.Add(model);
                    MSGExellModels.Add(model);
                    model.UpdateWorksheetCommonPart();
                    model.RealoadAll();


                }


            }

        }

        private void btnReloadWorksheets_Click(object sender, RibbonControlEventArgs e)
        {
            //  if (CurrentMSGExellModel.Owner != null)
            CurrentMSGExellModel.UpdateWorksheetCommonPart();

            CurrentMSGExellModel.RealoadAll();
        }
        private string Template_path;
        private void btnOpenMSGTemplate_Click(object sender, RibbonControlEventArgs e)
        {
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
                //    Template_path= openFileDialog1.
                MSGTemplateWorkbook = Globals.ThisAddIn.Application.Workbooks.Open(temlate_file_name);
                MSGTemplateWorkbook.Activate();
                if (CommonMSGExellModel != null)
                    btnFillTemlate.Enabled = true;
                else
                    btnFillTemlate.Enabled = false;
            }

        }

        private void btnFillTemlate_Click(object sender, RibbonControlEventArgs e)
        {
            const int TMP_NOW_DATE_ROW = 1;
            const int TMP_NOW_DATE_COL = 1;

            const int TMP_WORK_FIRST_INDEX_ROW = 6;

            const int TMP_WORK_NUMBER_COL = 1;
            const int TMP_WORK_NAME_COL = 2;
            const int TMP_WORK_PROJECT_QUANTITY_COL = 4;
            const int TMP_U_MRASURE_COL = 5;

            const int TMP_WORKDAY_DATE_ROW_COL = 2;
            const int TMP_WORKDAY_DATE_FIRST_COL = 19;

            const int WORKDAY_DATE_ROW = 2;
            const int WORKDAY_DATE_FIRST_COL = 18;
          
            int row_index = TMP_WORK_FIRST_INDEX_ROW;
            const int PLAN_PERIOD_MANTHS_NUMBER = 1;

            Excel.Worksheet MSGOutWorksheet = MSGTemplateWorkbook.Worksheets["МСГ"];
            Excel.Worksheet MSGTemplateWorksheet = MSGTemplateWorkbook.Worksheets["МСГ_Шаблон"];
            DateTime current_daye_date = DateTime.Now;

            MSGOutWorksheet.Cells[TMP_NOW_DATE_ROW, TMP_NOW_DATE_COL] = current_daye_date.ToString("d");

            int date_col_index = 0;
            int in_worksheet_number = 18;

            Excel.Range first_week_col_range = MSGTemplateWorksheet.UsedRange.Columns[TMP_WORKDAY_DATE_FIRST_COL - 1];
            Excel.Range first_date_col_range = MSGTemplateWorksheet.UsedRange.Columns[TMP_WORKDAY_DATE_FIRST_COL];

            //MSGOutWorksheet.Cells[TMP_WORKDAY_DATE_ROW_COL, TMP_WORKDAY_DATE_FIRST_COL - 1] =
            //    $"неделя\r\n {CommonMSGExellModel.WorksStartDate.ToString("dd")} - {this.GetLastNotVocationDate(CommonMSGExellModel.WorksStartDate).ToString("dd")}";
            int first_week_day_col;
            int last_week_day_col;
            for (DateTime date = CommonMSGExellModel.WorksStartDate; date <= CommonMSGExellModel.WorksStartDate.AddMonths(PLAN_PERIOD_MANTHS_NUMBER); date = date.AddDays(1))
            {

                if (date.DayOfWeek == DayOfWeek.Sunday || date == CommonMSGExellModel.WorksStartDate)
                {
                    first_week_col_range.Copy();
                    Excel.Range week_day_dest = MSGOutWorksheet.UsedRange.Columns[WORKDAY_DATE_FIRST_COL + date_col_index];
                    week_day_dest.PasteSpecial(XlPasteType.xlPasteAll);
                    MSGOutWorksheet.Cells[WORKDAY_DATE_ROW, WORKDAY_DATE_FIRST_COL + date_col_index] =
                      $"неделя\r\n {date.ToString("dd")} - {this.GetLastNotVocationDate(date).ToString("dd")}";
                    MSGOutWorksheet.Cells[WORKDAY_DATE_ROW+1, WORKDAY_DATE_FIRST_COL + date_col_index] = in_worksheet_number++;

                    first_week_day_col = date_col_index + 1;
                    last_week_day_col = first_week_day_col + (this.GetLastNotVocationDate(date) - date).Days;

                    Excel.Range project_week_pr_q = MSGOutWorksheet.Range[
                      MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW, WORKDAY_DATE_FIRST_COL + date_col_index],
                      MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW, WORKDAY_DATE_FIRST_COL + date_col_index]];

                    Excel.Range project_week_q = MSGOutWorksheet.Range[
                        MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW+1, WORKDAY_DATE_FIRST_COL + date_col_index],
                        MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW+1, WORKDAY_DATE_FIRST_COL + date_col_index]];
                
                    Excel.Range project_week_pr_q_first_day = MSGOutWorksheet.Range[
                         MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW, WORKDAY_DATE_FIRST_COL+ first_week_day_col],
                         MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW, WORKDAY_DATE_FIRST_COL+ first_week_day_col]];

                    Excel.Range project_week_pr_q_last_day = MSGOutWorksheet.Range[
                        MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW, WORKDAY_DATE_FIRST_COL+ last_week_day_col],
                        MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW, WORKDAY_DATE_FIRST_COL+ last_week_day_col]];
                  
                    Excel.Range project_week_q_first_day = MSGOutWorksheet.Range[
                        MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW+1, WORKDAY_DATE_FIRST_COL + first_week_day_col],
                        MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW+1, WORKDAY_DATE_FIRST_COL + first_week_day_col]];

                    Excel.Range project_week_q_last_day = MSGOutWorksheet.Range[
                        MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW+1, WORKDAY_DATE_FIRST_COL + last_week_day_col],
                        MSGOutWorksheet.Cells[TMP_WORK_FIRST_INDEX_ROW+1, WORKDAY_DATE_FIRST_COL + last_week_day_col]];
                    // project_week_q.Formula = $"=SUM(AQ6:AW6)";
                    project_week_q.Formula = $"=SUM({this.RangeAddress(project_week_q_first_day)}:{this.RangeAddress(project_week_q_last_day)})"; ;
                    project_week_pr_q.Formula = $"=SUM({this.RangeAddress(project_week_pr_q_first_day)}:{this.RangeAddress(project_week_pr_q_last_day)})"; ;


                    date_col_index++;

                    last_week_day_col = 0;
                    last_week_day_col = 0;
                }
                first_date_col_range.Copy();
                Excel.Range dest = MSGOutWorksheet.UsedRange.Columns[WORKDAY_DATE_FIRST_COL + date_col_index];
                dest.PasteSpecial(XlPasteType.xlPasteAll);
            
                MSGOutWorksheet.Cells[WORKDAY_DATE_ROW, WORKDAY_DATE_FIRST_COL + date_col_index] = date;
                MSGOutWorksheet.Cells[WORKDAY_DATE_ROW + 1,WORKDAY_DATE_FIRST_COL + date_col_index] = in_worksheet_number++;
             
                date_col_index++;
            }

            ///Заполнение формы данными из модели...
            foreach (MSGWork msg_work in CommonMSGExellModel.MSGWorks)
            {
                MSGOutWorksheet.Cells[row_index, TMP_WORK_NUMBER_COL] = msg_work.Number;
                MSGOutWorksheet.Cells[row_index, TMP_WORK_NAME_COL] = msg_work.Name;
                MSGOutWorksheet.Cells[row_index, TMP_WORK_PROJECT_QUANTITY_COL] = msg_work.ProjectQuantity;
                MSGOutWorksheet.Cells[row_index, TMP_U_MRASURE_COL] = msg_work.UnitOfMeasurement.Name;
                Excel.Range sourse = MSGOutWorksheet.Range[MSGOutWorksheet.Cells[row_index, TMP_WORK_NUMBER_COL], MSGOutWorksheet.Cells[row_index + 1, 10000]];
                sourse.Copy();
                Excel.Range dest = MSGOutWorksheet.Range[MSGOutWorksheet.Cells[row_index + 2, TMP_WORK_NUMBER_COL], MSGOutWorksheet.Cells[row_index + 3, 10000]];
                dest.PasteSpecial(XlPasteType.xlPasteAll);
                //  DateTime first_date = DateTime.Parse(MSGOutWorksheet.Cells[TMP_WORKDAY_DATE_ROW_COL, TMP_WORKDAY_DATE_FIRST_COL].Value.ToString());
                foreach (WorkScheduleChunk schedule_chunk in msg_work.WorkSchedules)
                {
                    int date_index = 0;
                    while (MSGOutWorksheet.Cells[TMP_WORKDAY_DATE_ROW_COL, TMP_WORKDAY_DATE_FIRST_COL + date_index].Value != null && date_index < 1000)
                    {
                        DateTime date;
                        DateTime.TryParse(MSGOutWorksheet.Cells[TMP_WORKDAY_DATE_ROW_COL, TMP_WORKDAY_DATE_FIRST_COL + date_index].Value.ToString(), out date);

                        if (date >= schedule_chunk.StartTime && date <= schedule_chunk.EndTime)
                        {

                            MSGOutWorksheet.Cells[row_index, TMP_WORKDAY_DATE_FIRST_COL + date_index] = msg_work.ProjectQuantity / msg_work.GetShedulesAllDaysNumber();
                        }
                        date_index++;
                    }
                }
                if (msg_work.ReportCard != null)
                    foreach (WorkDay msg_work_day in msg_work.ReportCard)
                    {
                        int date_index = 0;
                        while (MSGOutWorksheet.Cells[TMP_WORKDAY_DATE_ROW_COL, TMP_WORKDAY_DATE_FIRST_COL + date_index].Value != null && date_index < 1000)
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
                row_index += 2;
            }
            MSGTemplateWorkbook.SaveAs(@"D:\1234.xlsx");
            MSGTemplateWorkbook.Close();
        }

        private DateTime GetLastNotVocationDate(DateTime date)
        {
            DateTime out_date = date;
            while (out_date.AddDays(1).DayOfWeek != DayOfWeek.Sunday)
                out_date = out_date.AddDays(1);
            return out_date;
        }
        public string RangeAddress(Excel.Range rng)
        {
            return rng.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1,
                   Type.Missing,Type.Missing);
        }
    }
}
