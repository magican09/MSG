using ExellAddInsLib.MSG;

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using System.Collections;

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
            this.ReloadAllModels();



        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.OnActiveWorksheetChanged += OnActiveWorksheetChanged;

        }

        private void btnChangeCommonMSG_Click(object sender, RibbonControlEventArgs e)
        {
            this.ShowWorksheet(CommonMSGWorksheet);
            CurrentMSGExellModel = CommonMSGExellModel;
            btnCalcLabournes.Enabled = true;
            btnCalcQuantities.Enabled = true;
        }
        private void btnCalcLabournes_Click(object sender, RibbonControlEventArgs e)
        {
            CurrentMSGExellModel.CalcLabourness();
        }
        private void btnCalcQuantities_Click(object sender, RibbonControlEventArgs e)
        {
            CurrentMSGExellModel.CalcLabourness();
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
          //  CurrentMSGExellModel.RealoadAll();
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
            }
            CurrentMSGExellModel = empl_model;
            this.ShowWorksheet(empl_model.RegisterSheet);
            empl_model.UpdateWorksheetCommonPart();
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
            //        model.RealoadAll();
                }


            }

        }

        private void btnReloadWorksheets_Click(object sender, RibbonControlEventArgs e)
        {
            CurrentMSGExellModel.RealoadAll();
        }
    }
}
