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



        MSGExellModel CurrentMSGExellModel;
        MSGExellModel CommonMSGExellModel;
        ObservableCollection<MSGExellModel> MSGExellModels = new ObservableCollection<MSGExellModel>();

        ObservableCollection<Employer> Employers { get; set; } = new ObservableCollection<Employer>();
        Excel._Workbook CurrentWorkbook;
        
        Excel.Worksheet EmployersWorksheet;
        Excel.Worksheet PostsWorksheet;
        Excel.Worksheet MeasurementsWorksheet;
       
        private void OnActiveWorksheetChanged(Excel.Worksheet last_wsh, Excel.Worksheet new_wsh)
        {
            if (CurrentWorkbook == null)
                CurrentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
 
            if (EmployersWorksheet == null)
                EmployersWorksheet = CurrentWorkbook.Worksheets["Ответственные"];
            if (PostsWorksheet == null)
                PostsWorksheet = CurrentWorkbook.Worksheets["Должности"];
            if (MeasurementsWorksheet == null)
                MeasurementsWorksheet = CurrentWorkbook.Worksheets["Ед_изм"];

         
            this.ReloadEmployersList();

            switch (new_wsh.Name)
            {
                case "Ведомость_общая":
                    {

                        break;
                    }
            }

        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.OnActiveWorksheetChanged += OnActiveWorksheetChanged;

        }

        private void buttonMSGLoad_Click(object sender, RibbonControlEventArgs e)
        {
            MSGExellModels.Clear();
            CommonMSGExellModel = new MSGExellModel();
            CommonMSGExellModel.RegisterSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets["Ведомость_общая"];

            PostsWorksheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            EmployersWorksheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;

            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                if (worksheet.Name.Contains("_"))
                {
                    string emoloyer_namber_str = worksheet.Name.Substring(worksheet.Name.LastIndexOf("_") + 1, worksheet.Name.Length - worksheet.Name.LastIndexOf("_") - 1);
                    int employer_number;
                     int.TryParse(emoloyer_namber_str,out employer_number);
                    Employer employer = Employers.Where(em => em.Number == employer_number).FirstOrDefault();
                    if(employer!=null)
                    {
                        MSGExellModel model = new MSGExellModel();
                        model.RegisterSheet = worksheet;
                    }
                }

            }

            //CurrentMSGExellModel.MSGWorks.Clear();
            //CurrentMSGExellModel.VOVRWorks.Clear();
            //CurrentMSGExellModel.KSWorks.Clear();

            //CurrentMSGExellModel.LoadMSGWorks();
            //CurrentMSGExellModel.LoadVOVRWorks();
            //CurrentMSGExellModel.LoadKSWorks();
            //CurrentMSGExellModel. LoadWorksReportCards();

        }


        private void btnNotifyTest_Click(object sender, RibbonControlEventArgs e)
        {
            CurrentMSGExellModel.CalcLabourness();
            CurrentMSGExellModel.CalcQuantity();

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

        private void btnChangeEmployers_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeWorksheet = CurrentWorkbook.ActiveSheet;
            activeWorksheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            CurrentWorkbook.Worksheets["Ответственные"].Visible = Excel.XlSheetVisibility.xlSheetVisible;
            CurrentWorkbook.Worksheets["Ответственные"].Activate();
        }

        private void btnChangePosts_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeWorksheet = CurrentWorkbook.ActiveSheet;
            activeWorksheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            CurrentWorkbook.Worksheets["Должности"].Visible = Excel.XlSheetVisibility.xlSheetVisible;
            CurrentWorkbook.Worksheets["Должности"].Activate();
        }
    }
}
