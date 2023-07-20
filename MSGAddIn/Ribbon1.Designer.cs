namespace MSGAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnChangeCommonMSG = this.Factory.CreateRibbonButton();
            this.btnCalcLabournes = this.Factory.CreateRibbonButton();
            this.btnCalcQuantities = this.Factory.CreateRibbonButton();
            this.btnChangeUOM = this.Factory.CreateRibbonButton();
            this.grpInChargePersons = this.Factory.CreateRibbonGroup();
            this.comboBoxEmployerName = this.Factory.CreateRibbonComboBox();
            this.bntChangeEmployerMSG = this.Factory.CreateRibbonButton();
            this.btnChangeEmployers = this.Factory.CreateRibbonButton();
            this.btnChangePosts = this.Factory.CreateRibbonButton();
            this.btnSelectPerson = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnShowAlllHidenWorksheets = this.Factory.CreateRibbonButton();
            this.btnReloadWorksheets = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.grpInChargePersons.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.grpInChargePersons);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "МСГ";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnChangeCommonMSG);
            this.group1.Items.Add(this.btnCalcLabournes);
            this.group1.Items.Add(this.btnCalcQuantities);
            this.group1.Items.Add(this.btnChangeUOM);
            this.group1.Items.Add(this.btnReloadWorksheets);
            this.group1.Label = "Общие данные";
            this.group1.Name = "group1";
            // 
            // btnChangeCommonMSG
            // 
            this.btnChangeCommonMSG.Label = "Общая ведомость";
            this.btnChangeCommonMSG.Name = "btnChangeCommonMSG";
            this.btnChangeCommonMSG.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangeCommonMSG_Click);
            // 
            // btnCalcLabournes
            // 
            this.btnCalcLabournes.Enabled = false;
            this.btnCalcLabournes.Label = "Подсчет трудоемкостей";
            this.btnCalcLabournes.Name = "btnCalcLabournes";
            this.btnCalcLabournes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCalcLabournes_Click);
            // 
            // btnCalcQuantities
            // 
            this.btnCalcQuantities.Enabled = false;
            this.btnCalcQuantities.Label = "Подсчет  выполненных работ";
            this.btnCalcQuantities.Name = "btnCalcQuantities";
            this.btnCalcQuantities.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCalcQuantities_Click);
            // 
            // btnChangeUOM
            // 
            this.btnChangeUOM.Label = "Радактировать ед.изм.";
            this.btnChangeUOM.Name = "btnChangeUOM";
            this.btnChangeUOM.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangeUOM_Click);
            // 
            // grpInChargePersons
            // 
            this.grpInChargePersons.Items.Add(this.comboBoxEmployerName);
            this.grpInChargePersons.Items.Add(this.bntChangeEmployerMSG);
            this.grpInChargePersons.Items.Add(this.btnChangeEmployers);
            this.grpInChargePersons.Items.Add(this.btnChangePosts);
            this.grpInChargePersons.Items.Add(this.btnSelectPerson);
            this.grpInChargePersons.Label = "Отвественные";
            this.grpInChargePersons.Name = "grpInChargePersons";
            // 
            // comboBoxEmployerName
            // 
            this.comboBoxEmployerName.Label = "Выбор ответственного";
            this.comboBoxEmployerName.Name = "comboBoxEmployerName";
            this.comboBoxEmployerName.ShowLabel = false;
            this.comboBoxEmployerName.Text = null;
            this.comboBoxEmployerName.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBoxEmployerName_TextChanged);
            // 
            // bntChangeEmployerMSG
            // 
            this.bntChangeEmployerMSG.Enabled = false;
            this.bntChangeEmployerMSG.Label = "Открыть ведомость  ответственного";
            this.bntChangeEmployerMSG.Name = "bntChangeEmployerMSG";
            this.bntChangeEmployerMSG.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bntChangeEmployerMSG_Click);
            // 
            // btnChangeEmployers
            // 
            this.btnChangeEmployers.Label = "Редактировать список отвественных";
            this.btnChangeEmployers.Name = "btnChangeEmployers";
            this.btnChangeEmployers.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangeEmployers_Click);
            // 
            // btnChangePosts
            // 
            this.btnChangePosts.Label = "Редактировать список должностей";
            this.btnChangePosts.Name = "btnChangePosts";
            this.btnChangePosts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangePosts_Click);
            // 
            // btnSelectPerson
            // 
            this.btnSelectPerson.Label = "";
            this.btnSelectPerson.Name = "btnSelectPerson";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnShowAlllHidenWorksheets);
            this.group2.Label = "group2";
            this.group2.Name = "group2";
            // 
            // btnShowAlllHidenWorksheets
            // 
            this.btnShowAlllHidenWorksheets.Label = "Показать все скрытые листы";
            this.btnShowAlllHidenWorksheets.Name = "btnShowAlllHidenWorksheets";
            this.btnShowAlllHidenWorksheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShowAlllHidenWorksheets_Click);
            // 
            // btnReloadWorksheets
            // 
            this.btnReloadWorksheets.Label = "Обновить";
            this.btnReloadWorksheets.Name = "btnReloadWorksheets";
            this.btnReloadWorksheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReloadWorksheets_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.grpInChargePersons.ResumeLayout(false);
            this.grpInChargePersons.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCalcLabournes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCalcQuantities;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpInChargePersons;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectPerson;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBoxEmployerName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangeEmployers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangePosts;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShowAlllHidenWorksheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangeUOM;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bntChangeEmployerMSG;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangeCommonMSG;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReloadWorksheets;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
