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
            this.groupFileLaod = this.Factory.CreateRibbonGroup();
            this.btnLoadMSGFile = this.Factory.CreateRibbonButton();
            this.btnCloseMSGFile = this.Factory.CreateRibbonButton();
            this.groupMSGCommon = this.Factory.CreateRibbonGroup();
            this.btnChangeCommonMSG = this.Factory.CreateRibbonButton();
            this.btnCalcQuantities = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btnCalcLabournes = this.Factory.CreateRibbonButton();
            this.btnLabournessCoefficients = this.Factory.CreateRibbonButton();
            this.btnReloadWorksheets = this.Factory.CreateRibbonButton();
            this.grpInChargePersons = this.Factory.CreateRibbonGroup();
            this.comboBoxEmployerName = this.Factory.CreateRibbonComboBox();
            this.bntChangeEmployerMSG = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnChangeEmployers = this.Factory.CreateRibbonButton();
            this.btnChangePosts = this.Factory.CreateRibbonButton();
            this.btnChangeUOM = this.Factory.CreateRibbonButton();
            this.btnSelectPerson = this.Factory.CreateRibbonButton();
            this.groupMSG_OUT = this.Factory.CreateRibbonGroup();
            this.btnLoadTeplateFile = this.Factory.CreateRibbonButton();
            this.checkBoxSandayVocationrStatus = this.Factory.CreateRibbonCheckBox();
            this.checkBoxRerightDatePart = this.Factory.CreateRibbonCheckBox();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.btnFillTemlate = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnShowAlllHidenWorksheets = this.Factory.CreateRibbonButton();
            this.labelConractCode = this.Factory.CreateRibbonLabel();
            this.labelCurrentEmployerName = this.Factory.CreateRibbonLabel();
            this.openMSGTemplateFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.menuCreateCommand = this.Factory.CreateRibbonMenu();
            this.buttonCreateBasedOnSection = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupFileLaod.SuspendLayout();
            this.groupMSGCommon.SuspendLayout();
            this.grpInChargePersons.SuspendLayout();
            this.groupMSG_OUT.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.groupFileLaod);
            this.tab1.Groups.Add(this.groupMSGCommon);
            this.tab1.Groups.Add(this.grpInChargePersons);
            this.tab1.Groups.Add(this.groupMSG_OUT);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "МСГ";
            this.tab1.Name = "tab1";
            // 
            // groupFileLaod
            // 
            this.groupFileLaod.Items.Add(this.btnLoadMSGFile);
            this.groupFileLaod.Items.Add(this.btnCloseMSGFile);
            this.groupFileLaod.Label = "Загрузка";
            this.groupFileLaod.Name = "groupFileLaod";
            // 
            // btnLoadMSGFile
            // 
            this.btnLoadMSGFile.Label = "Загрузить ведомость";
            this.btnLoadMSGFile.Name = "btnLoadMSGFile";
            this.btnLoadMSGFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadMSGFile_Click);
            // 
            // btnCloseMSGFile
            // 
            this.btnCloseMSGFile.Label = "Закрыть";
            this.btnCloseMSGFile.Name = "btnCloseMSGFile";
            // 
            // groupMSGCommon
            // 
            this.groupMSGCommon.Items.Add(this.btnChangeCommonMSG);
            this.groupMSGCommon.Items.Add(this.btnCalcQuantities);
            this.groupMSGCommon.Items.Add(this.separator2);
            this.groupMSGCommon.Items.Add(this.btnCalcLabournes);
            this.groupMSGCommon.Items.Add(this.btnLabournessCoefficients);
            this.groupMSGCommon.Items.Add(this.btnReloadWorksheets);
            this.groupMSGCommon.Items.Add(this.menuCreateCommand);
            this.groupMSGCommon.Label = "Общие данные";
            this.groupMSGCommon.Name = "groupMSGCommon";
            // 
            // btnChangeCommonMSG
            // 
            this.btnChangeCommonMSG.Enabled = false;
            this.btnChangeCommonMSG.Label = "Общая ведомость";
            this.btnChangeCommonMSG.Name = "btnChangeCommonMSG";
            this.btnChangeCommonMSG.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangeCommonMSG_Click);
            // 
            // btnCalcQuantities
            // 
            this.btnCalcQuantities.Enabled = false;
            this.btnCalcQuantities.Label = "ПЕРЕСЧЕТ";
            this.btnCalcQuantities.Name = "btnCalcQuantities";
            this.btnCalcQuantities.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCalcQuantities_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // btnCalcLabournes
            // 
            this.btnCalcLabournes.Enabled = false;
            this.btnCalcLabournes.Label = "Подсчет трудоемкостей";
            this.btnCalcLabournes.Name = "btnCalcLabournes";
            this.btnCalcLabournes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCalcLabournes_Click);
            // 
            // btnLabournessCoefficients
            // 
            this.btnLabournessCoefficients.Enabled = false;
            this.btnLabournessCoefficients.Label = "Пересчет нулевых коэф.";
            this.btnLabournessCoefficients.Name = "btnLabournessCoefficients";
            this.btnLabournessCoefficients.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLabournessCoefficients_Click);
            // 
            // btnReloadWorksheets
            // 
            this.btnReloadWorksheets.Enabled = false;
            this.btnReloadWorksheets.Label = "Обновить";
            this.btnReloadWorksheets.Name = "btnReloadWorksheets";
            this.btnReloadWorksheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReloadWorksheets_Click);
            // 
            // grpInChargePersons
            // 
            this.grpInChargePersons.Items.Add(this.comboBoxEmployerName);
            this.grpInChargePersons.Items.Add(this.bntChangeEmployerMSG);
            this.grpInChargePersons.Items.Add(this.separator1);
            this.grpInChargePersons.Items.Add(this.btnChangeEmployers);
            this.grpInChargePersons.Items.Add(this.btnChangePosts);
            this.grpInChargePersons.Items.Add(this.btnChangeUOM);
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
            // separator1
            // 
            this.separator1.Name = "separator1";
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
            // btnChangeUOM
            // 
            this.btnChangeUOM.Label = "Радактировать ед.изм.";
            this.btnChangeUOM.Name = "btnChangeUOM";
            this.btnChangeUOM.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangeUOM_Click);
            // 
            // btnSelectPerson
            // 
            this.btnSelectPerson.Label = "";
            this.btnSelectPerson.Name = "btnSelectPerson";
            // 
            // groupMSG_OUT
            // 
            this.groupMSG_OUT.Items.Add(this.btnLoadTeplateFile);
            this.groupMSG_OUT.Items.Add(this.checkBoxSandayVocationrStatus);
            this.groupMSG_OUT.Items.Add(this.checkBoxRerightDatePart);
            this.groupMSG_OUT.Items.Add(this.separator3);
            this.groupMSG_OUT.Items.Add(this.btnFillTemlate);
            this.groupMSG_OUT.Label = "МСГ выход";
            this.groupMSG_OUT.Name = "groupMSG_OUT";
            // 
            // btnLoadTeplateFile
            // 
            this.btnLoadTeplateFile.Enabled = false;
            this.btnLoadTeplateFile.Label = "Создать МСГ из шаблона";
            this.btnLoadTeplateFile.Name = "btnLoadTeplateFile";
            this.btnLoadTeplateFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadTeplateFile_Click);
            // 
            // checkBoxSandayVocationrStatus
            // 
            this.checkBoxSandayVocationrStatus.Checked = true;
            this.checkBoxSandayVocationrStatus.Enabled = false;
            this.checkBoxSandayVocationrStatus.Label = "Вых. восскресенье";
            this.checkBoxSandayVocationrStatus.Name = "checkBoxSandayVocationrStatus";
            // 
            // checkBoxRerightDatePart
            // 
            this.checkBoxRerightDatePart.Checked = true;
            this.checkBoxRerightDatePart.Label = "Переписать каленраную часть";
            this.checkBoxRerightDatePart.Name = "checkBoxRerightDatePart";
            this.checkBoxRerightDatePart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBoxRerightDatePart_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // btnFillTemlate
            // 
            this.btnFillTemlate.Enabled = false;
            this.btnFillTemlate.Label = "Обновить существующий  МСГ";
            this.btnFillTemlate.Name = "btnFillTemlate";
            this.btnFillTemlate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFillTemlate_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnShowAlllHidenWorksheets);
            this.group2.Items.Add(this.labelConractCode);
            this.group2.Items.Add(this.labelCurrentEmployerName);
            this.group2.Label = "Вспомогательные";
            this.group2.Name = "group2";
            // 
            // btnShowAlllHidenWorksheets
            // 
            this.btnShowAlllHidenWorksheets.Label = "Показать все скрытые листы";
            this.btnShowAlllHidenWorksheets.Name = "btnShowAlllHidenWorksheets";
            this.btnShowAlllHidenWorksheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShowAlllHidenWorksheets_Click);
            // 
            // labelConractCode
            // 
            this.labelConractCode.Label = "________";
            this.labelConractCode.Name = "labelConractCode";
            // 
            // labelCurrentEmployerName
            // 
            this.labelCurrentEmployerName.Label = "________";
            this.labelCurrentEmployerName.Name = "labelCurrentEmployerName";
            // 
            // openMSGTemplateFileDialog
            // 
            this.openMSGTemplateFileDialog.FileName = "Шаблон МСГ";
            // 
            // menuCreateCommand
            // 
            this.menuCreateCommand.Items.Add(this.buttonCreateBasedOnSection);
            this.menuCreateCommand.Label = "Создать";
            this.menuCreateCommand.Name = "menuCreateCommand";
            // 
            // buttonCreateBasedOnSection
            // 
            this.buttonCreateBasedOnSection.Label = "Рзадел на основании";
            this.buttonCreateBasedOnSection.Name = "buttonCreateBasedOnSection";
            this.buttonCreateBasedOnSection.ShowImage = true;
            this.buttonCreateBasedOnSection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCreateBasedOnSection_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupFileLaod.ResumeLayout(false);
            this.groupFileLaod.PerformLayout();
            this.groupMSGCommon.ResumeLayout(false);
            this.groupMSGCommon.PerformLayout();
            this.grpInChargePersons.ResumeLayout(false);
            this.grpInChargePersons.PerformLayout();
            this.groupMSG_OUT.ResumeLayout(false);
            this.groupMSG_OUT.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMSGCommon;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFillTemlate;
        private System.Windows.Forms.OpenFileDialog openMSGTemplateFileDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMSG_OUT;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxSandayVocationrStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadTeplateFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupFileLaod;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadMSGFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelConractCode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCloseMSGFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxRerightDatePart;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelCurrentEmployerName;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLabournessCoefficients;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuCreateCommand;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCreateBasedOnSection;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
