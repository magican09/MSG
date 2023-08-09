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
            this.btnLoadInModel = this.Factory.CreateRibbonButton();
            this.btnLoadFromModel = this.Factory.CreateRibbonButton();
            this.menuCommon = this.Factory.CreateRibbonMenu();
            this.btnUpdateAll = this.Factory.CreateRibbonButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.btnChangeCommonMSG = this.Factory.CreateRibbonButton();
            this.comboBoxEmployerName = this.Factory.CreateRibbonComboBox();
            this.bntChangeEmployerMSG = this.Factory.CreateRibbonButton();
            this.groupMSGCommon = this.Factory.CreateRibbonGroup();
            this.btnCalcAll = this.Factory.CreateRibbonButton();
            this.buttonCalc = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btnCalcLabournes = this.Factory.CreateRibbonButton();
            this.groupCommands = this.Factory.CreateRibbonGroup();
            this.menuEditCommands = this.Factory.CreateRibbonMenu();
            this.buttonCopy = this.Factory.CreateRibbonButton();
            this.menuMSG = this.Factory.CreateRibbonMenu();
            this.btnCopyMSGWork = this.Factory.CreateRibbonButton();
            this.btnInitMSGContent = this.Factory.CreateRibbonButton();
            this.buttonPaste = this.Factory.CreateRibbonButton();
            this.grpInChargePersons = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnChangeEmployers = this.Factory.CreateRibbonButton();
            this.btnChangePosts = this.Factory.CreateRibbonButton();
            this.btnChangeUOM = this.Factory.CreateRibbonButton();
            this.btnSelectPerson = this.Factory.CreateRibbonButton();
            this.btnMachines = this.Factory.CreateRibbonButton();
            this.groupMSG_OUT = this.Factory.CreateRibbonGroup();
            this.btnCreateTemplateFile = this.Factory.CreateRibbonButton();
            this.checkBoxSandayVocationrStatus = this.Factory.CreateRibbonCheckBox();
            this.checkBoxRerightDatePart = this.Factory.CreateRibbonCheckBox();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.btnRefillTemlate = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnShowAlllHidenWorksheets = this.Factory.CreateRibbonButton();
            this.labelConractCode = this.Factory.CreateRibbonLabel();
            this.labelCurrentEmployerName = this.Factory.CreateRibbonLabel();
            this.openMSGTemplateFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.tab1.SuspendLayout();
            this.groupFileLaod.SuspendLayout();
            this.groupMSGCommon.SuspendLayout();
            this.groupCommands.SuspendLayout();
            this.grpInChargePersons.SuspendLayout();
            this.groupMSG_OUT.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.groupFileLaod);
            this.tab1.Groups.Add(this.groupMSGCommon);
            this.tab1.Groups.Add(this.groupCommands);
            this.tab1.Groups.Add(this.grpInChargePersons);
            this.tab1.Groups.Add(this.groupMSG_OUT);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "МСГ";
            this.tab1.Name = "tab1";
            // 
            // groupFileLaod
            // 
            this.groupFileLaod.Items.Add(this.btnLoadMSGFile);
            this.groupFileLaod.Items.Add(this.btnLoadInModel);
            this.groupFileLaod.Items.Add(this.btnLoadFromModel);
            this.groupFileLaod.Items.Add(this.menuCommon);
            this.groupFileLaod.Items.Add(this.separator4);
            this.groupFileLaod.Items.Add(this.btnChangeCommonMSG);
            this.groupFileLaod.Items.Add(this.comboBoxEmployerName);
            this.groupFileLaod.Items.Add(this.bntChangeEmployerMSG);
            this.groupFileLaod.Label = "Загрузка";
            this.groupFileLaod.Name = "groupFileLaod";
            // 
            // btnLoadMSGFile
            // 
            this.btnLoadMSGFile.Label = "Загрузить";
            this.btnLoadMSGFile.Name = "btnLoadMSGFile";
            this.btnLoadMSGFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadMSGFile_Click);
            // 
            // btnLoadInModel
            // 
            this.btnLoadInModel.Enabled = false;
            this.btnLoadInModel.Label = "В МОДЕЛЬ";
            this.btnLoadInModel.Name = "btnLoadInModel";
            this.btnLoadInModel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadInModel_Click);
            // 
            // btnLoadFromModel
            // 
            this.btnLoadFromModel.Enabled = false;
            this.btnLoadFromModel.Label = "ИЗ МОДЕЛИ";
            this.btnLoadFromModel.Name = "btnLoadFromModel";
            this.btnLoadFromModel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadFromModel_Click);
            // 
            // menuCommon
            // 
            this.menuCommon.Items.Add(this.btnUpdateAll);
            this.menuCommon.Label = "Глобальные команды";
            this.menuCommon.Name = "menuCommon";
            // 
            // btnUpdateAll
            // 
            this.btnUpdateAll.Enabled = false;
            this.btnUpdateAll.Label = "ОБНОВИТЬ ВСЁ";
            this.btnUpdateAll.Name = "btnUpdateAll";
            this.btnUpdateAll.ShowImage = true;
            this.btnUpdateAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateAll_Click);
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // btnChangeCommonMSG
            // 
            this.btnChangeCommonMSG.Enabled = false;
            this.btnChangeCommonMSG.Label = "Общая ведомость";
            this.btnChangeCommonMSG.Name = "btnChangeCommonMSG";
            this.btnChangeCommonMSG.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangeCommonMSG_Click);
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
            // groupMSGCommon
            // 
            this.groupMSGCommon.Items.Add(this.btnCalcAll);
            this.groupMSGCommon.Items.Add(this.buttonCalc);
            this.groupMSGCommon.Items.Add(this.separator2);
            this.groupMSGCommon.Items.Add(this.btnCalcLabournes);
            this.groupMSGCommon.Label = "Расчеты";
            this.groupMSGCommon.Name = "groupMSGCommon";
            // 
            // btnCalcAll
            // 
            this.btnCalcAll.Enabled = false;
            this.btnCalcAll.Label = "ПЕРЕСЧИТАТЬ ВСЕ";
            this.btnCalcAll.Name = "btnCalcAll";
            this.btnCalcAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCalcAll_Click);
            // 
            // buttonCalc
            // 
            this.buttonCalc.Enabled = false;
            this.buttonCalc.Label = "ПЕРЕСЧИТАТЬ";
            this.buttonCalc.Name = "buttonCalc";
            this.buttonCalc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCalc_Click);
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
            // groupCommands
            // 
            this.groupCommands.Items.Add(this.menuEditCommands);
            this.groupCommands.Items.Add(this.menuMSG);
            this.groupCommands.Items.Add(this.buttonPaste);
            this.groupCommands.Label = "_______";
            this.groupCommands.Name = "groupCommands";
            // 
            // menuEditCommands
            // 
            this.menuEditCommands.Enabled = false;
            this.menuEditCommands.Items.Add(this.buttonCopy);
            this.menuEditCommands.Label = "РАЗДЕЛ";
            this.menuEditCommands.Name = "menuEditCommands";
            // 
            // buttonCopy
            // 
            this.buttonCopy.Label = "Копировать";
            this.buttonCopy.Name = "buttonCopy";
            this.buttonCopy.ShowImage = true;
            this.buttonCopy.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCopy_Click);
            // 
            // menuMSG
            // 
            this.menuMSG.Items.Add(this.btnCopyMSGWork);
            this.menuMSG.Items.Add(this.btnInitMSGContent);
            this.menuMSG.Label = "МСГ";
            this.menuMSG.Name = "menuMSG";
            // 
            // btnCopyMSGWork
            // 
            this.btnCopyMSGWork.Label = "Копировать";
            this.btnCopyMSGWork.Name = "btnCopyMSGWork";
            this.btnCopyMSGWork.ShowImage = true;
            this.btnCopyMSGWork.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopyMSGWork_Click);
            // 
            // btnInitMSGContent
            // 
            this.btnInitMSGContent.Label = "Дописать ...";
            this.btnInitMSGContent.Name = "btnInitMSGContent";
            this.btnInitMSGContent.ShowImage = true;
            this.btnInitMSGContent.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInitMSGContent_Click);
            // 
            // buttonPaste
            // 
            this.buttonPaste.Enabled = false;
            this.buttonPaste.Label = "Вставить";
            this.buttonPaste.Name = "buttonPaste";
            this.buttonPaste.ShowImage = true;
            this.buttonPaste.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonPaste_Click);
            // 
            // grpInChargePersons
            // 
            this.grpInChargePersons.Items.Add(this.separator1);
            this.grpInChargePersons.Items.Add(this.btnChangeEmployers);
            this.grpInChargePersons.Items.Add(this.btnChangePosts);
            this.grpInChargePersons.Items.Add(this.btnChangeUOM);
            this.grpInChargePersons.Items.Add(this.btnSelectPerson);
            this.grpInChargePersons.Items.Add(this.btnMachines);
            this.grpInChargePersons.Label = "Общие данные";
            this.grpInChargePersons.Name = "grpInChargePersons";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnChangeEmployers
            // 
            this.btnChangeEmployers.Label = "Отвественные";
            this.btnChangeEmployers.Name = "btnChangeEmployers";
            this.btnChangeEmployers.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangeEmployers_Click);
            // 
            // btnChangePosts
            // 
            this.btnChangePosts.Label = "Должности";
            this.btnChangePosts.Name = "btnChangePosts";
            this.btnChangePosts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangePosts_Click);
            // 
            // btnChangeUOM
            // 
            this.btnChangeUOM.Label = "Ед.изм.";
            this.btnChangeUOM.Name = "btnChangeUOM";
            this.btnChangeUOM.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangeUOM_Click);
            // 
            // btnSelectPerson
            // 
            this.btnSelectPerson.Label = "";
            this.btnSelectPerson.Name = "btnSelectPerson";
            // 
            // btnMachines
            // 
            this.btnMachines.Label = "Техника";
            this.btnMachines.Name = "btnMachines";
            this.btnMachines.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMachines_Click);
            // 
            // groupMSG_OUT
            // 
            this.groupMSG_OUT.Items.Add(this.btnCreateTemplateFile);
            this.groupMSG_OUT.Items.Add(this.checkBoxSandayVocationrStatus);
            this.groupMSG_OUT.Items.Add(this.checkBoxRerightDatePart);
            this.groupMSG_OUT.Items.Add(this.separator3);
            this.groupMSG_OUT.Items.Add(this.btnRefillTemlate);
            this.groupMSG_OUT.Label = "МСГ выход";
            this.groupMSG_OUT.Name = "groupMSG_OUT";
            // 
            // btnCreateTemplateFile
            // 
            this.btnCreateTemplateFile.Enabled = false;
            this.btnCreateTemplateFile.Label = "Создать МСГ из шаблона";
            this.btnCreateTemplateFile.Name = "btnCreateTemplateFile";
            this.btnCreateTemplateFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadTeplateFile_Click);
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
            // btnRefillTemlate
            // 
            this.btnRefillTemlate.Enabled = false;
            this.btnRefillTemlate.Label = "Обновить существующий  МСГ";
            this.btnRefillTemlate.Name = "btnRefillTemlate";
            this.btnRefillTemlate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFillTemlate_Click);
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
            this.groupCommands.ResumeLayout(false);
            this.groupCommands.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCalcAll;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadFromModel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRefillTemlate;
        private System.Windows.Forms.OpenFileDialog openMSGTemplateFileDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMSG_OUT;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxSandayVocationrStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateTemplateFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupFileLaod;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadMSGFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelConractCode;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxRerightDatePart;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelCurrentEmployerName;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCopy;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupCommands;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonPaste;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCalc;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuEditCommands;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopyMSGWork;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadInModel;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuCommon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMachines;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuMSG;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInitMSGContent;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
