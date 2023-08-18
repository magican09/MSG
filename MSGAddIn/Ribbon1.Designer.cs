﻿namespace MSGAddIn
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
            this.buttonCalc = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btnCalcLabournes = this.Factory.CreateRibbonButton();
            this.groupCommands = this.Factory.CreateRibbonGroup();
            this.menuSection = this.Factory.CreateRibbonMenu();
            this.buttonCopyWorkSection = this.Factory.CreateRibbonButton();
            this.menuMSG = this.Factory.CreateRibbonMenu();
            this.btnCopyMSGWork = this.Factory.CreateRibbonButton();
            this.btnInitMSGContent = this.Factory.CreateRibbonButton();
            this.btnCopyWorkerComposition = this.Factory.CreateRibbonButton();
            this.btnCopyMachineComposition = this.Factory.CreateRibbonButton();
            this.menuVOVR = this.Factory.CreateRibbonMenu();
            this.btnCopyVOVRWork = this.Factory.CreateRibbonButton();
            this.menuKS = this.Factory.CreateRibbonMenu();
            this.btnCopyKSWork = this.Factory.CreateRibbonButton();
            this.menuRC = this.Factory.CreateRibbonMenu();
            this.btnCopyRCWork = this.Factory.CreateRibbonButton();
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
            this.checkBoxRerightDatePart = this.Factory.CreateRibbonCheckBox();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.groupInfo = this.Factory.CreateRibbonGroup();
            this.btnShowAlllHidenWorksheets = this.Factory.CreateRibbonButton();
            this.labelConractCode = this.Factory.CreateRibbonLabel();
            this.labelCurrentEmployerName = this.Factory.CreateRibbonLabel();
            this.openMSGTemplateFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnLoadInModelLocal = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupFileLaod.SuspendLayout();
            this.groupMSGCommon.SuspendLayout();
            this.groupCommands.SuspendLayout();
            this.grpInChargePersons.SuspendLayout();
            this.groupMSG_OUT.SuspendLayout();
            this.groupInfo.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.groupFileLaod);
            this.tab1.Groups.Add(this.groupMSGCommon);
            this.tab1.Groups.Add(this.groupCommands);
            this.tab1.Groups.Add(this.grpInChargePersons);
            this.tab1.Groups.Add(this.groupMSG_OUT);
            this.tab1.Groups.Add(this.groupInfo);
            this.tab1.Label = "МСГ";
            this.tab1.Name = "tab1";
            // 
            // groupFileLaod
            // 
            this.groupFileLaod.Items.Add(this.btnLoadMSGFile);
            this.groupFileLaod.Items.Add(this.btnLoadInModel);
            this.groupFileLaod.Items.Add(this.btnLoadFromModel);
            this.groupFileLaod.Items.Add(this.menuCommon);
            this.groupFileLaod.Items.Add(this.btnLoadInModelLocal);
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
            this.groupMSGCommon.Items.Add(this.buttonCalc);
            this.groupMSGCommon.Items.Add(this.separator2);
            this.groupMSGCommon.Items.Add(this.btnCalcLabournes);
            this.groupMSGCommon.Label = "Расчеты";
            this.groupMSGCommon.Name = "groupMSGCommon";
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
            this.groupCommands.Items.Add(this.menuSection);
            this.groupCommands.Items.Add(this.menuMSG);
            this.groupCommands.Items.Add(this.menuVOVR);
            this.groupCommands.Items.Add(this.menuKS);
            this.groupCommands.Items.Add(this.menuRC);
            this.groupCommands.Items.Add(this.buttonPaste);
            this.groupCommands.Label = "Команды";
            this.groupCommands.Name = "groupCommands";
            // 
            // menuSection
            // 
            this.menuSection.Enabled = false;
            this.menuSection.Items.Add(this.buttonCopyWorkSection);
            this.menuSection.Label = "РАЗДЕЛ";
            this.menuSection.Name = "menuSection";
            // 
            // buttonCopyWorkSection
            // 
            this.buttonCopyWorkSection.Label = "Копировать";
            this.buttonCopyWorkSection.Name = "buttonCopyWorkSection";
            this.buttonCopyWorkSection.ShowImage = true;
            this.buttonCopyWorkSection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCopyWorkSection_Click);
            // 
            // menuMSG
            // 
            this.menuMSG.Enabled = false;
            this.menuMSG.Items.Add(this.btnCopyMSGWork);
            this.menuMSG.Items.Add(this.btnInitMSGContent);
            this.menuMSG.Items.Add(this.btnCopyWorkerComposition);
            this.menuMSG.Items.Add(this.btnCopyMachineComposition);
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
            // btnCopyWorkerComposition
            // 
            this.btnCopyWorkerComposition.Label = "Копировать работников";
            this.btnCopyWorkerComposition.Name = "btnCopyWorkerComposition";
            this.btnCopyWorkerComposition.ShowImage = true;
            this.btnCopyWorkerComposition.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopyWorkerComposition_Click);
            // 
            // btnCopyMachineComposition
            // 
            this.btnCopyMachineComposition.Label = "Копировать технику";
            this.btnCopyMachineComposition.Name = "btnCopyMachineComposition";
            this.btnCopyMachineComposition.ShowImage = true;
            this.btnCopyMachineComposition.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopyMachineComposition_Click);
            // 
            // menuVOVR
            // 
            this.menuVOVR.Enabled = false;
            this.menuVOVR.Items.Add(this.btnCopyVOVRWork);
            this.menuVOVR.Label = " ВОВР";
            this.menuVOVR.Name = "menuVOVR";
            // 
            // btnCopyVOVRWork
            // 
            this.btnCopyVOVRWork.Label = "Копировать";
            this.btnCopyVOVRWork.Name = "btnCopyVOVRWork";
            this.btnCopyVOVRWork.ShowImage = true;
            this.btnCopyVOVRWork.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopyVOVRWork_Click);
            // 
            // menuKS
            // 
            this.menuKS.Enabled = false;
            this.menuKS.Items.Add(this.btnCopyKSWork);
            this.menuKS.Label = "КС-2";
            this.menuKS.Name = "menuKS";
            // 
            // btnCopyKSWork
            // 
            this.btnCopyKSWork.Label = "Копировать";
            this.btnCopyKSWork.Name = "btnCopyKSWork";
            this.btnCopyKSWork.ShowImage = true;
            this.btnCopyKSWork.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopyKSWork_Click);
            // 
            // menuRC
            // 
            this.menuRC.Enabled = false;
            this.menuRC.Items.Add(this.btnCopyRCWork);
            this.menuRC.Label = "ТУВР";
            this.menuRC.Name = "menuRC";
            // 
            // btnCopyRCWork
            // 
            this.btnCopyRCWork.Label = "Копировать";
            this.btnCopyRCWork.Name = "btnCopyRCWork";
            this.btnCopyRCWork.ShowImage = true;
            this.btnCopyRCWork.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopyRCWork_Click);
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
            this.btnChangeEmployers.Enabled = false;
            this.btnChangeEmployers.Label = "Отвественные";
            this.btnChangeEmployers.Name = "btnChangeEmployers";
            this.btnChangeEmployers.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangeEmployers_Click);
            // 
            // btnChangePosts
            // 
            this.btnChangePosts.Enabled = false;
            this.btnChangePosts.Label = "Должности";
            this.btnChangePosts.Name = "btnChangePosts";
            this.btnChangePosts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangePosts_Click);
            // 
            // btnChangeUOM
            // 
            this.btnChangeUOM.Enabled = false;
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
            this.btnMachines.Enabled = false;
            this.btnMachines.Label = "Техника";
            this.btnMachines.Name = "btnMachines";
            this.btnMachines.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMachines_Click);
            // 
            // groupMSG_OUT
            // 
            this.groupMSG_OUT.Items.Add(this.btnCreateTemplateFile);
            this.groupMSG_OUT.Items.Add(this.checkBoxRerightDatePart);
            this.groupMSG_OUT.Items.Add(this.separator3);
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
            // groupInfo
            // 
            this.groupInfo.Items.Add(this.btnShowAlllHidenWorksheets);
            this.groupInfo.Items.Add(this.labelConractCode);
            this.groupInfo.Items.Add(this.labelCurrentEmployerName);
            this.groupInfo.Label = "Вспомогательные";
            this.groupInfo.Name = "groupInfo";
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
            // btnLoadInModelLocal
            // 
            this.btnLoadInModelLocal.Label = "В МОДЕЛЬ (част.)";
            this.btnLoadInModelLocal.Name = "btnLoadInModelLocal";
            this.btnLoadInModelLocal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadInModelLocal_Click);
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
            this.groupInfo.ResumeLayout(false);
            this.groupInfo.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMSGCommon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCalcLabournes;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpInChargePersons;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectPerson;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBoxEmployerName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangeEmployers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangePosts;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShowAlllHidenWorksheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangeUOM;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bntChangeEmployerMSG;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangeCommonMSG;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadFromModel;
        private System.Windows.Forms.OpenFileDialog openMSGTemplateFileDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMSG_OUT;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateTemplateFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupFileLaod;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadMSGFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelConractCode;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxRerightDatePart;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelCurrentEmployerName;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCopyWorkSection;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupCommands;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonPaste;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCalc;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuSection;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopyMSGWork;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadInModel;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuCommon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMachines;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuMSG;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInitMSGContent;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopyWorkerComposition;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopyMachineComposition;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuVOVR;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopyVOVRWork;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuRC;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopyRCWork;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuKS;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopyKSWork;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadInModelLocal;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
