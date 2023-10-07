using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MSGAddIn
{
    public class SetGlobalsForm : Form
    {
        private TextBox textBoxContructionObjectCode;
        private Label label1;
        private DateTimePicker dateTimePickerStartDate;
        private Label label2;
        private Label label3;
        private TextBox textBoxContractCode;
        private Label label4;
        private TextBox textBoxConstructionSubObjectCode;
        private Button btnSave;
        private DateTime _recordCardStartDate;

        public DateTime RecordCardStartDate
        {
            get { return _recordCardStartDate; }
            set { _recordCardStartDate = value;
                this.dateTimePickerStartDate.Value = _recordCardStartDate;
            } }
        public string ContractCode { get; set; }
        public string ContructionObjectCode { get; set; }
        public string ConstructionSubObjectCode { get; set; }



        public SetGlobalsForm()
        {
            this.InitializeComponent();
        }
        private void InitializeComponent()
        {
            this.btnSave = new System.Windows.Forms.Button();
            this.textBoxContructionObjectCode = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dateTimePickerStartDate = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxContractCode = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxConstructionSubObjectCode = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(622, 176);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 0;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // textBoxContructionObjectCode
            // 
            this.textBoxContructionObjectCode.Location = new System.Drawing.Point(207, 68);
            this.textBoxContructionObjectCode.Multiline = true;
            this.textBoxContructionObjectCode.Name = "textBoxContructionObjectCode";
            this.textBoxContructionObjectCode.Size = new System.Drawing.Size(490, 48);
            this.textBoxContructionObjectCode.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(69, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(132, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Дата  начала всех работ";
            // 
            // dateTimePickerStartDate
            // 
            this.dateTimePickerStartDate.Location = new System.Drawing.Point(207, 15);
            this.dateTimePickerStartDate.Name = "dateTimePickerStartDate";
            this.dateTimePickerStartDate.Size = new System.Drawing.Size(140, 20);
            this.dateTimePickerStartDate.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(24, 68);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(177, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Наменоваение обекта(договора):";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(162, 49);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(39, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Шифр:";
            // 
            // textBoxContractCode
            // 
            this.textBoxContractCode.Location = new System.Drawing.Point(207, 42);
            this.textBoxContractCode.Name = "textBoxContractCode";
            this.textBoxContractCode.Size = new System.Drawing.Size(275, 20);
            this.textBoxContractCode.TabIndex = 6;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(24, 122);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(149, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Нименоваение подобъекта:";
            // 
            // textBoxConstructionSubObjectCode
            // 
            this.textBoxConstructionSubObjectCode.Location = new System.Drawing.Point(207, 122);
            this.textBoxConstructionSubObjectCode.Multiline = true;
            this.textBoxConstructionSubObjectCode.Name = "textBoxConstructionSubObjectCode";
            this.textBoxConstructionSubObjectCode.Size = new System.Drawing.Size(490, 48);
            this.textBoxConstructionSubObjectCode.TabIndex = 7;
            // 
            // SetGlobalsForm
            // 
            this.ClientSize = new System.Drawing.Size(709, 265);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBoxConstructionSubObjectCode);
            this.Controls.Add(this.textBoxContractCode);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dateTimePickerStartDate);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxContructionObjectCode);
            this.Controls.Add(this.btnSave);
            this.Name = "SetGlobalsForm";
            this.Load += new System.EventHandler(this.SetGlobalsForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            _recordCardStartDate = this.dateTimePickerStartDate.Value;
            ContractCode = this.textBoxContractCode.Text;
            ConstructionSubObjectCode = this.textBoxConstructionSubObjectCode.Text;
            ContructionObjectCode = this.textBoxContructionObjectCode.Text;
      
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void SetGlobalsForm_Load(object sender, EventArgs e)
        {
            this.dateTimePickerStartDate.Value =_recordCardStartDate;
            this.textBoxContractCode.Text = ContractCode;
            this.textBoxConstructionSubObjectCode.Text = ConstructionSubObjectCode;
            this.textBoxContructionObjectCode.Text = ContructionObjectCode;
        }


    }
}
