namespace excelTemplate
{
    partial class mainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(mainForm));
            this.mainPanel = new System.Windows.Forms.Panel();
            this.lbAutoContact = new System.Windows.Forms.Label();
            this.tbSearchID = new System.Windows.Forms.TextBox();
            this.lbSearchID = new System.Windows.Forms.Label();
            this.btSearchID = new System.Windows.Forms.Button();
            this.btBrowseTemplate = new System.Windows.Forms.Button();
            this.btBrowseExcel = new System.Windows.Forms.Button();
            this.lbBrowseTemplate = new System.Windows.Forms.Label();
            this.lbBrowseExcel = new System.Windows.Forms.Label();
            this.tbBrowseTemplate = new System.Windows.Forms.TextBox();
            this.tbBrowseExcel = new System.Windows.Forms.TextBox();
            this.contactPanel = new System.Windows.Forms.Panel();
            this.btAutoContact = new System.Windows.Forms.Button();
            this.mainPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainPanel
            // 
            this.mainPanel.Controls.Add(this.btAutoContact);
            this.mainPanel.Controls.Add(this.lbAutoContact);
            this.mainPanel.Controls.Add(this.tbSearchID);
            this.mainPanel.Controls.Add(this.lbSearchID);
            this.mainPanel.Controls.Add(this.btSearchID);
            this.mainPanel.Controls.Add(this.btBrowseTemplate);
            this.mainPanel.Controls.Add(this.btBrowseExcel);
            this.mainPanel.Controls.Add(this.lbBrowseTemplate);
            this.mainPanel.Controls.Add(this.lbBrowseExcel);
            this.mainPanel.Controls.Add(this.tbBrowseTemplate);
            this.mainPanel.Controls.Add(this.tbBrowseExcel);
            this.mainPanel.Location = new System.Drawing.Point(12, 12);
            this.mainPanel.Name = "mainPanel";
            this.mainPanel.Size = new System.Drawing.Size(760, 154);
            this.mainPanel.TabIndex = 0;
            // 
            // lbAutoContact
            // 
            this.lbAutoContact.AutoSize = true;
            this.lbAutoContact.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.lbAutoContact.Location = new System.Drawing.Point(517, 126);
            this.lbAutoContact.Name = "lbAutoContact";
            this.lbAutoContact.Size = new System.Drawing.Size(98, 16);
            this.lbAutoContact.TabIndex = 9;
            this.lbAutoContact.Text = "สร้างสัญญาทั้งหมด";
            // 
            // tbSearchID
            // 
            this.tbSearchID.Location = new System.Drawing.Point(80, 125);
            this.tbSearchID.Name = "tbSearchID";
            this.tbSearchID.Size = new System.Drawing.Size(167, 20);
            this.tbSearchID.TabIndex = 7;
            // 
            // lbSearchID
            // 
            this.lbSearchID.AutoSize = true;
            this.lbSearchID.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.lbSearchID.Location = new System.Drawing.Point(3, 126);
            this.lbSearchID.Name = "lbSearchID";
            this.lbSearchID.Size = new System.Drawing.Size(71, 16);
            this.lbSearchID.TabIndex = 8;
            this.lbSearchID.Text = "เลขที่ใบสมัคร";
            // 
            // btSearchID
            // 
            this.btSearchID.Location = new System.Drawing.Point(253, 123);
            this.btSearchID.Name = "btSearchID";
            this.btSearchID.Size = new System.Drawing.Size(75, 23);
            this.btSearchID.TabIndex = 6;
            this.btSearchID.Text = "ค้นหา";
            this.btSearchID.UseVisualStyleBackColor = true;
            // 
            // btBrowseTemplate
            // 
            this.btBrowseTemplate.Location = new System.Drawing.Point(652, 77);
            this.btBrowseTemplate.Name = "btBrowseTemplate";
            this.btBrowseTemplate.Size = new System.Drawing.Size(95, 23);
            this.btBrowseTemplate.TabIndex = 5;
            this.btBrowseTemplate.Text = "เลือก";
            this.btBrowseTemplate.UseVisualStyleBackColor = true;
            // 
            // btBrowseExcel
            // 
            this.btBrowseExcel.Location = new System.Drawing.Point(652, 26);
            this.btBrowseExcel.Name = "btBrowseExcel";
            this.btBrowseExcel.Size = new System.Drawing.Size(95, 23);
            this.btBrowseExcel.TabIndex = 4;
            this.btBrowseExcel.Text = "เลือก";
            this.btBrowseExcel.UseVisualStyleBackColor = true;
            this.btBrowseExcel.Click += new System.EventHandler(this.btBrowseExcel_Click);
            // 
            // lbBrowseTemplate
            // 
            this.lbBrowseTemplate.AutoSize = true;
            this.lbBrowseTemplate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.lbBrowseTemplate.Location = new System.Drawing.Point(3, 60);
            this.lbBrowseTemplate.Name = "lbBrowseTemplate";
            this.lbBrowseTemplate.Size = new System.Drawing.Size(84, 16);
            this.lbBrowseTemplate.TabIndex = 3;
            this.lbBrowseTemplate.Text = "เลือก template";
            // 
            // lbBrowseExcel
            // 
            this.lbBrowseExcel.AutoSize = true;
            this.lbBrowseExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.lbBrowseExcel.Location = new System.Drawing.Point(3, 9);
            this.lbBrowseExcel.Name = "lbBrowseExcel";
            this.lbBrowseExcel.Size = new System.Drawing.Size(92, 16);
            this.lbBrowseExcel.TabIndex = 2;
            this.lbBrowseExcel.Text = "เลือกเอกสารสัญญา";
            // 
            // tbBrowseTemplate
            // 
            this.tbBrowseTemplate.BackColor = System.Drawing.SystemColors.Window;
            this.tbBrowseTemplate.Location = new System.Drawing.Point(69, 79);
            this.tbBrowseTemplate.Name = "tbBrowseTemplate";
            this.tbBrowseTemplate.ReadOnly = true;
            this.tbBrowseTemplate.Size = new System.Drawing.Size(564, 20);
            this.tbBrowseTemplate.TabIndex = 1;
            // 
            // tbBrowseExcel
            // 
            this.tbBrowseExcel.BackColor = System.Drawing.SystemColors.Window;
            this.tbBrowseExcel.Location = new System.Drawing.Point(69, 28);
            this.tbBrowseExcel.Name = "tbBrowseExcel";
            this.tbBrowseExcel.ReadOnly = true;
            this.tbBrowseExcel.Size = new System.Drawing.Size(564, 20);
            this.tbBrowseExcel.TabIndex = 0;
            // 
            // contactPanel
            // 
            this.contactPanel.Location = new System.Drawing.Point(12, 172);
            this.contactPanel.Name = "contactPanel";
            this.contactPanel.Size = new System.Drawing.Size(760, 377);
            this.contactPanel.TabIndex = 1;
            // 
            // btAutoContact
            // 
            this.btAutoContact.Location = new System.Drawing.Point(621, 123);
            this.btAutoContact.Name = "btAutoContact";
            this.btAutoContact.Size = new System.Drawing.Size(126, 23);
            this.btAutoContact.TabIndex = 10;
            this.btAutoContact.Text = "สร้างสัญญาอัตโนมัติ";
            this.btAutoContact.UseVisualStyleBackColor = true;
            // 
            // mainForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(784, 561);
            this.Controls.Add(this.contactPanel);
            this.Controls.Add(this.mainPanel);
            this.Enabled = false;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "mainForm";
            this.Text = "โปรแกรมสร้างสัญญา";
            this.mainPanel.ResumeLayout(false);
            this.mainPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel mainPanel;
        private System.Windows.Forms.Label lbAutoContact;
        private System.Windows.Forms.TextBox tbSearchID;
        private System.Windows.Forms.Label lbSearchID;
        private System.Windows.Forms.Button btSearchID;
        private System.Windows.Forms.Button btBrowseTemplate;
        private System.Windows.Forms.Button btBrowseExcel;
        private System.Windows.Forms.Label lbBrowseTemplate;
        private System.Windows.Forms.Label lbBrowseExcel;
        private System.Windows.Forms.TextBox tbBrowseTemplate;
        private System.Windows.Forms.TextBox tbBrowseExcel;
        private System.Windows.Forms.Panel contactPanel;
        private System.Windows.Forms.Button btAutoContact;
    }
}

