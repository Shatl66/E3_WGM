namespace E3_WGM
{
    partial class E3DocListBrowser
    {
        /// <summary> 
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary> 
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.buttonUploadDoc = new System.Windows.Forms.Button();
            this.StructureBrowserPanel = new System.Windows.Forms.Panel();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.isRowSelected = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.oidMaster = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.oid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.number = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ATR_BOM_RS = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ATR_DOC_TYPE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ATR_DOC_FORMAT = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.list_Amount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.StructureBrowserPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonUploadDoc
            // 
            this.buttonUploadDoc.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.buttonUploadDoc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonUploadDoc.Location = new System.Drawing.Point(42, 8);
            this.buttonUploadDoc.Name = "buttonUploadDoc";
            this.buttonUploadDoc.Size = new System.Drawing.Size(100, 25);
            this.buttonUploadDoc.TabIndex = 0;
            this.buttonUploadDoc.Text = "Выгрузить";
            this.buttonUploadDoc.UseVisualStyleBackColor = true;
            this.buttonUploadDoc.Click += new System.EventHandler(this.buttonUploadDoc_Click);
            // 
            // StructureBrowserPanel
            // 
            this.StructureBrowserPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.StructureBrowserPanel.Controls.Add(this.buttonUploadDoc);
            this.StructureBrowserPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.StructureBrowserPanel.Location = new System.Drawing.Point(0, 0);
            this.StructureBrowserPanel.Name = "StructureBrowserPanel";
            this.StructureBrowserPanel.Size = new System.Drawing.Size(624, 40);
            this.StructureBrowserPanel.TabIndex = 3;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.isRowSelected,
            this.oidMaster,
            this.oid,
            this.number,
            this.name,
            this.ATR_BOM_RS,
            this.ATR_DOC_TYPE,
            this.ATR_DOC_FORMAT,
            this.list_Amount});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 40);
            this.dataGridView1.MultiSelect = false;
            this.dataGridView1.Name = "dataGridView1";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridView1.Size = new System.Drawing.Size(624, 124);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.CurrentCellDirtyStateChanged += new System.EventHandler(this.dataGridView1_CurrentCellDirtyStateChanged);
            // 
            // isRowSelected
            // 
            this.isRowSelected.HeaderText = "";
            this.isRowSelected.Name = "isRowSelected";
            this.isRowSelected.Width = 50;
            // 
            // oidMaster
            // 
            this.oidMaster.HeaderText = "oidMaster";
            this.oidMaster.Name = "oidMaster";
            this.oidMaster.Visible = false;
            // 
            // oid
            // 
            this.oid.HeaderText = "oid";
            this.oid.Name = "oid";
            this.oid.Visible = false;
            // 
            // number
            // 
            this.number.HeaderText = "Обозначение";
            this.number.Name = "number";
            this.number.ReadOnly = true;
            this.number.Width = 150;
            // 
            // name
            // 
            this.name.HeaderText = "Наименование";
            this.name.Name = "name";
            this.name.ReadOnly = true;
            this.name.Width = 250;
            // 
            // ATR_BOM_RS
            // 
            this.ATR_BOM_RS.HeaderText = "Раздел в СП";
            this.ATR_BOM_RS.Name = "ATR_BOM_RS";
            this.ATR_BOM_RS.ReadOnly = true;
            this.ATR_BOM_RS.Width = 120;
            // 
            // ATR_DOC_TYPE
            // 
            this.ATR_DOC_TYPE.HeaderText = "Вид документа";
            this.ATR_DOC_TYPE.Name = "ATR_DOC_TYPE";
            this.ATR_DOC_TYPE.ReadOnly = true;
            this.ATR_DOC_TYPE.Width = 250;
            // 
            // ATR_DOC_FORMAT
            // 
            this.ATR_DOC_FORMAT.HeaderText = "Формат документа";
            this.ATR_DOC_FORMAT.Name = "ATR_DOC_FORMAT";
            this.ATR_DOC_FORMAT.ReadOnly = true;
            this.ATR_DOC_FORMAT.Width = 80;
            // 
            // list_Amount
            // 
            this.list_Amount.HeaderText = "Количество листов";
            this.list_Amount.Name = "list_Amount";
            this.list_Amount.ReadOnly = true;
            this.list_Amount.Width = 80;
            // 
            // E3DocListBrowser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.StructureBrowserPanel);
            this.Name = "E3DocListBrowser";
            this.Size = new System.Drawing.Size(624, 164);
            this.StructureBrowserPanel.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonUploadDoc;
        private System.Windows.Forms.Panel StructureBrowserPanel;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewCheckBoxColumn isRowSelected;
        private System.Windows.Forms.DataGridViewTextBoxColumn oidMaster;
        private System.Windows.Forms.DataGridViewTextBoxColumn oid;
        private System.Windows.Forms.DataGridViewTextBoxColumn number;
        private System.Windows.Forms.DataGridViewTextBoxColumn name;
        private System.Windows.Forms.DataGridViewTextBoxColumn ATR_BOM_RS;
        private System.Windows.Forms.DataGridViewTextBoxColumn ATR_DOC_TYPE;
        private System.Windows.Forms.DataGridViewTextBoxColumn ATR_DOC_FORMAT;
        private System.Windows.Forms.DataGridViewTextBoxColumn list_Amount;
    }
}
