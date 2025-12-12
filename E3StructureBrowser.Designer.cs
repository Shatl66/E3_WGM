namespace E3_WGM
{
    partial class E3StructureBrowser
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
            this._treeView = new Aga.Controls.Tree.TreeViewAdv();
            this.treeColumnNumber = new Aga.Controls.Tree.TreeColumn();
            this.treeColumnName = new Aga.Controls.Tree.TreeColumn();
            this.treeColumnEntry = new Aga.Controls.Tree.TreeColumn();
            this.treeColumnBomRS = new Aga.Controls.Tree.TreeColumn();
            this.treeColumnAmount = new Aga.Controls.Tree.TreeColumn();
            this.treeColumnUnit = new Aga.Controls.Tree.TreeColumn();
            this.treeColumnLineNumber = new Aga.Controls.Tree.TreeColumn();
            this.treeColumnRefDes = new Aga.Controls.Tree.TreeColumn();
            this.nodeCheckBox1 = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._icon = new Aga.Controls.Tree.NodeControls.NodeIcon();
            this._number = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._name = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._entry = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._bomRS = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._amount = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._unit = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._lineNumber = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this._refDes = new Aga.Controls.Tree.NodeControls.NodeTextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // _treeView
            // 
            this._treeView.AllowColumnReorder = true;
            this._treeView.AutoRowHeight = true;
            this._treeView.BackColor = System.Drawing.SystemColors.Window;
            this._treeView.Columns.Add(this.treeColumnNumber);
            this._treeView.Columns.Add(this.treeColumnName);
            this._treeView.Columns.Add(this.treeColumnEntry);
            this._treeView.Columns.Add(this.treeColumnBomRS);
            this._treeView.Columns.Add(this.treeColumnAmount);
            this._treeView.Columns.Add(this.treeColumnUnit);
            this._treeView.Columns.Add(this.treeColumnLineNumber);
            this._treeView.Columns.Add(this.treeColumnRefDes);
            this._treeView.DefaultToolTipProvider = null;
            this._treeView.Dock = System.Windows.Forms.DockStyle.Fill;
            this._treeView.DragDropMarkColor = System.Drawing.Color.Black;
            this._treeView.FullRowSelect = true;
            this._treeView.LineColor = System.Drawing.SystemColors.ControlDark;
            this._treeView.LoadOnDemand = true;
            this._treeView.Location = new System.Drawing.Point(0, 40);
            this._treeView.Model = null;
            this._treeView.Name = "_treeView";
            this._treeView.NodeControls.Add(this.nodeCheckBox1);
            this._treeView.NodeControls.Add(this._icon);
            this._treeView.NodeControls.Add(this._number);
            this._treeView.NodeControls.Add(this._name);
            this._treeView.NodeControls.Add(this._entry);
            this._treeView.NodeControls.Add(this._bomRS);
            this._treeView.NodeControls.Add(this._amount);
            this._treeView.NodeControls.Add(this._unit);
            this._treeView.NodeControls.Add(this._lineNumber);
            this._treeView.NodeControls.Add(this._refDes);
            this._treeView.SelectedNode = null;
            this._treeView.ShowNodeToolTips = true;
            this._treeView.Size = new System.Drawing.Size(573, 129);
            this._treeView.TabIndex = 1;
            this._treeView.Text = "treeViewAdv1";
            this._treeView.UseColumns = true;
            // 
            // treeColumnNumber
            // 
            this.treeColumnNumber.Header = "Обозначение";
            this.treeColumnNumber.SortOrder = System.Windows.Forms.SortOrder.None;
            this.treeColumnNumber.TooltipText = "Number";
            this.treeColumnNumber.Width = 250;
            // 
            // treeColumnName
            // 
            this.treeColumnName.Header = "Наименование";
            this.treeColumnName.SortOrder = System.Windows.Forms.SortOrder.None;
            this.treeColumnName.TooltipText = null;
            this.treeColumnName.Width = 100;
            // 
            // treeColumnEntry
            // 
            this.treeColumnEntry.Header = "Имя ??";
            this.treeColumnEntry.SortOrder = System.Windows.Forms.SortOrder.None;
            this.treeColumnEntry.TooltipText = null;
            this.treeColumnEntry.Width = 150;
            // 
            // treeColumnBomRS
            // 
            this.treeColumnBomRS.Header = "Раздел спецификации";
            this.treeColumnBomRS.SortOrder = System.Windows.Forms.SortOrder.None;
            this.treeColumnBomRS.TooltipText = null;
            this.treeColumnBomRS.Width = 100;
            // 
            // treeColumnAmount
            // 
            this.treeColumnAmount.Header = "Количество";
            this.treeColumnAmount.SortOrder = System.Windows.Forms.SortOrder.None;
            this.treeColumnAmount.TooltipText = null;
            // 
            // treeColumnUnit
            // 
            this.treeColumnUnit.Header = "Единица измерения";
            this.treeColumnUnit.SortOrder = System.Windows.Forms.SortOrder.None;
            this.treeColumnUnit.TooltipText = null;
            // 
            // treeColumnLineNumber
            // 
            this.treeColumnLineNumber.Header = "Позиция";
            this.treeColumnLineNumber.SortOrder = System.Windows.Forms.SortOrder.None;
            this.treeColumnLineNumber.TooltipText = null;
            // 
            // treeColumnRefDes
            // 
            this.treeColumnRefDes.Header = "Позиц. обозначение";
            this.treeColumnRefDes.SortOrder = System.Windows.Forms.SortOrder.None;
            this.treeColumnRefDes.TooltipText = null;
            this.treeColumnRefDes.Width = 150;
            // 
            // nodeCheckBox1
            // 
            this.nodeCheckBox1.DataPropertyName = "IsChecked";
            this.nodeCheckBox1.IncrementalSearchEnabled = true;
            this.nodeCheckBox1.LeftMargin = 3;
            this.nodeCheckBox1.ParentColumn = this.treeColumnNumber;
            // 
            // _icon
            // 
            this._icon.DataPropertyName = "Icon";
            this._icon.LeftMargin = 1;
            this._icon.ParentColumn = this.treeColumnNumber;
            this._icon.ScaleMode = Aga.Controls.Tree.ImageScaleMode.Clip;
            // 
            // _number
            // 
            this._number.DataPropertyName = "NUMBER";
            this._number.IncrementalSearchEnabled = true;
            this._number.LeftMargin = 3;
            this._number.ParentColumn = this.treeColumnNumber;
            this._number.Trimming = System.Drawing.StringTrimming.EllipsisCharacter;
            // 
            // _name
            // 
            this._name.DataPropertyName = "NAME";
            this._name.IncrementalSearchEnabled = true;
            this._name.LeftMargin = 3;
            this._name.ParentColumn = this.treeColumnName;
            // 
            // _entry
            // 
            this._entry.DataPropertyName = "ATR_E3_ENTRY";
            this._entry.IncrementalSearchEnabled = true;
            this._entry.LeftMargin = 3;
            this._entry.ParentColumn = this.treeColumnEntry;
            // 
            // _bomRS
            // 
            this._bomRS.DataPropertyName = "ATR_BOM_RS";
            this._bomRS.IncrementalSearchEnabled = true;
            this._bomRS.LeftMargin = 3;
            this._bomRS.ParentColumn = this.treeColumnBomRS;
            // 
            // _amount
            // 
            this._amount.DataPropertyName = "Amount";
            this._amount.IncrementalSearchEnabled = true;
            this._amount.LeftMargin = 3;
            this._amount.ParentColumn = this.treeColumnAmount;
            // 
            // _unit
            // 
            this._unit.DataPropertyName = "Unit";
            this._unit.IncrementalSearchEnabled = true;
            this._unit.LeftMargin = 3;
            this._unit.ParentColumn = this.treeColumnUnit;
            // 
            // _lineNumber
            // 
            this._lineNumber.DataPropertyName = "LineNumber";
            this._lineNumber.IncrementalSearchEnabled = true;
            this._lineNumber.LeftMargin = 3;
            this._lineNumber.ParentColumn = this.treeColumnLineNumber;
            // 
            // _refDes
            // 
            this._refDes.DataPropertyName = "RefDes";
            this._refDes.IncrementalSearchEnabled = true;
            this._refDes.LeftMargin = 3;
            this._refDes.ParentColumn = this.treeColumnRefDes;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(573, 40);
            this.panel1.TabIndex = 2;
            // 
            // button1
            // 
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Location = new System.Drawing.Point(221, 7);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(100, 25);
            this.button1.TabIndex = 2;
            this.button1.Text = "Выгрузить";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // E3StructureBrowser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this._treeView);
            this.Controls.Add(this.panel1);
            this.Name = "E3StructureBrowser";
            this.Size = new System.Drawing.Size(573, 169);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private Aga.Controls.Tree.TreeViewAdv _treeView;
        private Aga.Controls.Tree.TreeColumn treeColumnNumber;
        private Aga.Controls.Tree.TreeColumn treeColumnName;
        private Aga.Controls.Tree.TreeColumn treeColumnEntry;
        private Aga.Controls.Tree.TreeColumn treeColumnBomRS;
        private Aga.Controls.Tree.TreeColumn treeColumnAmount;
        private Aga.Controls.Tree.TreeColumn treeColumnUnit;
        private Aga.Controls.Tree.TreeColumn treeColumnLineNumber;
        private Aga.Controls.Tree.TreeColumn treeColumnRefDes;
        private Aga.Controls.Tree.NodeControls.NodeTextBox nodeCheckBox1;
        private Aga.Controls.Tree.NodeControls.NodeIcon _icon;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button1;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _number;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _name;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _entry;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _bomRS;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _amount;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _unit;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _lineNumber;
        private Aga.Controls.Tree.NodeControls.NodeTextBox _refDes;
    }
}
