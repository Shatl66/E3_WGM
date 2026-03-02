namespace E3_WGM
{
    partial class E3ProjectBrowser
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
            this.StructureBrowserPanel = new System.Windows.Forms.Panel();
            this.labelNameProject = new System.Windows.Forms.Label();
            this.labelNumberProject = new System.Windows.Forms.Label();
            this.textBoxNameProject = new System.Windows.Forms.TextBox();
            this.textBoxNumberProject = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.StructureBrowserPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // StructureBrowserPanel
            // 
            this.StructureBrowserPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.StructureBrowserPanel.Controls.Add(this.labelNameProject);
            this.StructureBrowserPanel.Controls.Add(this.labelNumberProject);
            this.StructureBrowserPanel.Controls.Add(this.textBoxNameProject);
            this.StructureBrowserPanel.Controls.Add(this.textBoxNumberProject);
            this.StructureBrowserPanel.Controls.Add(this.button1);
            this.StructureBrowserPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.StructureBrowserPanel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.StructureBrowserPanel.Location = new System.Drawing.Point(0, 0);
            this.StructureBrowserPanel.Name = "StructureBrowserPanel";
            this.StructureBrowserPanel.Size = new System.Drawing.Size(618, 174);
            this.StructureBrowserPanel.TabIndex = 3;
            // 
            // labelNameProject
            // 
            this.labelNameProject.AutoSize = true;
            this.labelNameProject.Location = new System.Drawing.Point(9, 109);
            this.labelNameProject.Name = "labelNameProject";
            this.labelNameProject.Size = new System.Drawing.Size(233, 17);
            this.labelNameProject.TabIndex = 4;
            this.labelNameProject.Text = "Наименование выбранной сборки";
            // 
            // labelNumberProject
            // 
            this.labelNumberProject.AutoSize = true;
            this.labelNumberProject.Location = new System.Drawing.Point(9, 80);
            this.labelNumberProject.Name = "labelNumberProject";
            this.labelNumberProject.Size = new System.Drawing.Size(225, 17);
            this.labelNumberProject.TabIndex = 3;
            this.labelNumberProject.Text = "Обозначение выбранной сборки";
            // 
            // textBoxNameProject
            // 
            this.textBoxNameProject.Location = new System.Drawing.Point(252, 109);
            this.textBoxNameProject.Multiline = true;
            this.textBoxNameProject.Name = "textBoxNameProject";
            this.textBoxNameProject.Size = new System.Drawing.Size(349, 50);
            this.textBoxNameProject.TabIndex = 2;
            // 
            // textBoxNumberProject
            // 
            this.textBoxNumberProject.Location = new System.Drawing.Point(252, 80);
            this.textBoxNumberProject.Name = "textBoxNumberProject";
            this.textBoxNumberProject.Size = new System.Drawing.Size(349, 23);
            this.textBoxNumberProject.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Location = new System.Drawing.Point(12, 14);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(158, 40);
            this.button1.TabIndex = 0;
            this.button1.Text = "Выгрузить проект Е3";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.buttonUploadProjectDoc_Click);
            // 
            // E3ProjectBrowser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.StructureBrowserPanel);
            this.Name = "E3ProjectBrowser";
            this.Size = new System.Drawing.Size(618, 174);
            this.StructureBrowserPanel.ResumeLayout(false);
            this.StructureBrowserPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel StructureBrowserPanel;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBoxNameProject;
        private System.Windows.Forms.TextBox textBoxNumberProject;
        private System.Windows.Forms.Label labelNameProject;
        private System.Windows.Forms.Label labelNumberProject;
    }
}
