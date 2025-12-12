namespace E3_WGM
{
    partial class E3WGMForm
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
            this.E3Log = new System.Windows.Forms.TextBox();
            this.e3CommonControl1 = new E3_WGM.E3CommonControl();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageProject = new System.Windows.Forms.TabPage();
            this.e3ProjectBrowser1 = new E3_WGM.E3ProjectBrowser();
            this.tabPageStructureBrowser = new System.Windows.Forms.TabPage();
            this.e3StructureBrowser1 = new E3_WGM.E3StructureBrowser();
            this.tabPageDocListBrowser = new System.Windows.Forms.TabPage();
            this.e3DocListBrowser1 = new E3_WGM.E3DocListBrowser();
            this.tabControl1.SuspendLayout();
            this.tabPageProject.SuspendLayout();
            this.tabPageStructureBrowser.SuspendLayout();
            this.tabPageDocListBrowser.SuspendLayout();
            this.SuspendLayout();
            // 
            // E3Log
            // 
            this.E3Log.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.E3Log.Location = new System.Drawing.Point(0, 480);
            this.E3Log.Multiline = true;
            this.E3Log.Name = "E3Log";
            this.E3Log.Size = new System.Drawing.Size(901, 98);
            this.E3Log.TabIndex = 1;
            // 
            // e3CommonControl1
            // 
            this.e3CommonControl1.Dock = System.Windows.Forms.DockStyle.Top;
            this.e3CommonControl1.Location = new System.Drawing.Point(0, 0);
            this.e3CommonControl1.Name = "e3CommonControl1";
            this.e3CommonControl1.Size = new System.Drawing.Size(901, 73);
            this.e3CommonControl1.TabIndex = 3;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPageProject);
            this.tabControl1.Controls.Add(this.tabPageStructureBrowser);
            this.tabControl1.Controls.Add(this.tabPageDocListBrowser);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 73);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(901, 407);
            this.tabControl1.TabIndex = 4;
            // 
            // tabPageProject
            // 
            this.tabPageProject.Controls.Add(this.e3ProjectBrowser1);
            this.tabPageProject.Location = new System.Drawing.Point(4, 22);
            this.tabPageProject.Name = "tabPageProject";
            this.tabPageProject.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageProject.Size = new System.Drawing.Size(893, 381);
            this.tabPageProject.TabIndex = 2;
            this.tabPageProject.Text = "Проект";
            this.tabPageProject.UseVisualStyleBackColor = true;
            // 
            // e3ProjectBrowser1
            // 
            this.e3ProjectBrowser1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.e3ProjectBrowser1.Location = new System.Drawing.Point(3, 3);
            this.e3ProjectBrowser1.Name = "e3ProjectBrowser1";
            this.e3ProjectBrowser1.Size = new System.Drawing.Size(887, 375);
            this.e3ProjectBrowser1.TabIndex = 0;
            // 
            // tabPageStructureBrowser
            // 
            this.tabPageStructureBrowser.Controls.Add(this.e3StructureBrowser1);
            this.tabPageStructureBrowser.Location = new System.Drawing.Point(4, 22);
            this.tabPageStructureBrowser.Name = "tabPageStructureBrowser";
            this.tabPageStructureBrowser.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageStructureBrowser.Size = new System.Drawing.Size(893, 381);
            this.tabPageStructureBrowser.TabIndex = 0;
            this.tabPageStructureBrowser.Text = "Состав";
            this.tabPageStructureBrowser.UseVisualStyleBackColor = true;
            // 
            // e3StructureBrowser1
            // 
            this.e3StructureBrowser1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.e3StructureBrowser1.Location = new System.Drawing.Point(3, 3);
            this.e3StructureBrowser1.Name = "e3StructureBrowser1";
            this.e3StructureBrowser1.Size = new System.Drawing.Size(887, 375);
            this.e3StructureBrowser1.TabIndex = 0;
            // 
            // tabPageDocListBrowser
            // 
            this.tabPageDocListBrowser.Controls.Add(this.e3DocListBrowser1);
            this.tabPageDocListBrowser.Location = new System.Drawing.Point(4, 22);
            this.tabPageDocListBrowser.Name = "tabPageDocListBrowser";
            this.tabPageDocListBrowser.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageDocListBrowser.Size = new System.Drawing.Size(893, 381);
            this.tabPageDocListBrowser.TabIndex = 1;
            this.tabPageDocListBrowser.Text = "Документация";
            this.tabPageDocListBrowser.UseVisualStyleBackColor = true;
            // 
            // e3DocListBrowser1
            // 
            this.e3DocListBrowser1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.e3DocListBrowser1.Location = new System.Drawing.Point(3, 3);
            this.e3DocListBrowser1.Name = "e3DocListBrowser1";
            this.e3DocListBrowser1.Size = new System.Drawing.Size(887, 375);
            this.e3DocListBrowser1.TabIndex = 0;
            // 
            // E3WGMForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(901, 578);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.e3CommonControl1);
            this.Controls.Add(this.E3Log);
            this.Name = "E3WGMForm";
            this.Text = "E3 WGM";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.E3WGMForm_FormClosed);
            this.Load += new System.EventHandler(this.E3WGMForm_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPageProject.ResumeLayout(false);
            this.tabPageStructureBrowser.ResumeLayout(false);
            this.tabPageDocListBrowser.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox E3Log;
        private E3CommonControl e3CommonControl1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPageProject;
        private System.Windows.Forms.TabPage tabPageStructureBrowser;
        private System.Windows.Forms.TabPage tabPageDocListBrowser;
        private E3StructureBrowser e3StructureBrowser1;
        private E3DocListBrowser e3DocListBrowser1;
        private E3ProjectBrowser e3ProjectBrowser1;
    }
}