namespace ATM_3._0
{
    partial class StartUpForm
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
            tableLayoutPanel1 = new TableLayoutPanel();
            NewATMTile = new MetroFramework.Controls.MetroTile();
            ImportATMTile = new MetroFramework.Controls.MetroTile();
            tableLayoutPanel1.SuspendLayout();
            SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 2;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
            tableLayoutPanel1.Controls.Add(NewATMTile, 1, 0);
            tableLayoutPanel1.Controls.Add(ImportATMTile, 0, 0);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(20, 60);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 1;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            tableLayoutPanel1.Size = new Size(760, 120);
            tableLayoutPanel1.TabIndex = 0;
            // 
            // NewATMTile
            // 
            NewATMTile.ActiveControl = null;
            NewATMTile.Dock = DockStyle.Fill;
            NewATMTile.Location = new Point(383, 3);
            NewATMTile.Name = "NewATMTile";
            NewATMTile.Size = new Size(374, 114);
            NewATMTile.TabIndex = 1;
            NewATMTile.Text = "New MTA";
            NewATMTile.UseSelectable = true;
            NewATMTile.Click += NewATMTile_Click;
            // 
            // ImportATMTile
            // 
            ImportATMTile.ActiveControl = null;
            ImportATMTile.Dock = DockStyle.Fill;
            ImportATMTile.Location = new Point(3, 3);
            ImportATMTile.Name = "ImportATMTile";
            ImportATMTile.Size = new Size(374, 114);
            ImportATMTile.TabIndex = 0;
            ImportATMTile.Text = "Import MTA";
            ImportATMTile.UseSelectable = true;
            ImportATMTile.Click += ImportATMTile_Click;
            // 
            // StartUpForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 200);
            Controls.Add(tableLayoutPanel1);
            MaximumSize = new Size(800, 200);
            MinimumSize = new Size(800, 200);
            Name = "StartUpForm";
            Text = "StartUp Form";
            tableLayoutPanel1.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private TableLayoutPanel tableLayoutPanel1;
        private MetroFramework.Controls.MetroTile ImportATMTile;
        private MetroFramework.Controls.MetroTile NewATMTile;
    }
}