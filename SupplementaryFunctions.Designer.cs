namespace ATM_4._0
{
    partial class Fonctions_supplémentaires
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
            ReturnButton = new MetroFramework.Controls.MetroButton();
            metroButton4 = new MetroFramework.Controls.MetroButton();
            syncProvidorButton = new MetroFramework.Controls.MetroButton();
            LanguageButton = new MetroFramework.Controls.MetroButton();
            ResetButton = new MetroFramework.Controls.MetroButton();
            tableLayoutPanel1.SuspendLayout();
            SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 5;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            tableLayoutPanel1.Controls.Add(ReturnButton, 4, 0);
            tableLayoutPanel1.Controls.Add(metroButton4, 3, 0);
            tableLayoutPanel1.Controls.Add(syncProvidorButton, 2, 0);
            tableLayoutPanel1.Controls.Add(LanguageButton, 0, 0);
            tableLayoutPanel1.Controls.Add(ResetButton, 1, 0);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(20, 60);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 2;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 14.822134F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 85.1778641F));
            tableLayoutPanel1.Size = new Size(1025, 506);
            tableLayoutPanel1.TabIndex = 0;
            // 
            // ReturnButton
            // 
            ReturnButton.Dock = DockStyle.Fill;
            ReturnButton.Location = new Point(823, 3);
            ReturnButton.Name = "ReturnButton";
            ReturnButton.Size = new Size(199, 69);
            ReturnButton.TabIndex = 4;
            ReturnButton.Text = "Retour";
            ReturnButton.UseSelectable = true;
            ReturnButton.Click += ReturnButton_Click;
            // 
            // metroButton4
            // 
            metroButton4.Dock = DockStyle.Fill;
            metroButton4.Location = new Point(618, 3);
            metroButton4.Name = "metroButton4";
            metroButton4.Size = new Size(199, 69);
            metroButton4.TabIndex = 3;
            metroButton4.Text = "Reserve";
            metroButton4.UseSelectable = true;
            // 
            // syncProvidorButton
            // 
            syncProvidorButton.Dock = DockStyle.Fill;
            syncProvidorButton.Location = new Point(413, 3);
            syncProvidorButton.Name = "syncProvidorButton";
            syncProvidorButton.Size = new Size(199, 69);
            syncProvidorButton.TabIndex = 2;
            syncProvidorButton.Text = "Sync Providor";
            syncProvidorButton.UseSelectable = true;
            syncProvidorButton.Click += syncProvidorButton_Click;
            // 
            // LanguageButton
            // 
            LanguageButton.Dock = DockStyle.Fill;
            LanguageButton.Location = new Point(3, 3);
            LanguageButton.Name = "LanguageButton";
            LanguageButton.Size = new Size(199, 69);
            LanguageButton.TabIndex = 1;
            LanguageButton.Text = "Français/English";
            LanguageButton.UseSelectable = true;
            // 
            // ResetButton
            // 
            ResetButton.Dock = DockStyle.Fill;
            ResetButton.Location = new Point(208, 3);
            ResetButton.Name = "ResetButton";
            ResetButton.Size = new Size(199, 69);
            ResetButton.TabIndex = 0;
            ResetButton.Text = "Reset";
            ResetButton.UseSelectable = true;
            // 
            // Fonctions_supplémentaires
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1065, 586);
            Controls.Add(tableLayoutPanel1);
            Name = "Fonctions_supplémentaires";
            Text = "Fonctions Supplémentaires";
            tableLayoutPanel1.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private TableLayoutPanel tableLayoutPanel1;
        private MetroFramework.Controls.MetroButton ResetButton;
        private MetroFramework.Controls.MetroButton ReturnButton;
        private MetroFramework.Controls.MetroButton metroButton4;
        private MetroFramework.Controls.MetroButton syncProvidorButton;
        private MetroFramework.Controls.MetroButton LanguageButton;
    }
}