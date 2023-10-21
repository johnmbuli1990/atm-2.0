using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ATM_4._0;

namespace ATM_3._0
{
    public partial class StartUpForm : MetroFramework.Forms.MetroForm
    {

        public StartUpForm()
        {
            InitializeComponent();

        }



        private void NewATMTile_Click(object sender, EventArgs e)
        {
            WorkingDirectory form2 = new WorkingDirectory();
            form2.Show();
            this.Hide();
        }

        public string Excelfile { get; set; }
        private void ImportATMTile_Click(object sender, EventArgs e)
        {
            //Open the directory of the existing ATM

            // Create an instance of the OpenFileDialog class
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Set the file filter to only show Excel files
            openFileDialog.Filter = "Excel files (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm";

            // Show the file dialog and check if the user clicked the OK button
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                Excelfile = openFileDialog.FileName;

                // Open the ATMForm
                MainATMForm form3 = new MainATMForm(Excelfile: Excelfile);

                //MainATMForm form3 = new MainATMForm(Excelfile);
                form3.Show();
                this.Hide();

                // Further processing of the Excel file can be done here
            }
        }
    }
}
