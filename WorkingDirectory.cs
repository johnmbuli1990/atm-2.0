
using ATM_4._0;
using Excel = Microsoft.Office.Interop.Excel;

namespace ATM_3._0
{
    public partial class WorkingDirectory : MetroFramework.Forms.MetroForm
    {
        private string copyFilePath;
        private Excel.Application ExcelApp;
        private Excel.Workbook Workbook;
        private Excel.Worksheet Worksheet;
        private string selectedDirectory2;
        public WorkingDirectory()
        {
            InitializeComponent();
            
        }

        public string selectedDirectory { get; set; }
        private void ATMDirectory_Click(object sender, EventArgs e)
        {

            // Create a new instance of the FolderBrowserDialog class
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            // Set the initial directory to the current directory
            folderBrowserDialog.SelectedPath = Environment.CurrentDirectory;

            // Show the dialog and wait for the user to select a folder
            DialogResult result = folderBrowserDialog.ShowDialog();

            // If the user clicked the OK button, register the selected directory name
            if (result == DialogResult.OK)
            {
                selectedDirectory = folderBrowserDialog.SelectedPath;
                // Do something with the selected directory name, such as display it in a label control
                //selectedDirectoryLabel.Text = selectedDirectory;

                // Open the ATMForm
                MainATMForm form3 = new MainATMForm(selectedDirectory: selectedDirectory);

                //MainATMForm form3 = new MainATMForm(selectedDirectory);
                form3.Show();
                this.Hide();
            }
        }
    }
}
