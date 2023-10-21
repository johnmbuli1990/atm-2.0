using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.DataFormats;

namespace ATM_4._0
{
    public partial class Fonctions_supplémentaires : MetroFramework.Forms.MetroForm
    {
        //Declaring Excel object
        private Excel.Application ExcelApp;
        private Excel.Workbook Workbook;
        private Excel.Worksheet Worksheet;
        private MainATMForm MainATMFormInstance;
        public Fonctions_supplémentaires()
        {
            InitializeComponent();
        }

        private void Fonctions_supplémentaires_Load(object sender, EventArgs e)
        {

        }

        private void ReturnButton_Click(object sender, EventArgs e)
        {
           
        }

        private void syncProvidorButton_Click(object sender, EventArgs e)
        {
            // Get the reference to the existing ComboBox on Form1
            MainATMForm form1 = System.Windows.Forms.Application.OpenForms.OfType<MainATMForm>().FirstOrDefault();
            ComboBox RechangeChoicesCombo = form1.GetComboBox();


            // Clear combobox contents
            RechangeChoicesCombo.Items.Clear();
            // Create a new OpenFileDialog object
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Set the file filter and initial directory
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.xlsb";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // Show the file dialog and check if the user clicked OK
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Get the selected file path
                string filePath = openFileDialog.FileName;

                // Check if the file is already open
                bool isFileOpen = IsFileOpen(filePath);

                if (isFileOpen)
                {
                    MessageBox.Show("File is already open");
                }
                else
                {
                    // Open the Excel file using the default program
                    ExcelApp = new Excel.Application();
                    Workbook = ExcelApp.Workbooks.Open(filePath);
                    Excel.Worksheet worksheet = (Worksheet)Workbook.Worksheets["Material list"];

                    // Find the last used row in the third column
                    Excel.Range lastCell = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                    int lastRow = lastCell.Row;

                    Excel.Range usedRange = worksheet.Range["C4", "C" + lastRow.ToString()];

                    // Get the values from the third column
                    Array values = (Array)usedRange.Value;

                    // Populate the values into the ComboBox
                    // Populate the values into the ComboBox
                    foreach (object value in values)
                    {
                        if (value != null && value.ToString() != "System.DBNull")
                        {
                            RechangeChoicesCombo.Items.Add(value.ToString());
                        }
                    }
                    MessageBox.Show("Items have been added to the ComboBox.", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);


                    // Close the Excel workbook and release resources
                    Workbook.Close();
                    ExcelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
                    worksheet = null;
                    Workbook = null;
                    ExcelApp = null;
                    GC.Collect();

                }
            }
        }

        private bool IsFileOpen(string filePath)
        {
            try
            {
                // Try to open the file with write access
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    // If the file can be opened with write access, it is not already open
                    return false;
                }
            }
            catch (IOException)
            {
                // If the file cannot be opened with write access, it is already open
                return true;
            }
        }
    }
}
