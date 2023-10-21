using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.IO;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using ComponentFactory.Krypton.Toolkit;
using ATM_3._0;
using System.Xml.Linq;
using System.Diagnostics;
using System.Reflection;

namespace ATM_4._0
{
    public partial class MainATMForm : MetroFramework.Forms.MetroForm
    {

        //Declaring Excel object
        private Excel.Application ExcelApp;
        private Excel.Workbook Workbook;
        private Excel.Worksheet Worksheet;
        private string selectedDirectory2;
        //private string Excelfile;
        private string selectedRowId;
        private List<ComboBox> comboBoxes;

        public MainATMForm(string Excelfile = null, string selectedDirectory = null)
        {
            InitializeComponent();
            this.FormClosing += MainATMClosing;

            //selectedDirectory = st;
            //Excelfile = st2;
            if (!string.IsNullOrEmpty(Excelfile))
            {
                RefreshButton.Enabled = false;
                ExcelApp = new Excel.Application();
                Workbook = ExcelApp.Workbooks.Open(Excelfile);
                ExcelApp.Visible = true;

                Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
                if (worksheet != null)
                {
                    //clear comboboxes
                    ClearComboBoxes();

                    int dataStartRow = 2;

                    // Populate combo boxes with cell values and row IDs
                    WireUpComboBoxes();
                    foreach (ComboBox comboBox in comboBoxes)
                    {
                        int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                        AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                        // Wire up the event handler
                        comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                    }
                }
            }
            else if (!string.IsNullOrEmpty(selectedDirectory))
            {
                selectedDirectory2 = selectedDirectory;
            }

            // Attach an event handler to the ButtonSpecClick event
            AddDesButton.Click += AddDesButton_Click;
            AddTacheInduite.Click += AddTacheInduite_Click;
            AddFigureButton.Click += AddFigureButton_Click;
            AddTimeButton.Click += AddTimeButton_Click;

            //Event handlers to activate the operation and note lines -- combo choices and text input
            ActionComboBox.SelectedIndexChanged += ActionComboBox_SelectedIndexChanged;
            PictoComboBox.SelectedIndexChanged += ActionComboBox_SelectedIndexChanged;
            OrderComboBox.SelectedIndexChanged += ActionComboBox_SelectedIndexChanged;
            DesignationComboBox.SelectedIndexChanged += ActionComboBox_SelectedIndexChanged;
            TacheInduiteComboBox.SelectedIndexChanged += ActionComboBox_SelectedIndexChanged;
            FigureComboBox.SelectedIndexChanged += ActionComboBox_SelectedIndexChanged;
            CompetenceComboBox.SelectedIndexChanged += ActionComboBox_SelectedIndexChanged;
            OutillagesComboBox.SelectedIndexChanged += ActionComboBox_SelectedIndexChanged;
            RechangeComboBox.SelectedIndexChanged += ActionComboBox_SelectedIndexChanged;
            ConsommablesComboBox.SelectedIndexChanged += ActionComboBox_SelectedIndexChanged;
            VisserieComboBox.SelectedIndexChanged += ActionComboBox_SelectedIndexChanged;
            TimeComboBox.SelectedIndexChanged += ActionComboBox_SelectedIndexChanged;



            //Update initially active combos based on user's input
            TagPictoChoicesCombo.SelectedIndexChanged += UpdateTagCombos_SelectedIndexChanged;
            OrdreChoicesCombo.SelectedIndexChanged += Update_OrdreCombo;
            ActionChoicesCombo.SelectedIndexChanged += Update_ActionCombo;
            CompetenceChoicesCombo.SelectedIndexChanged += Update_Competence;
            OutillageChoicesCombo.SelectedIndexChanged += Update_Outillage;
            RechangeChoicesCombo.SelectedIndexChanged += Update_Rechange;
            ConsommableChoicesCombo.SelectedIndexChanged += Update_Consommable;
            VisserieChoicesCombo.SelectedIndexChanged += VisserieComboBox_SelectedIndexChanged;


        }

        // Make the ComboBox accessible to other forms
        public ComboBox GetComboBox()
        {
            return RechangeChoicesCombo;
        }

        private void AddTimeButton_Click(object sender, EventArgs e)
        {
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            // Update the selected item in ComboBox1 to match ComboBox2
            if (TimeComboBox.SelectedItem != null)
            {
                //PictoComboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                if (TempsTextbox.Text != null)
                {
                    // Get the row index
                    KeyValuePair<string, int> selectedItem = (KeyValuePair<string, int>)TacheInduiteComboBox.SelectedItem;
                    int selectedRowIndex = selectedItem.Value;


                    // Get the worksheet and the range of the selected row
                    Excel.Worksheet xlWorksheet = (Excel.Worksheet)Workbook.Sheets["DESCRIPTION_TACHE"];
                    Excel.Range selectedRowRange = (Excel.Range)xlWorksheet.Rows[selectedRowIndex];

                    //Add the selected item in excel
                    ((Excel.Range)xlWorksheet.Cells[selectedRowIndex, 12]).Value = TempsTextbox.Text;

                    // Save the changes
                    Workbook.Save();

                    // Clear the comboboxes
                    ClearComboBoxes();

                    // Update the combo boxes with the new data
                    int dataStartRow = 2;

                    // Populate combo boxes with cell values and row IDs
                    WireUpComboBoxes();
                    foreach (ComboBox comboBox in comboBoxes)
                    {
                        int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                        AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                        // Wire up the event handler
                        comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                    }
                    //refresh the tag combo box of the selected row

                    PictoComboBox.SelectedIndex = selectedRowIndex - 2;
                    TempsTextbox.Text = null;
                }
                else
                {
                    // Handle the case when no item is selected in the ComboBox
                    Console.WriteLine("No item selected in Picto");
                }

            }
        }

        private void AddFigureButton_Click(object? sender, EventArgs e)
        {
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            // Update the selected item in ComboBox1 to match ComboBox2
            if (FigureComboBox.SelectedItem != null)
            {
                //Open file directory and add a figure in a new tab


                //PictoComboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                if (AddFigureText.Text != null)
                {
                    // Get the row index
                    KeyValuePair<string, int> selectedItem = (KeyValuePair<string, int>)FigureComboBox.SelectedItem;
                    int selectedRowIndex = selectedItem.Value;


                    // Get the worksheet and the range of the selected row
                    Excel.Worksheet xlWorksheet = (Excel.Worksheet)Workbook.Sheets["DESCRIPTION_TACHE"];
                    Excel.Range selectedRowRange = (Excel.Range)xlWorksheet.Rows[selectedRowIndex];

                    //Add the selected item in excel
                    ((Excel.Range)xlWorksheet.Cells[selectedRowIndex, 6]).Value = AddFigureText.Text;

                    // Open file directory dialog to choose an image file
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.bmp";
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string imagePath = openFileDialog.FileName;


                        // Create a new worksheet for the image
                        Excel.Worksheet imageSheet = (Excel.Worksheet)Workbook.Sheets.Add(After: worksheet);
                        imageSheet.Name = AddFigureText.Text;

                        // Add the image to the worksheet
                        Excel.Range imageCell = (Excel.Range)imageSheet.Cells[1, 1];
                        Excel.Pictures pictures = (Excel.Pictures)imageSheet.Pictures();
                        Excel.Picture picture = pictures.Insert(imagePath);
                        picture.Left = (double)imageCell.Left;
                        picture.Top = (double)imageCell.Top;
                        picture.Width = 200;
                        picture.Height = 200;

                        // Add a hyperlink from the figure to the added tab
                        string address = string.Format("#{0}!A1", AddFigureText.Text);
                        Excel.Hyperlink hyperlink = (Excel.Hyperlink)xlWorksheet.Hyperlinks.Add(
                            Anchor: ((Excel.Range)xlWorksheet.Cells[selectedRowIndex, 6]),
                            Address: string.Empty,
                            SubAddress: address,
                            TextToDisplay: AddFigureText.Text
                        );
                        hyperlink.Range.Hyperlinks.Add(hyperlink.Range, address);

                        MessageBox.Show("Une figure a été ajoutée");

                    }

                    // Save the changes
                    Workbook.Save();

                    // Clear the comboboxes
                    ClearComboBoxes();

                    // Update the combo boxes with the new data
                    int dataStartRow = 2;

                    // Populate combo boxes with cell values and row IDs
                    WireUpComboBoxes();
                    foreach (ComboBox comboBox in comboBoxes)
                    {
                        int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                        AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                        // Wire up the event handler
                        comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                    }
                    //refresh the tag combo box of the selected row

                    PictoComboBox.SelectedIndex = selectedRowIndex - 2;
                    AddFigureText.Text = null;
                }
                else
                {
                    MessageBox.Show("Ajouter le nom de la figure d'abord");
                    return;
                }

            }
        }

        private void AddTacheInduite_Click(object? sender, EventArgs e)
        {
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            // Update the selected item in ComboBox1 to match ComboBox2
            if (TacheInduiteComboBox.SelectedItem != null)
            {
                //PictoComboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                if (TacheInduiteTextBox.Text != null)
                {
                    // Get the row index
                    KeyValuePair<string, int> selectedItem = (KeyValuePair<string, int>)TacheInduiteComboBox.SelectedItem;
                    int selectedRowIndex = selectedItem.Value;


                    // Get the worksheet and the range of the selected row
                    Excel.Worksheet xlWorksheet = (Excel.Worksheet)Workbook.Sheets["DESCRIPTION_TACHE"];
                    Excel.Range selectedRowRange = (Excel.Range)xlWorksheet.Rows[selectedRowIndex];

                    //Add the selected item in excel
                    ((Excel.Range)xlWorksheet.Cells[selectedRowIndex, 5]).Value = TacheInduiteTextBox.Text;

                    // Save the changes
                    Workbook.Save();

                    // Clear the comboboxes
                    ClearComboBoxes();

                    // Update the combo boxes with the new data
                    int dataStartRow = 2;

                    // Populate combo boxes with cell values and row IDs
                    WireUpComboBoxes();
                    foreach (ComboBox comboBox in comboBoxes)
                    {
                        int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                        AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                        // Wire up the event handler
                        comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                    }
                    //refresh the tag combo box of the selected row

                    PictoComboBox.SelectedIndex = selectedRowIndex - 2;
                    TacheInduiteTextBox.Text = null;
                }
                else
                {
                    // Handle the case when no item is selected in the ComboBox
                    Console.WriteLine("No item selected in Picto");
                }

            }
        }

        private void AddDesButton_Click(object? sender, EventArgs e)
        {
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            // Update the selected item in ComboBox1 to match ComboBox2
            if (DesignationComboBox.SelectedItem != null)
            {
                //PictoComboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                if (DesignationTextBox.Text != null)
                {
                    // Get the row index
                    KeyValuePair<string, int> selectedItem = (KeyValuePair<string, int>)DesignationComboBox.SelectedItem;
                    int selectedRowIndex = selectedItem.Value;


                    // Get the worksheet and the range of the selected row
                    Excel.Worksheet xlWorksheet = (Excel.Worksheet)Workbook.Sheets["DESCRIPTION_TACHE"];
                    Excel.Range selectedRowRange = (Excel.Range)xlWorksheet.Rows[selectedRowIndex];

                    //Add the selected item in excel
                    ((Excel.Range)xlWorksheet.Cells[selectedRowIndex, 4]).Value = DesignationTextBox.Text;

                    // Save the changes
                    Workbook.Save();

                    // Clear the comboboxes
                    ClearComboBoxes();

                    // Update the combo boxes with the new data
                    int dataStartRow = 2;

                    // Populate combo boxes with cell values and row IDs
                    WireUpComboBoxes();
                    foreach (ComboBox comboBox in comboBoxes)
                    {
                        int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                        AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                        // Wire up the event handler
                        comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                    }
                    //refresh the tag combo box of the selected row
                    /*
                     * 1. get the desired combo box
                     * 2. simulate selection 
                     */
                    PictoComboBox.SelectedIndex = selectedRowIndex - 2;
                    DesignationTextBox.Text = null;
                }
                else
                {
                    // Handle the case when no item is selected in the ComboBox
                    Console.WriteLine("No item selected in Picto");
                }

            }
        }

        public string copyFilePath { get; set; }
        public string SelectedDirectory2 { get; }

        private void button1_Click(object sender, EventArgs e)
        {

        }
        private void ActionComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (ActionComboBox.SelectedItem != null && DesignationComboBox.SelectedItem != null)
            {
                string selectedItem = ActionComboBox.SelectedItem.ToString();
                string selectedItem2 = DesignationComboBox.SelectedItem.ToString();
                string[] substrings = { "Mise en sécurité", "Opérations proprement dites", "Opérations complémentaires" };
                TagPictoChoicesCombo.Enabled = true;
                OrdreChoicesCombo.Enabled = true;
                ActionChoicesCombo.Enabled = true;
                DesignationTextBox.Enabled = true;
                TacheInduiteTextBox.Enabled = true;
                AddFigureText.Enabled = true;
                CompetenceChoicesCombo.Enabled = true;
                OutillageChoicesCombo.Enabled = true;
                RechangeChoicesCombo.Enabled = true;
                ConsommableChoicesCombo.Enabled = true;
                VisserieChoicesCombo.Enabled = true;
                TempsTextbox.Enabled = true;


                if (selectedItem.Contains("SOUS"))
                {

                    TagPictoChoicesCombo.Enabled = false;
                    //OrdreInputText.Enabled = true;
                    ActionChoicesCombo.Enabled = false;
                    DesignationTextBox.Enabled = false;
                    TacheInduiteTextBox.Enabled = false;
                    AddFigureText.Enabled = false;
                    CompetenceChoicesCombo.Enabled = false;
                    OutillageChoicesCombo.Enabled = false;
                    RechangeChoicesCombo.Enabled = false;
                    ConsommableChoicesCombo.Enabled = false;
                    VisserieChoicesCombo.Enabled = false;
                    TempsTextbox.Enabled = false;

                }
                else if (selectedItem.Contains("Note"))
                {
                    //TagPictoChoicesCombo.Enabled = false;
                    OrdreChoicesCombo.Enabled = false;
                    ActionChoicesCombo.Enabled = false;
                    //DesignationTextBox.Enabled = false;
                    //TacheInduiteTextBox.Enabled = false;
                    //AddFigureText.Enabled = false;
                    //CompetenceChoicesCombo.Enabled = false;
                    //OutillageChoicesCombo.Enabled = false;
                    //RechangeChoicesCombo.Enabled = false;
                    //ConsommableChoicesCombo.Enabled = false;
                    //VisserieChoicesCombo.Enabled = false;
                    //TempsTextbox.Enabled = false;
                }


                foreach (string substring in substrings)
                {
                    if (selectedItem2.Contains(substring))
                    {
                        TagPictoChoicesCombo.Enabled = false;
                        OrdreChoicesCombo.Enabled = false;
                        ActionChoicesCombo.Enabled = false;
                        DesignationTextBox.Enabled = false;
                        TacheInduiteTextBox.Enabled = false;
                        AddFigureText.Enabled = false;
                        CompetenceChoicesCombo.Enabled = false;
                        OutillageChoicesCombo.Enabled = false;
                        RechangeChoicesCombo.Enabled = false;
                        ConsommableChoicesCombo.Enabled = false;
                        VisserieChoicesCombo.Enabled = false;
                        //TempsTextbox.Enabled = false;

                    }
                }


            }

        }

        private void UpdateTagCombos_SelectedIndexChanged(object sender, EventArgs e)
        {
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            // Update the selected item in ComboBox1 to match ComboBox2
            if (TagPictoChoicesCombo.SelectedItem != null)
            {
                //PictoComboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                if (PictoComboBox.SelectedItem != null)
                {
                    // Get the row index
                    KeyValuePair<string, int> selectedItem = (KeyValuePair<string, int>)PictoComboBox.SelectedItem;
                    int selectedRowIndex = selectedItem.Value;


                    // Get the worksheet and the range of the selected row
                    Excel.Worksheet xlWorksheet = (Excel.Worksheet)Workbook.Sheets["DESCRIPTION_TACHE"];
                    Excel.Range selectedRowRange = (Excel.Range)xlWorksheet.Rows[selectedRowIndex];

                    //Add the selected item in excel
                    ((Excel.Range)xlWorksheet.Cells[selectedRowIndex, 1]).Value = TagPictoChoicesCombo.SelectedItem;

                    // Save the changes
                    Workbook.Save();

                    // Clear the comboboxes
                    ClearComboBoxes();

                    // Update the combo boxes with the new data
                    int dataStartRow = 2;

                    // Populate combo boxes with cell values and row IDs
                    WireUpComboBoxes();
                    foreach (ComboBox comboBox in comboBoxes)
                    {
                        int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                        AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                        // Wire up the event handler
                        comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                    }
                    //refresh the tag combo box of the selected row
                    /*
                     * 1. get the desired combo box
                     * 2. simulate selection 
                     */
                    PictoComboBox.SelectedIndex = selectedRowIndex - 2;
                    TagPictoChoicesCombo.SelectedItem = null;
                }
                else
                {
                    // Handle the case when no item is selected in the ComboBox
                    Console.WriteLine("No item selected in Picto");
                }

            }

        }

        private void Update_OrdreCombo(object sender, EventArgs e)
        {
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            if (OrdreChoicesCombo.SelectedItem != null)
            {
                // OrderComboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                if (OrderComboBox.SelectedItem != null)
                {
                    // Get the row index
                    KeyValuePair<string, int> selectedItem = (KeyValuePair<string, int>)OrderComboBox.SelectedItem;
                    int selectedRowIndex = selectedItem.Value;

                    // Get the worksheet and the range of the selected row
                    Excel.Worksheet xlWorksheet = (Excel.Worksheet)Workbook.Sheets["DESCRIPTION_TACHE"];
                    Excel.Range selectedRowRange = (Excel.Range)xlWorksheet.Rows[selectedRowIndex];

                    //Add the selected item in excel
                    ((Excel.Range)xlWorksheet.Cells[selectedRowIndex, 2]).Value = OrdreChoicesCombo.SelectedItem;

                    //update the order combo
                    //OrderComboBox.SelectedItem = OrdreChoicesCombo.SelectedItem;
                    // Save the changes
                    Workbook.Save();

                    // Clear the comboboxes
                    ClearComboBoxes();

                    // Update the combo boxes with the new data
                    int dataStartRow = 2;

                    // Populate combo boxes with cell values and row IDs
                    WireUpComboBoxes();
                    foreach (ComboBox comboBox in comboBoxes)
                    {
                        int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                        AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                        // Wire up the event handler
                        comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                    }
                    PictoComboBox.SelectedIndex = selectedRowIndex - 2;
                    OrdreChoicesCombo.SelectedItem = null;
                }
                else
                {
                    // Handle the case when no item is selected in the ComboBox
                    Console.WriteLine("No item selected in Order Combo box");
                }

            }
        }

        private void Update_ActionCombo(object sender, EventArgs e)
        {
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            if (ActionChoicesCombo.SelectedItem != null)
            {
                if (ActionComboBox.SelectedItem != null)
                {
                    // Get the row index
                    KeyValuePair<string, int> selectedItem = (KeyValuePair<string, int>)ActionComboBox.SelectedItem;
                    int selectedRowIndex = selectedItem.Value;

                    // Get the worksheet and the range of the selected row
                    Excel.Worksheet xlWorksheet = (Excel.Worksheet)Workbook.Sheets["DESCRIPTION_TACHE"];
                    Excel.Range selectedRowRange = (Excel.Range)xlWorksheet.Rows[selectedRowIndex];

                    //Add the selected item in excel
                    ((Excel.Range)xlWorksheet.Cells[selectedRowIndex, 3]).Value = ActionChoicesCombo.SelectedItem;

                    //update the order combo
                    //OrderComboBox.SelectedItem = OrdreChoicesCombo.SelectedItem;
                    // Save the changes
                    Workbook.Save();

                    // Clear the comboboxes
                    ClearComboBoxes();

                    // Update the combo boxes with the new data
                    int dataStartRow = 2;

                    // Populate combo boxes with cell values and row IDs
                    WireUpComboBoxes();
                    foreach (ComboBox comboBox in comboBoxes)
                    {
                        int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                        AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                        // Wire up the event handler
                        comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                    }
                    PictoComboBox.SelectedIndex = selectedRowIndex - 2;
                    ActionChoicesCombo.SelectedItem = null;
                }
                else
                {
                    // Handle the case when no item is selected in the ComboBox
                    Console.WriteLine("No item selected in Action Combo box");
                }

            }
        }
        private void Update_Competence(object sender, EventArgs e)
        {
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            if (CompetenceChoicesCombo.SelectedItem != null)
            {

                if (CompetenceComboBox.SelectedItem != null)
                {
                    // Get the row index
                    KeyValuePair<string, int> selectedItem = (KeyValuePair<string, int>)CompetenceComboBox.SelectedItem;
                    int selectedRowIndex = selectedItem.Value;

                    // Get the worksheet and the range of the selected row
                    Excel.Worksheet xlWorksheet = (Excel.Worksheet)Workbook.Sheets["DESCRIPTION_TACHE"];
                    Excel.Range selectedRowRange = (Excel.Range)xlWorksheet.Rows[selectedRowIndex];

                    //Add the selected item in excel
                    ((Excel.Range)xlWorksheet.Cells[selectedRowIndex, 7]).Value = CompetenceChoicesCombo.SelectedItem;

                    //update the order combo
                    //OrderComboBox.SelectedItem = OrdreChoicesCombo.SelectedItem;
                    // Save the changes
                    Workbook.Save();

                    // Clear the comboboxes
                    ClearComboBoxes();

                    // Update the combo boxes with the new data
                    int dataStartRow = 2;

                    // Populate combo boxes with cell values and row IDs
                    WireUpComboBoxes();
                    foreach (ComboBox comboBox in comboBoxes)
                    {
                        int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                        AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                        // Wire up the event handler
                        comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                    }
                    PictoComboBox.SelectedIndex = selectedRowIndex - 2;
                    CompetenceChoicesCombo.SelectedItem = null;
                }
                else
                {
                    // Handle the case when no item is selected in the ComboBox
                    Console.WriteLine("No item selected in Action Combo box");
                }

            }
        }

        private void Update_Outillage(object sender, EventArgs e)
        {
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            if (OutillageChoicesCombo.SelectedItem != null)
            {

                if (OutillagesComboBox.SelectedItem != null)
                {
                    // Get the row index
                    KeyValuePair<string, int> selectedItem = (KeyValuePair<string, int>)OutillagesComboBox.SelectedItem;
                    int selectedRowIndex = selectedItem.Value;

                    // Get the worksheet and the range of the selected row
                    Excel.Worksheet xlWorksheet = (Excel.Worksheet)Workbook.Sheets["DESCRIPTION_TACHE"];
                    Excel.Range selectedRowRange = (Excel.Range)xlWorksheet.Rows[selectedRowIndex];

                    //Add the selected item in excel
                    ((Excel.Range)xlWorksheet.Cells[selectedRowIndex, 8]).Value = OutillageChoicesCombo.SelectedItem;

                    //update the order combo
                    //OrderComboBox.SelectedItem = OrdreChoicesCombo.SelectedItem;
                    // Save the changes
                    Workbook.Save();

                    // Clear the comboboxes
                    ClearComboBoxes();

                    // Update the combo boxes with the new data
                    int dataStartRow = 2;

                    // Populate combo boxes with cell values and row IDs
                    WireUpComboBoxes();
                    foreach (ComboBox comboBox in comboBoxes)
                    {
                        int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                        AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                        // Wire up the event handler
                        comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                    }
                    PictoComboBox.SelectedIndex = selectedRowIndex - 2;
                    OutillageChoicesCombo.SelectedItem = null;
                }
                else
                {
                    // Handle the case when no item is selected in the ComboBox
                    Console.WriteLine("No item selected in Action Combo box");
                }

            }
        }

        private void Update_Rechange(object sender, EventArgs e)
        {
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            if (RechangeChoicesCombo.SelectedItem != null)
            {

                if (RechangeComboBox.SelectedItem != null)
                {
                    // Get the row index
                    KeyValuePair<string, int> selectedItem = (KeyValuePair<string, int>)RechangeComboBox.SelectedItem;
                    int selectedRowIndex = selectedItem.Value;

                    // Get the worksheet and the range of the selected row
                    Excel.Worksheet xlWorksheet = (Excel.Worksheet)Workbook.Sheets["DESCRIPTION_TACHE"];
                    Excel.Range selectedRowRange = (Excel.Range)xlWorksheet.Rows[selectedRowIndex];

                    //Add the selected item in excel
                    ((Excel.Range)xlWorksheet.Cells[selectedRowIndex, 9]).Value = RechangeChoicesCombo.SelectedItem;

                    //update the order combo
                    //OrderComboBox.SelectedItem = OrdreChoicesCombo.SelectedItem;
                    // Save the changes
                    Workbook.Save();

                    // Clear the comboboxes
                    ClearComboBoxes();

                    // Update the combo boxes with the new data
                    int dataStartRow = 2;

                    // Populate combo boxes with cell values and row IDs
                    WireUpComboBoxes();
                    foreach (ComboBox comboBox in comboBoxes)
                    {
                        int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                        AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                        // Wire up the event handler
                        comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                    }
                    PictoComboBox.SelectedIndex = selectedRowIndex - 2;
                    RechangeChoicesCombo.SelectedItem = null;
                }
                else
                {
                    // Handle the case when no item is selected in the ComboBox
                    Console.WriteLine("No item selected in Action Combo box");
                }

            }
        }

        private void Update_Consommable(object sender, EventArgs e)
        {
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            if (ConsommableChoicesCombo.SelectedItem != null)
            {

                if (ConsommablesComboBox.SelectedItem != null)
                {
                    // Get the row index
                    KeyValuePair<string, int> selectedItem = (KeyValuePair<string, int>)ConsommablesComboBox.SelectedItem;
                    int selectedRowIndex = selectedItem.Value;

                    // Get the worksheet and the range of the selected row
                    Excel.Worksheet xlWorksheet = (Excel.Worksheet)Workbook.Sheets["DESCRIPTION_TACHE"];
                    Excel.Range selectedRowRange = (Excel.Range)xlWorksheet.Rows[selectedRowIndex];

                    //Add the selected item in excel
                    ((Excel.Range)xlWorksheet.Cells[selectedRowIndex, 10]).Value = ConsommableChoicesCombo.SelectedItem;

                    //update the order combo
                    //OrderComboBox.SelectedItem = OrdreChoicesCombo.SelectedItem;
                    // Save the changes
                    Workbook.Save();

                    // Clear the comboboxes
                    ClearComboBoxes();

                    // Update the combo boxes with the new data
                    int dataStartRow = 2;

                    // Populate combo boxes with cell values and row IDs
                    WireUpComboBoxes();
                    foreach (ComboBox comboBox in comboBoxes)
                    {
                        int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                        AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                        // Wire up the event handler
                        comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                    }
                    PictoComboBox.SelectedIndex = selectedRowIndex - 2;
                    ConsommableChoicesCombo.SelectedItem = null;
                }
                else
                {
                    // Handle the case when no item is selected in the ComboBox
                    Console.WriteLine("No item selected in Action Combo box");
                }

            }
        }

        private void VisserieComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            if (VisserieChoicesCombo.SelectedItem != null)
            {

                if (VisserieComboBox.SelectedItem != null)
                {
                    // Get the row index
                    KeyValuePair<string, int> selectedItem = (KeyValuePair<string, int>)VisserieComboBox.SelectedItem;
                    int selectedRowIndex = selectedItem.Value;

                    // Get the worksheet and the range of the selected row
                    Excel.Worksheet xlWorksheet = (Excel.Worksheet)Workbook.Sheets["DESCRIPTION_TACHE"];
                    Excel.Range selectedRowRange = (Excel.Range)xlWorksheet.Rows[selectedRowIndex];

                    //Add the selected item in excel
                    ((Excel.Range)xlWorksheet.Cells[selectedRowIndex, 11]).Value = VisserieChoicesCombo.SelectedItem;

                    //update the order combo
                    //OrderComboBox.SelectedItem = OrdreChoicesCombo.SelectedItem;
                    // Save the changes
                    Workbook.Save();

                    // Clear the comboboxes
                    ClearComboBoxes();

                    // Update the combo boxes with the new data
                    int dataStartRow = 2;

                    // Populate combo boxes with cell values and row IDs
                    WireUpComboBoxes();
                    foreach (ComboBox comboBox in comboBoxes)
                    {
                        int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                        AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                        // Wire up the event handler
                        comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                    }
                    PictoComboBox.SelectedIndex = selectedRowIndex - 2;
                    VisserieChoicesCombo.SelectedItem = null;
                }
                else
                {
                    // Handle the case when no item is selected in the ComboBox
                    Console.WriteLine("No item selected in Action Combo box");
                }

            }
        }

        // Event handler for the ButtonSpecClick event

        private void RefreshButton_Click(object sender, EventArgs e)
        {
            if (selectedDirectory2 == null)
            {
                MessageBox.Show("veuillez d'abord choisir le répertoire de travail", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string resourceName = "ATM_4._0.AtmTemplate.xlsx";
            string filePath = Assembly.GetExecutingAssembly().GetName().Name + "." + resourceName;
            copyFilePath = selectedDirectory2 + @"\AtmTemplateExample.xlsx";

            try
            {
                if (File.Exists(copyFilePath))
                {
                    // Generate a new file name
                    string directory = Path.GetDirectoryName(copyFilePath);
                    string fileName = Path.GetFileNameWithoutExtension(copyFilePath);
                    string extension = Path.GetExtension(copyFilePath);
                    string newFileName = fileName + "_duplicate" + extension;
                    copyFilePath = Path.Combine(directory, newFileName);
                    return;
                }


                using (Stream resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    if (resourceStream != null)
                    {
                        using (FileStream fileStream = File.Create(copyFilePath))
                        {
                            resourceStream.CopyTo(fileStream);
                        }
                    }
                }


                ExcelApp = new Excel.Application();
                Workbook = ExcelApp.Workbooks.Open(copyFilePath);
                ExcelApp.Visible = true;

                Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
                if (worksheet != null)
                {
                    ClearComboBoxes();
                    int dataStartRow = 2;

                    WireUpComboBoxes();
                    foreach (ComboBox comboBox in comboBoxes)
                    {
                        int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                        AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                        comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                    }
                }
                else
                {
                    MessageBox.Show("The worksheet 'DESCRIPTION_TACHE' does not exist in the workbook.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show("The Excel file is already open. Please close it and try again.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Excel.Worksheet GetWorksheetByName(Excel.Workbook workbook, string sheetName)
        {
            try
            {
                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name == sheetName)
                    {
                        return worksheet;
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle the exception (e.g., log, display an error message, etc.)
                MessageBox.Show("The Excel file is closed.", "Info",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return null;
        }


        private List<ComboBox> GetComboBoxes()
        {
            return new List<ComboBox>()
    {
        PictoComboBox, OrderComboBox, ActionComboBox, DesignationComboBox, TacheInduiteComboBox,
        FigureComboBox, CompetenceComboBox, OutillagesComboBox, RechangeComboBox,
        ConsommablesComboBox, VisserieComboBox, TimeComboBox
    };
        }


        private void AddCellValuesToComboBox(ComboBox comboBox, Excel.Worksheet worksheet, int column, int startRow)
        {
            for (int row = startRow; row <= worksheet.UsedRange.Rows.Count; row++)
            {
                Excel.Range cellRange = (Excel.Range)worksheet.Cells[row, column];
                object cellValue = cellRange.Value;

                string valueToAdd = (cellValue != null) ? cellValue.ToString() : string.Empty;
                int rowId = row; // Get the row ID

                // Create a KeyValuePair with the value and row ID
                KeyValuePair<string, int> item = new KeyValuePair<string, int>(valueToAdd, rowId);
                comboBox.Items.Add(item);

            }
        }


        private void WireUpComboBoxes()
        {
            comboBoxes = new List<ComboBox>()
            {
                PictoComboBox, OrderComboBox, ActionComboBox, DesignationComboBox, TacheInduiteComboBox,
                FigureComboBox, CompetenceComboBox, OutillagesComboBox, RechangeComboBox,
                ConsommablesComboBox, VisserieComboBox, TimeComboBox
            };

            foreach (var comboBox in comboBoxes)
            {
                comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
            }
        }


        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox selectedComboBox = (ComboBox)sender;
            int selectedRowIndex = selectedComboBox.SelectedIndex;
            //string selectedText = selectedComboBox.GetItemText(selectedComboBox.SelectedItem);

            // Synchronize the selection in other comboboxes
            foreach (ComboBox comboBox in comboBoxes)
            {
                if (comboBox != selectedComboBox)
                {
                    comboBox.SelectedIndex = selectedRowIndex;
                    //comboBox.SelectedItem = selectedText;
                }


            }

        }


        private void ComboBox_DropDownClosed(object sender, EventArgs e)
        {
            ComboBox selectedComboBox = (ComboBox)sender;

            // Get the selected row ID from the Tag property
            if (selectedComboBox.Tag != null && int.TryParse(selectedComboBox.Tag.ToString(), out int rowId))
            {
                // Get the selected item
                object selectedItem = selectedComboBox.SelectedItem;

                // Iterate through all the comboboxes
                foreach (ComboBox comboBox in GetComboBoxes())
                {
                    // Set the selected item of each combobox to the corresponding row ID
                    if (comboBox.Tag != null && int.TryParse(comboBox.Tag.ToString(), out int comboBoxRowId) && comboBoxRowId == rowId)
                    {
                        comboBox.SelectedItem = selectedItem;
                    }
                }
            }
        }


        private void AtmSave_Click(object sender, EventArgs e)
        {
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            KeyValuePair<string, int> selectedItem = (KeyValuePair<string, int>)PictoComboBox.SelectedItem;
            int selectedRowIndex = selectedItem.Value;
            // Call your desired method or perform the necessary actions here
            // Clear the comboboxes
            ClearComboBoxes();

            // Update the combo boxes with the new data
            int dataStartRow = 2;

            // Populate combo boxes with cell values and row IDs
            WireUpComboBoxes();
            foreach (ComboBox comboBox in comboBoxes)
            {
                int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                // Wire up the event handler
                comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
            }
            PictoComboBox.SelectedIndex = selectedRowIndex - 2;
        }

        private void SyncProvidorButton_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void ClearComboBoxes()
        {
            PictoComboBox.Items.Clear();
            OrderComboBox.Items.Clear();
            ActionComboBox.Items.Clear();
            DesignationComboBox.Items.Clear();
            TacheInduiteComboBox.Items.Clear();
            FigureComboBox.Items.Clear();
            CompetenceComboBox.Items.Clear();
            OutillagesComboBox.Items.Clear();
            RechangeComboBox.Items.Clear();
            ConsommablesComboBox.Items.Clear();
            VisserieComboBox.Items.Clear();
            TimeComboBox.Items.Clear();
        }




        private List<string> GetComboBoxDataSource(Excel.Worksheet worksheet)
        {
            List<string> comboBoxDataSource = new List<string>();
            Excel.Range lastCell = (Excel.Range)worksheet.Cells[worksheet.Rows.Count, 2];
            Excel.Range usedRange = worksheet.Range[worksheet.Cells[1, 1], lastCell];
            Excel.Range lastUsedCell = usedRange.End[Excel.XlDirection.xlUp];
            int lastUsedRow = lastUsedCell.Row;

            for (int i = 2; i <= lastUsedRow; i++)
            {
                string value = ((Excel.Range)worksheet.Cells[i, 2]).Value?.ToString();
                if (!string.IsNullOrEmpty(value))
                {
                    comboBoxDataSource.Add(value);
                }
            }

            return comboBoxDataSource;
        }

        private void MainATMClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                // Release Excel resources
                Workbook.Close();
                ExcelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
                Worksheet = null;
                Workbook = null;
                ExcelApp = null;
                GC.Collect();

                // Terminate the Excel process forcefully
                Process[] excelProcesses = Process.GetProcessesByName("excel");
                foreach (Process process in excelProcesses)
                {
                    process.Kill();
                }
            }
            catch (Exception ex)
            {
                // Handle the exception or display an error message
                //MessageBox.Show("An error occurred while releasing Excel resources: " + ex.Message);
            }
        }



        private void tableLayoutPanel1_Paint_1(object sender, PaintEventArgs e)
        {

        }



        private void DesignationTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void AddInducedTaskTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void SupFuctionsButton_Click(object sender, EventArgs e)
        {
            Fonctions_supplémentaires form3 = new Fonctions_supplémentaires();
            form3.TopMost = true;
            form3.Show();
            //this.Hide();
        }

        private void GoToImportATM_Click(object sender, EventArgs e)
        {
            try
            {
                // Check if the "DESCRIPTION_TACHE" worksheet is already open
                Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
                if (worksheet == null)
                {
                    // The worksheet is not open, show the StartUpForm
                    StartUpForm form3 = new StartUpForm();
                    form3.Show();
                    this.Hide();
                }
                else
                {
                    // The worksheet is already open, display a message box
                    MessageBox.Show("Fermez l'ATM d'abord");
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions that occurred
                MessageBox.Show("An error occurred: " + ex.Message);
            }

        }

        //Add a subchapter
        private void button5_Click_1(object sender, EventArgs e)
        {
            // Check if an Excel file is already open
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            if (worksheet == null)
            {
                MessageBox.Show("Excel file is not open.");
                return;
            }

            // Get the selected row index from the currently selected item in the combobox
            int selectedRowIndex = OrderComboBox.SelectedIndex;
            if (selectedRowIndex < 0)
            {
                MessageBox.Show("Please select a row first.");
                return;
            }

            // Get the worksheet and the range of the selected row
            Excel.Worksheet xlWorksheet = (Excel.Worksheet)Workbook.Sheets["DESCRIPTION_TACHE"];
            Excel.Range selectedRowRange = (Excel.Range)xlWorksheet.Rows[selectedRowIndex + 2];


            // Insert a new row below the selected row
            selectedRowRange.Offset[1].Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            // Get the inserted row
            Excel.Range insertedRow = (Excel.Range)xlWorksheet.Rows[selectedRowIndex + 3];

            // Set the interior color of the inserted row
            insertedRow.Interior.ColorIndex = 0;



            // Set the value of the new row in column 2 (B)
            ((Excel.Range)xlWorksheet.Cells[selectedRowIndex + 3, 2]).Value = "";

            // Set the value of the new row in column 3 (C)
            ((Excel.Range)xlWorksheet.Cells[selectedRowIndex + 3, 3]).Value = "SOUS CHAPITRE";

            // Save the changes
            Workbook.Save();

            // Clear the comboboxes
            ClearComboBoxes();

            // Update the combo boxes with the new data
            int dataStartRow = 2;

            // Populate combo boxes with cell values and row IDs
            WireUpComboBoxes();
            foreach (ComboBox comboBox in comboBoxes)
            {
                int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                // Wire up the event handler
                comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
            }
            PictoComboBox.SelectedIndex = selectedRowIndex + 1;
        }

        //Add note
        private void NoteButton_Click_1(object sender, EventArgs e)
        {
            // Check if an Excel file is already open
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            if (worksheet == null)
            {
                MessageBox.Show("Excel file is not open.");
                return;
            }

            // Get the selected row index from the currently selected item in the combobox
            int selectedRowIndex = OrderComboBox.SelectedIndex;
            if (selectedRowIndex < 0)
            {
                MessageBox.Show("Please select a row first.");
                return;
            }

            // Get the worksheet and the range of the selected row
            Excel.Worksheet xlWorksheet = (Excel.Worksheet)Workbook.Sheets["DESCRIPTION_TACHE"];
            Excel.Range selectedRowRange = (Excel.Range)xlWorksheet.Rows[selectedRowIndex + 2];

            // Insert a new row below the selected row
            selectedRowRange.Offset[1].Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            // Get the inserted row
            Excel.Range insertedRow = (Excel.Range)xlWorksheet.Rows[selectedRowIndex + 3];

            // Set the interior color of the inserted row
            insertedRow.Interior.ColorIndex = 0;

            // Set the value of the new row in column 2 (B)
            ((Excel.Range)xlWorksheet.Cells[selectedRowIndex + 3, 2]).Value = "NOTE";

            // Set the value of the new row in column 3 (C)
            ((Excel.Range)xlWorksheet.Cells[selectedRowIndex + 3, 3]).Value = "Note";

            // Save the changes
            Workbook.Save();

            // Clear the comboboxes
            ClearComboBoxes();

            // Update the combo boxes with the new data
            int dataStartRow = 2;

            // Populate combo boxes with cell values and row IDs
            WireUpComboBoxes();
            foreach (ComboBox comboBox in comboBoxes)
            {
                int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                // Wire up the event handler
                comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
            }
            PictoComboBox.SelectedIndex = selectedRowIndex + 1;
        }

        //Add a step
        private void OperationButton_Click_1(object sender, EventArgs e)
        {
            // Check if an Excel file is already open
            Excel.Worksheet worksheet = GetWorksheetByName(Workbook, "DESCRIPTION_TACHE");
            if (worksheet == null)
            {
                MessageBox.Show("Please open an Excel file first.");
                return;
            }

            // Get the selected row index from the currently selected item in the combobox
            int selectedRowIndex = OrderComboBox.SelectedIndex;
            if (selectedRowIndex < 0)
            {
                MessageBox.Show("Please select a row first.");
                return;
            }

            // Get the worksheet and the range of the selected row
            Excel.Worksheet xlWorksheet = (Excel.Worksheet)Workbook.Sheets["DESCRIPTION_TACHE"];
            Excel.Range selectedRowRange = (Excel.Range)xlWorksheet.Rows[selectedRowIndex + 2];

            // Get the address of the new row
            string newRowAddress = $"B{selectedRowIndex + 3}";

            // Insert a new row below the selected row
            selectedRowRange.Offset[1].Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            // Get the inserted row
            Excel.Range insertedRow = (Excel.Range)xlWorksheet.Rows[selectedRowIndex + 3];

            // Set the interior color of the inserted row
            insertedRow.Interior.ColorIndex = 0;

            // Set the value of the new row in column 2 (B)
            ((Excel.Range)xlWorksheet.Cells[selectedRowIndex + 3, 2]).Value = "";

            // Set the value of the new row in column 3 (C)
            ((Excel.Range)xlWorksheet.Cells[selectedRowIndex + 3, 3]).Value = "";

            // Save the changes
            Workbook.Save();

            // Clear the comboboxes
            ClearComboBoxes();

            // Update the combo boxes with the new data
            int dataStartRow = 2;

            // Populate combo boxes with cell values and row IDs
            WireUpComboBoxes();
            foreach (ComboBox comboBox in comboBoxes)
            {
                int columnIndex = comboBoxes.IndexOf(comboBox) + 1;
                AddCellValuesToComboBox(comboBox, worksheet, columnIndex, dataStartRow);

                // Wire up the event handler
                comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
            }
            PictoComboBox.SelectedIndex = selectedRowIndex + 1;
        }

        private void OutillageChoicesCombo_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}