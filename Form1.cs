using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelUnion
{
    public partial class Form1 : Form
    {
        //private List<Column> Columns1 { get; set; } = new List<Column>();
        //private List<Column> Columns2 { get; set; } = new List<Column>();

        //private List<string> Sheets1;
        //private List<string> Sheets2;

        private ExcelLoader excelLoader1;
        private ExcelLoader excelLoader2;

        public Form1()
        {
            InitializeComponent();
        }

        private void file1Button_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            file1TextBox.Text = openFileDialog1.FileName;

        }
        private void file2Button_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            file2TextBox.Text = openFileDialog1.FileName;

        }

        private ExcelLoader UpdateColumnsAndSheets(TextBox fileTextBox, ComboBox sheetComboBox, ComboBox keyColumnComboBox, CheckedListBox chosenColsListBox)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (string.IsNullOrWhiteSpace(fileTextBox.Text))
            {
                return null;
            }
            ExcelLoader excelLoader = new ExcelLoader(fileTextBox.Text);
            List<string> Sheets = excelLoader.Sheets();
            sheetComboBox.Items.Clear();
            sheetComboBox.Items.AddRange(Sheets.ToArray());
            sheetComboBox.SelectedIndex = 0;           

            UpdateColumns(fileTextBox, sheetComboBox, keyColumnComboBox, excelLoader, chosenColsListBox);
            Cursor.Current = Cursors.Default;
            return excelLoader;
        }

        private void UpdateColumns(TextBox fileTextBox, ComboBox sheetComboBox, ComboBox keyColumnComboBox, ExcelLoader excelLoader, CheckedListBox chosenColsListBox)
        {
            string sheetname = sheetComboBox.Text;
            if (string.IsNullOrWhiteSpace(sheetname))
            {
                return;
            }

            List<Column> Columns = excelLoader.Columns(sheetname);
            keyColumnComboBox.Items.Clear();
            keyColumnComboBox.Items.AddRange(Columns.ToArray());
            keyColumnComboBox.SelectedIndex = 0;
            foreach(Column col in keyColumnComboBox.Items)
            {
                if(col.Title.Equals("Matricule", StringComparison.OrdinalIgnoreCase) || col.Title.Equals("Identifiant MAIF", StringComparison.OrdinalIgnoreCase))
                {
                    keyColumnComboBox.SelectedItem = col;
                }
            }
            chosenColsListBox.Items.Clear();
            chosenColsListBox.Items.AddRange(Columns.ToArray());
            HashSet<string> defaultCols = new HashSet<string>();
            defaultCols.Add("Matricule");
            defaultCols.Add("Nom");
            defaultCols.Add("Prénom");
            defaultCols.Add("Statut de l'affiliation");
            defaultCols.Add("Nom RRH");
            defaultCols.Add("Prénom RRH");
            defaultCols.Add("Nom du manager");
            defaultCols.Add("Prénom du manager");

            foreach (Column col in Columns)
            {
                if (defaultCols.Contains(col.Title, StringComparer.OrdinalIgnoreCase))
                {
                    chosenColsListBox.SetItemChecked(chosenColsListBox.Items.IndexOf(col), true);
                }
            }
        }


        private void buttonLaunch_Click(object sender, EventArgs e)
        {
            saveFileDialog.ShowDialog();
            string fileOut = saveFileDialog.FileName;
            
            Cursor.Current = Cursors.WaitCursor;
            int lines1 = excelLoader1.Lines(sheet1ComboBox.Text);
            int lines2 = excelLoader2.Lines(sheet2ComboBox.Text);
            progressBarA.Maximum = (lines1 + lines2 ) *2;
            progressBarA.Value = 0;
            //excelLoader1.Compare(sheet1ComboBox.Text, ((Column)keyColumn1ComboBox.SelectedItem).index, this.progressBarA);

            Content contents1 = this.excelLoader1.Content(sheet1ComboBox.Text, ((Column)keyColumn1ComboBox.SelectedItem).index, chosenCols1ListBox.CheckedItems.OfType<Column>().ToList(), this.progressBarA);
            Content contents2 = this.excelLoader2.Content(sheet2ComboBox.Text, ((Column)keyColumn2ComboBox.SelectedItem).index, chosenCols2ListBox.CheckedItems.OfType<Column>().ToList(), this.progressBarA);
            //int count = contents1.Keys.Intersect(contents2.Keys).Count();

            ExcelLoader.GenerateUnion(fileOut, contents1, contents2, this.progressBarA);

            Cursor.Current = Cursors.Default;
            MessageBox.Show("C'est fait ! " + fileOut);
        }

        private void file1TextBox_TextChanged(object sender, EventArgs e)
        {
            this.excelLoader1 = UpdateColumnsAndSheets(file1TextBox, sheet1ComboBox, keyColumn1ComboBox, chosenCols1ListBox);
        }

        private void file2TextBox_TextChanged(object sender, EventArgs e)
        {
            this.excelLoader2 = UpdateColumnsAndSheets(file2TextBox, sheet2ComboBox, keyColumn2ComboBox, chosenCols2ListBox);
        }

        private void sheet1ComboBox_TextChanged(object sender, EventArgs e)
        {
            if (this.excelLoader1 != null) UpdateColumns(file1TextBox, sheet1ComboBox, keyColumn1ComboBox, excelLoader1, chosenCols1ListBox);
        }

        private void sheet2ComboBox_TextChanged(object sender, EventArgs e)
        {
            if (this.excelLoader2 != null) UpdateColumns(file2TextBox, sheet2ComboBox, keyColumn2ComboBox, excelLoader2, chosenCols2ListBox);
        }

        private void outputDirButton_Click(object sender, EventArgs e)
        {

        }
    }
}
