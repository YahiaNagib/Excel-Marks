using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel_Edit_App
{
    public partial class Form1 : Form
    {
        string fileLocation;
        Excel excel;
        Dictionary<int, string> data;
        public Form1()
        {
            InitializeComponent();
            workSheetNumberTextEdit.Text = "1";
            IdColNumberTextEdit.Text = "2";
        }

        // Opening and reading the file
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            int col, sheet, idCol, nameCol;
            try
            {
                col = int.Parse(columnNumberTextEdit.Text);
                sheet = int.Parse(workSheetNumberTextEdit.Text);
                idCol = int.Parse(IdColNumberTextEdit.Text);
                nameCol = int.Parse(NameColumnNumberTextEdit.Text);
            }
            catch
            {
                MessageBox.Show("Please enter the missing data");
                return;
            }
            openFileDialog1.Title = "Choose file";
            openFileDialog1.ShowDialog();
            fileLocation = openFileDialog1.FileName;
            label2.Text = fileLocation + " Loaded";
            label1.Text = "Opening File, Please Wait";
            excel = new Excel(fileLocation, sheet, col, idCol, nameCol);
            data = excel.ReadFile();
            label1.Text = "File opened successfully";
        }

        // Save button
        private void simpleButton2_Click(object sender, EventArgs e)
        {
            label1.Text = "Saving, Please Wait";
            var inputData = new Dictionary<int, float>();
            int id; float mark;
            for (int i = 0; i < gridView1.DataRowCount; i++)
            {
                id = int.Parse(gridView1.GetRowCellValue(i, "id").ToString());
                mark = float.Parse(gridView1.GetRowCellValue(i, "mark").ToString());
                inputData.Add(id, mark);
            }
            // Applying edits
            var notSavedIds = excel.EditCell(inputData);

            // Saving the file
            excel.Save();

            if (notSavedIds.Count == 0)
            {
                label1.Text = "Done!";
            }
            else
            {
                string ids = "";
                foreach (var item in notSavedIds)
                {
                    ids = ids + item.Key + " ";
                }
                MessageBox.Show("All is saved except " + ids);
            }
            // markBindingSource.Clear();

        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            markBindingSource.Clear();
        }

        private void gridView1_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }

        private void gridView1_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            int id;

            if (excel == null)
            {
                MessageBox.Show("Please, choose a file", "File Not Specified", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (e.Column.FieldName == "id")
            {
                id = int.Parse(gridView1.GetRowCellValue(e.RowHandle, "id").ToString());
                try
                {
                    gridView1.SetRowCellValue(e.RowHandle, "name", data[id]);
                }
                catch
                {
                    gridView1.SetRowCellValue(e.RowHandle, "name", "Not Found!");
                }
            }
        }
    }
}
