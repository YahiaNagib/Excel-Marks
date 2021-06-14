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
        public Form1()
        {
            InitializeComponent();
            workSheetNumberTextEdit.Text = "1";
            IdColNumberTextEdit.Text = "2";
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Choose file";
            openFileDialog1.ShowDialog();
            fileLocation = openFileDialog1.FileName;
            label2.Text = fileLocation + " Loaded";
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            int col = int.Parse(columnNumberTextEdit.Text);
            int sheet = int.Parse(workSheetNumberTextEdit.Text);
            int idCol = int.Parse(IdColNumberTextEdit.Text);
            var excel = new Excel(fileLocation, sheet, col, idCol);
            int id; float mark;
            for (int i = 0; i < gridView1.DataRowCount; i++)
            {
                id = int.Parse(gridView1.GetRowCellValue(i, "id").ToString());
                mark = float.Parse(gridView1.GetRowCellValue(i, "mark").ToString());
                excel.EditCell(id, mark);
            }

            excel.Save();
            var notSavedIds = excel.check();
            if (notSavedIds.Count == 0)
            {
                MessageBox.Show("All is saved");
            }
            else
            {
                string ids = "";
                foreach (var item in notSavedIds)
                {
                    ids = ids + item.ToString() + " ";
                }
                MessageBox.Show("All is saved except " + ids);
            }
            markBindingSource.Clear();

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
    }
}
