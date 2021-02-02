using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace Excel_Edit_App
{
    class Excel
    {
        string path;
        int col;
        List<int> notSavedIds = new List<int>();

        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel(string path, int sheet, int col)
        {
            this.path = path;
            this.col = col;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public void EditCell(int id, float mark)
        {
            string excelId;
            int number;
            for (int i = 1; i <= ws.Rows.Count; i++)
            {
                if (i > 600)
                {
                    notSavedIds.Add(id);
                    break;
                }
                excelId = ws.Cells[i, 2].Text;

                if (int.TryParse(excelId, out number) && checkId(excelId, id))
                {
                    ws.Cells[i, col].Value2 = mark;
                    break;
                }
            }
        }

        public bool checkId(string excelId, int id)
        {
            if (excelId.Length == 4)
            {
                return int.Parse(excelId) == id;
            }
            else
            {
                return int.Parse(excelId.Substring(excelId.Length - 4)) == id;
            }

        }

        public void Save()
        {
            wb.Save();
            excel.Quit();
        }

        public List<int> check()
        {
            return notSavedIds;
        }

    }
}
