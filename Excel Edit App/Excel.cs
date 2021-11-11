using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace Excel_Edit_App
{
    class Excel
    {
        string path;
        int col;
        int idCol;
        int nameCol;
        int sheet;

        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel(string path, int sheet, int col, int idCol, int nameCol)
        {
            this.path = path;
            this.col = col;
            this.idCol = idCol;
            this.sheet = sheet;
            this.nameCol = nameCol;
        }

        private void OpenFile()
        {
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public Dictionary<int, string> ReadFile()
        {
            OpenFile();
            var data = new Dictionary<int, string>();
            string excelId, name;
            for (int i = 1; i<= ws.Rows.Count; i++)
            {
                if (i > 600) break;

                excelId = ws.Cells[i, idCol].Text;
                excelId = Regex.Replace(excelId, @"\s+", "");
                name = ws.Cells[i, nameCol].Text;

                if (int.TryParse(excelId, out int intId))
                {
                    data.Add(intId, name);
                }
            }
            excel.Quit();
            return data;
        }

        public Dictionary<int, float> EditCell(Dictionary<int,float> inputData)
        {
            OpenFile();
            string excelId;
            for (int i = 1; i <= ws.Rows.Count; i++)
            {

                if (i > 600 || inputData.Count == 0)  break; 

                excelId = ws.Cells[i, idCol].Text;
                excelId = Regex.Replace(excelId, @"\s+", "");

                if (int.TryParse(excelId, out int intId))
                {
                    if (inputData.ContainsKey(intId))
                    {
                        ws.Cells[i, col].Value2 = inputData[intId];
                        inputData.Remove(intId);
                    }
                    else if ( excelId.Length > 4 &&
                            inputData.ContainsKey(int.Parse(excelId.Substring(excelId.Length - 4))))
                    {
                        ws.Cells[i, col].Value2 = inputData[int.Parse(excelId.Substring(excelId.Length - 4))];
                        inputData.Remove(int.Parse(excelId.Substring(excelId.Length - 4)));
                    }
                }
            }
            return inputData;
        }

        public void Save()
        {
            wb.Save();
            excel.Quit();
        }

    }
}
