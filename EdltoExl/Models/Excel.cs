//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Web;
//using System.Threading.Tasks;
//using Microsoft.Office.Interop.Excel;
//using _Excel = Microsoft.Office.Interop.Excel;

//namespace EdltoExl.Models
//{
//    public class Excel
//    {

//        string path = "";
//        _Application excel = new _Excel.Application();
//        Workbook wb;
//        Worksheet ws;
         

//        public Excel(string path, int Sheet)
//        {
//            this.path = path;
//            wb = excel.Workbooks.Open(path);

//            ws = wb.Worksheets[Sheet];
           
//        }
      


//        public void WriteToCell(int i, int j, string s)
//        {
//            // i++;
//            //j++;
//            ws.Cells[i, j].Value2 = s;

//        }

//        public string ReadCell(int i, int j)
//        {
//            i++;
//            j++;
//            if (ws.Cells[i, j].Value2 != null)
//                return ws.Cells[i, j].Value2;
//            else
//                return "";

//        }
//        public void Save()
//        {
//            wb.Save();

//        }

//        public void SaveAs(string path)
//        {
//            wb.SaveAs(path);

//        }

//        public void close()
//        {
//            wb.Close();
       

//        }
//    }
//}
