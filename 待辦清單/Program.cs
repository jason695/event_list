using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using HtmlAgilityPack;
using System.Net;
using Excel=Microsoft.Office.Interop.Excel;
using System.Data;
//using System.Text.RegularExpressions;
//using System.Reflection;

namespace cut_sinocloud
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Data.DataTable dt = CUT_WEB(@"C:\Users\113720\Desktop\待辦事項.htm");
            
            if (dt.Rows.Count>0) EXPORT_XLS(dt);
            
            Console.ReadLine();
        }

        /// <summary>
        /// 截取WEB畫面
        /// </summary>
        /// <returns></returns>
        static System.Data.DataTable CUT_WEB(String path)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            DataColumn column;
            DataRow row;

            try
            {
                WebClient url = new WebClient();
                MemoryStream ms = new MemoryStream(url.DownloadData(path));
                Console.WriteLine("讀取檔案:" + path);

                HtmlDocument doc = new HtmlDocument();
                doc.Load(ms, Encoding.Default);

                //Xpath
                var res = doc.DocumentNode.SelectSingleNode(@"/html/body/form/div[3]/div/div/div/div[1]/table");
                if (res != null)
                {
                    // 明細
                    var list = res.SelectNodes(@"tr");

                    var hh = list[0].SelectNodes(@"th");
                    //Console.Write(hh[3].InnerText.Replace("\r\n", "").Replace(" ", "").Trim() + "|");
                    //Console.Write(hh[4].InnerText.Replace("\r\n", "").Replace(" ", "").Trim() + "|");
                    //Console.Write(hh[5].InnerText.Replace("\r\n", "").Replace(" ", "").Trim() + "|");
                    //Console.Write(hh[6].InnerText.Replace("\r\n", "").Replace(" ", "").Trim() + "|");
                    //Console.WriteLine(hh[7].InnerText.Replace("\r\n", "").Replace(" ", "").Trim());

                    // Create column
                    column = new DataColumn();
                    column.DataType = Type.GetType("System.String");
                    column.ColumnName = hh[3].InnerText.Replace("\r\n", "").Replace(" ", "").Trim();
                    table.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = Type.GetType("System.String");
                    column.ColumnName = hh[4].InnerText.Replace("\r\n", "").Replace(" ", "").Trim();
                    table.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = Type.GetType("System.String");
                    column.ColumnName = hh[5].InnerText.Replace("\r\n", "").Replace(" ", "").Trim();
                    table.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = Type.GetType("System.String");
                    column.ColumnName = hh[6].InnerText.Replace("\r\n", "").Replace(" ", "").Trim();
                    table.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = Type.GetType("System.String");
                    column.ColumnName = hh[7].InnerText.Replace("\r\n", "").Replace(" ", "").Trim();
                    table.Columns.Add(column);

                    list.RemoveAt(0);

                    foreach (var item in list)
                    {
                        var dd = item.SelectNodes(@"td");
                        //Console.Write(dd[3].InnerText.Replace("\r\n", "").Replace(" ", "").Trim() + "|");
                        //Console.Write(dd[4].InnerText.Replace("\r\n", "").Replace(" ", "").Trim() + "|");
                        //Console.Write(dd[5].InnerText.Replace("\r\n", "").Replace(" ", "").Trim() + "|");
                        //Console.Write(dd[6].InnerText.Replace("\r\n", "").Replace(" ", "").Trim() + "|");
                        //Console.WriteLine(dd[7].InnerText.Replace("\r\n", "").Replace(" ", "").Trim());

                        row = table.NewRow();
                        row[0] = dd[3].InnerText.Replace("\r\n", "").Replace(" ", "").Trim();
                        row[1] = dd[4].InnerText.Replace("\r\n", "").Replace(" ", "").Trim();
                        row[2] = dd[5].InnerText.Replace("\r\n", "").Replace(" ", "").Trim();
                        row[3] = dd[6].InnerText.Replace("\r\n", "").Replace(" ", "").Trim();
                        row[4] = dd[7].InnerText.Replace("\r\n", "").Replace(" ", "").Trim();
                        table.Rows.Add(row);
                    }

                    Console.WriteLine("截取畫面OK!!");
                }
                else
                {
                    Console.WriteLine("[ALERT]無資料可讀取");
                }
            }
            catch (Exception ex)
            {
                //throw ex;                
                Console.WriteLine("[ERROR]"+ex.Message);
                Console.Read();
            }
            
            return table;
        }

        /// <summary>
        /// 產生EXCEL
        /// </summary>
        /// <param name="dt"></param>
        static void EXPORT_XLS(System.Data.DataTable dt)
        {
            FileInfo fi = new FileInfo("Data.xlsx");
            Excel.Application xlapp = new Excel.Application();
            if (xlapp == null)
            {
                Console.WriteLine("[ALERT]請安裝office!!");
            }
            xlapp.Visible = false;//不顯示excel程式
            Excel.Workbook wb = xlapp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets[1];

            ws.Name = "需求單_" + DateTime.Today.ToString("yyyyMMdd");
            
            try
            {
                //表頭
                ws.Cells[1, 1] = dt.Columns[0].ColumnName;
                ws.Cells[1, 2] = dt.Columns[1].ColumnName;
                ws.Cells[1, 3] = dt.Columns[2].ColumnName;
                ws.Cells[1, 4] = dt.Columns[3].ColumnName;
                ws.Cells[1, 5] = dt.Columns[4].ColumnName;

                //明細
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow row = dt.Rows[i];

                    ws.Cells[i + 2, 1] = "'" + row[0];
                    ws.Cells[i + 2, 2] = row[1];
                    ws.Cells[i + 2, 3] = row[2];
                    ws.Cells[i + 2, 4] = row[3];
                    ws.Cells[i + 2, 5] = row[4];

                    //格式定義
                    Excel.Range ra;
                    ra = ((Excel.Range)ws.Cells[i + 2, 1]);
                    ra.ColumnWidth = 14;

                    ra = ((Excel.Range)ws.Cells[i + 2, 2]);
                    ra.ColumnWidth = 100;
                    ra.WrapText = true; // 自動換行

                    ra = ((Excel.Range)ws.Cells[i + 2, 3]);
                    ra.ColumnWidth = 10;

                    ra = ((Excel.Range)ws.Cells[i + 2, 4]);
                    ra.ColumnWidth = 20;

                }

                if (ws == null)
                {
                    Console.WriteLine("[ALERT]建立sheet失敗");
                }
                //wb.SaveAs(@fi.DirectoryName + "\\Data_" + DateTime.Today.ToString("yyyyMMdd") + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                string f = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                wb.SaveAs(f + @"/Data_" + DateTime.Today.ToString("yyyyMMdd") + ".xlsx"
                    , Excel.XlFileFormat.xlOpenXMLWorkbook
                    , Type.Missing
                    , Type.Missing
                    , Type.Missing
                    , Type.Missing
                    , Excel.XlSaveAsAccessMode.xlNoChange
                    , Type.Missing
                    , Type.Missing
                    , Type.Missing
                    , Type.Missing
                    , Type.Missing);
                wb.Close(false, Type.Missing, Type.Missing);

                Console.WriteLine("產檔:" + f + @"/需求單_" + DateTime.Today.ToString("yyyyMMdd") + ".xlsx");

                xlapp.Workbooks.Close();
                xlapp.Quit();                                
            }
            catch (Exception ex)
            {
                //throw ex;
                Console.WriteLine("[ERROR]" + ex.Message);
                Console.Read();
            }
            finally
            {
                //刪除 Windows工作管理員中的Excel.exe process，  
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);  
            }
        }

    }    
}
