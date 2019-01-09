using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;

namespace 工作日誌
{
    class Program
    {
        static void Main(string[] args)
        {
        //(1)讀取C:\Users\113720\Desktop\工作日誌\下 最新的txt
        //COPY一份,更名為今天日期
            
            String dir = @"C:\Users\113720\Desktop\工作日誌";
            String fil_copy = "";
            String fil_new = "W" + DateTime.Now.ToString("yyyyMMdd") +".txt";
            
            DirectoryInfo dirinfo = new DirectoryInfo(dir);
            FileInfo[] sortList = dirinfo.GetFiles();
            Array.Sort(sortList, new MyDateSorter());
                        
            foreach (FileInfo item in sortList)
            {
                fil_copy = item.Name;
                Console.WriteLine("來源來:" + fil_copy);
                break;
            }
            Console.WriteLine("目的檔:" + fil_new);

            //File.Copy(dir + @"\" + fil_copy ,dir + @"\" + fil_new ,true);
            if (fil_copy == fil_new)
            {
                Console.WriteLine("目的檔已存在,確認是否刪除(Y/N):" + fil_new);

                string v1 = Console.ReadLine();

                if (v1 == "Y")
                {
                    File.Delete(dir + @"\" + fil_new);
                    Console.WriteLine("目的檔已存在,刪除:" + fil_new);
                    Main(args);
                    return;
                }
                else
                {
                    return;
                }                
            }

        //(2)將前7行改為
        //========================================================================
        //==TODAY=================================================================
        //========================================================================
        //
        //========================================================================
        //==前 次=================================================================
        //========================================================================
            using (StreamReader sr = File.OpenText(dir + @"\" + fil_copy))
            {
                using (StreamWriter sw = File.AppendText(dir + @"\" + fil_new))
                {
                    string s = "";
                    int i = 1; //行數
                    while ((s = sr.ReadLine()) != null)
                    {
                        if (i == 4)
                        {
                            sw.WriteLine(Environment.NewLine 
                                + "========================================================================" + Environment.NewLine
                                + "==" + fil_copy.Substring(5, 2) + "/" + fil_copy.Substring(7, 2) + "=================================================================" + Environment.NewLine
                                + "========================================================================" + Environment.NewLine
                                + s);
                        }
                        else {
                            sw.WriteLine(s);
                        }

                        i += 1;
                    }
                }
            }
            Console.WriteLine("OK!!");
            Console.Read();
        }
    }

    public class MyDateSorter : IComparer
    {
        #region IComparer Members
        public int Compare(object x, object y)
        {
            if (x == null && y == null)
            {
                return 0;
            }
            if (x == null)
            {
                return -1;
            }
            if (y == null)
            {
                return 1;
            }
            FileInfo xInfo = (FileInfo)x;
            FileInfo yInfo = (FileInfo)y;


            //依名稱排序
            //return xInfo.FullName.CompareTo(yInfo.FullName);//遞增
            return yInfo.FullName.CompareTo(xInfo.FullName);//遞減

            //依修改日期排序
            //return xInfo.LastWriteTime.CompareTo(yInfo.LastWriteTime);//遞增
            //return yInfo.LastWriteTime.CompareTo(xInfo.LastWriteTime);//遞減
        }
        #endregion
    }
}
