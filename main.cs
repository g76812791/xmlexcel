using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace xmlexcel
{
    class Program
    {
        static void Main(string[] args)
        {

            //使用appdomain获取当前应用程序集的执行目录
            string dir = AppDomain.CurrentDomain.BaseDirectory+ "Viewers Report 2019x.xml";
            //使用path获取当前应用程序集的执行的上级目录
            //  dir = Path.GetFullPath("..");
            //使用path获取当前应用程序集的执行目录的上级的上级目录
            //  dir = Path.GetFullPath("../..");
            //相互转化方式 https://www.cnblogs.com/Yesi/p/6473741.html
            Workbook workbook1 = new Workbook();
            workbook1.LoadFromXml(dir);
            workbook1.SaveToFile("test1.xlsx", ExcelVersion.Version2013);

            byte[] buf = GetData(dir);//我得到的比特数组

            FileStream fs = new FileStream(@"D:\123.xls", FileMode.Create, FileAccess.Write);
            fs.Write(buf, 0, buf.Length);
            fs.Flush();
            fs.Close();
        }

        static byte[] GetData(string path)
        {
            //实例化内存流
            MemoryStream fileStream = new MemoryStream();
            //写文件流
            StreamWriter fileWriter = new StreamWriter(fileStream);
            //读文件流
            StreamReader fileReader = new StreamReader(path, System.Text.Encoding.UTF8);
            //读取整个文件的内容
            string tempStr = fileReader.ReadToEnd();
          
            fileWriter.WriteLine(tempStr);
            fileWriter.Flush();

            byte[] bytes = fileStream.ToArray();
            fileStream.Close();
            fileWriter.Close();
            fileReader.Close();
            return bytes;
        }
        //360图书馆
        //public byte[] GetData(string path, string tempName)
        //{
        //    //实例化内存流
        //    MemoryStream fileStream = new MemoryStream();
        //    //写文件流
        //    StreamWriter fileWriter = new StreamWriter(fileStream);
        //    //读文件流
        //    StreamReader fileReader = new StreamReader(path, System.Text.Encoding.UTF8);
        //    //读取整个文件的内容
        //    string tempStr = fileReader.ReadToEnd();
        //    //得到要导出的信息
        //    DataTable dt = airport.GetDataView();
        //    //定义变量获取表的数据行数
        //    int rows = 0;
        //    //定义StringBuilder变量来存储字符串
        //    StringBuilder sb = new StringBuilder();
        //    //判断表是否有数据
        //    if (dt != null && dt.Rows.Count > 0)
        //    {
        //        //给变量赋值
        //        rows = dt.Rows.Count;
        //        int num = rows + 1;//此处1表示模板中表头的行数，且默认为1，可自行增加表头并定义
        //        if (tempStr.IndexOf("+#RowCount#+") > 0)
        //        {
        //            tempStr = tempStr.Replace("+#RowCount#+", "" + num + "");
        //        }
        //        //开始循环生成标准字符串
        //        for (int i = 0; i < rows; i++)
        //        {
        //            sb.Append("<Row>\n"
        //                + "<Cell><Data ss:Type=\"String\">" + dt.Rows[i]["Name"] + "</Data></Cell>\n"
        //                + "<Cell><Data ss:Type=\"String\">" + dt.Rows[i]["Sex"] + "</Data></Cell>\n"
        //                + "<Cell><Data ss:Type=\"String\">" + dt.Rows[i]["Age"] + "</Data></Cell>\n"
        //                + "</Row>\n");
        //        }
        //        if (tempStr.IndexOf("+%Data%+") > 0)
        //        {
        //            tempStr = tempStr.Replace("+%Data%+", sb.ToString());
        //        }
        //    }
        //    fileWriter.WriteLine(tempStr);
        //    fileWriter.Flush();

        //    byte[] bytes = fileStream.ToArray();

        //    fileStream.Close();
        //    fileWriter.Close();
        //    fileReader.Close();

        //    return bytes;
        //}


        //https://www.cnblogs.com/waitingfor/archive/2011/12/19/2293469.html com 的引用
        //public static void ConvertExcel(string savePath)
        //{
        //    //将xml文件转换为标准的Excel格式 
        //    Object Nothing = Missing.Value;//由于yongCOM组件很多值需要用Missing.Value代替   
        //    Microsoft.Office.Interop.Excel.Application ExclApp = new Microsoft.Office.Interop.Excel.ApplicationClass();// 初始化
        //    Microsoft.Office.Interop.Excel.Workbook ExclDoc = ExclApp.Workbooks.Open(savePath, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing);//打开Excl工作薄   
        //    try
        //    {
        //        Object format = Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal;//获取Excl 2007文件格式   
        //        ExclApp.DisplayAlerts = false;
        //        ExclDoc.SaveAs(savePath, format, Nothing, Nothing, Nothing, Nothing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Nothing, Nothing, Nothing, Nothing, Nothing);//保存为Excl 2007格式   
        //    }
        //    catch (Exception ex) { }
        //    ExclDoc.Close(Nothing, Nothing, Nothing);
        //    ExclApp.Quit();
        //}

    }
}
