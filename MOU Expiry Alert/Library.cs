using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MOU_Expiry_Alert
{
    public static class Library
    {
        public static void WriteErrorLog(Exception ex)
        {
            StreamWriter sw = null;
            try
            {
                //string AppLocation = "";
                //AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
                // AppLocation = AppLocation.Replace("file:\\", "");@"C:\MOUAlertsWindowsService\ExcelFiles\MOUList.xls"
                //sw = new StreamWriter(AppLocation + "\\LogFiles\\LogFile"+DateTime.Now+".txt", true);
                sw = new StreamWriter(@"C:\MOUAlertsWindowsService\LogFiles\LogFile_" + DateTime.Now + ".txt", true);
                sw.WriteLine(DateTime.Now.ToString() + ": " + ex.Source.ToString().Trim() + "; " + ex.Message.ToString().Trim());
                sw.Flush();
                sw.Close();
            }
            catch
            {
            }
        }

        public static void WriteErrorLog(string Message)
        {
            StreamWriter sw = null;
            try
            {
                 sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);
               // sw = new StreamWriter(@"C:\MOUAlertsWindowsService\LogFiles\LogFile_" + DateTime.Now + ".txt", true);
                sw.WriteLine(DateTime.Now.ToString() + ": " + Message);
                sw.Flush();
                sw.Close();
            }
            catch
            {
            }
        }
    }
}
