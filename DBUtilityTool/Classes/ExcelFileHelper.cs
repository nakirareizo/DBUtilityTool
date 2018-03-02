using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DBUtilityTool.Classes
{
    class ExcelFileHelper
    {
        internal static void GenerateExcelFile(DataTable dt, out string sFileName)
        {
            #region "EXPORT spBigFile INTO EXCEL"
            //Export spBigFile into EXCEL File
            string folderPath = ConfigurationSettings.AppSettings["ExcelLocation"].ToString();
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            sFileName = "";
            if (IsDirectoryEmpty(folderPath))
            {
                sFileName = ConfigurationSettings.AppSettings["ExcelLocation"].ToString() + "DBRecord_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
            }
            else
            {
                int LastCounter = getLastFileCounter(folderPath);

                if (LastCounter == 0)
                    sFileName = ConfigurationSettings.AppSettings["ExcelLocation"].ToString() + "DBRecord_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
                else
                    sFileName = ConfigurationSettings.AppSettings["ExcelLocation"].ToString() + "DBRecord_" + DateTime.Now.ToString("ddMMyyyy") + "_" + LastCounter.ToString() + ".xlsx";
            }
            ConvertToExcel(dt, sFileName);
            #endregion
        }

        internal static void GenerateExcelFileApprovalDates(DataTable dt, string SyncedDate)
        {
            #region "EXPORT spBigFile INTO EXCEL"
            //Export spBigFile into EXCEL File
            string folderPath = ConfigurationSettings.AppSettings["ExcelLocation"].ToString();
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            string ExcelFile = "";
            string SyncType = "ApprovalDates";
            if (IsDirectoryEmpty(folderPath))
            {
                ExcelFile = ConfigurationSettings.AppSettings["ExcelLocation"].ToString() + "DBRecord_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
            }
            else
            {

                int LastCounter = getLastFileCounter(folderPath);

                if (LastCounter == 0)
                    ExcelFile = ConfigurationSettings.AppSettings["ExcelLocation"].ToString() + "DBRecord_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
                else
                    ExcelFile = ConfigurationSettings.AppSettings["ExcelLocation"].ToString() + "DBRecord_" + DateTime.Now.ToString("ddMMyyyy") + "_" + LastCounter.ToString() + ".xlsx";
            }
            ConvertToExcel(dt, ExcelFile);
            #endregion
        }

        private static bool IsDirectoryEmpty(string path)
        {
            return !Directory.EnumerateFileSystemEntries(path).Any();
        }
        private static void ConvertToExcel(DataTable dt, string ExcelFile)
        {        
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "DBRecord_" + DateTime.Now.ToString("ddMMyyyy"));
                wb.SaveAs(ExcelFile);
            }
        }

        private static int getLastFileCounter(string folderPath)
        {
            int output = 0;
            var directory = new DirectoryInfo(folderPath);
            var fileName = directory.GetFiles()
            .OrderByDescending(f => f.LastWriteTime)
            .First();

            string[] arrFile = fileName.ToString().Split('_');
            string sDate = arrFile[1].ToString().Substring(0, 8); //21032016
            DateTime dDate = getDateTime(sDate);
            //1.xlsx, 10.xlsx
            string lastCounter = "";
            if (dDate.ToString("ddMMyyyy") == DateTime.Now.ToString("ddMMyyyy"))
            {
                if (arrFile.Count() > 2)
                {
                    if (arrFile[2].Length <= 6)
                        lastCounter = arrFile[2].Substring(0, 1);
                    else
                        lastCounter = arrFile[2].Substring(0, 2);
                }
                if (arrFile.Length == 2)
                    output = 1;
                else if (Convert.ToInt32(lastCounter) == 10)
                {
                    output = 0;
                }
                else
                {
                    output = Convert.ToInt32(lastCounter) + 1;
                }
            }
            else
                output = 0;
            return output;
        }
        private static DateTime getDateTime(string sDate)
        {
            DateTime myDate = new DateTime();
            string[] formats = { "ddMMyyyy" };
            return myDate = DateTime.ParseExact(sDate, formats, new CultureInfo(Thread.CurrentThread.CurrentCulture.Name), DateTimeStyles.None);
        }
    }
}
