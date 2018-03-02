using DBUtilityTool.Classes;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DBUtilityTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataTable dt1 = GetFirsTable();
            DataTable dt2 = Get2ndTable();

            foreach (DataRow dr in dt2.Rows)
            {
                DataRow row = dt1.NewRow();
                row[0] = dr[0].ToString();
                row[1] = dr[1].ToString();
                dt1.Rows.Add(row);

            }
            ConvertToExcel(dt1);
        }

        private DataTable Get2ndTable()
        {
            DataTable output = new DataTable();
            string connectionString = @"Server=10.32.0.213;userid=appadmin;password=Mdec@cyber1215;Database=mscapp;Convert Zero Datetime=True";
            MySqlConnection conn = null;
            conn = new MySqlConnection(connectionString);
            conn.Open();

            try
            {
                //StringBuilder sql = new StringBuilder();
                string sQuery = txtQuery2.Text.Trim();
                //sql.AppendLine(" SELECT * FROM ( SELECT c.name, f.company_id ");
                //sql.AppendLine(" FROM companies AS c ");
                //sql.AppendLine(" LEFT JOIN foreign_knowledge_workers AS f ON f.company_id = c.id ");
                //sql.AppendLine(" ORDER BY name ASC ");
                //sql.AppendLine(" ) AS tmp_table GROUP BY name ");
                MySqlDataAdapter returnVal = new MySqlDataAdapter(txtQuery2.Text.Trim(), conn);
                returnVal.Fill(output);

                //ConvertToExcel(output);

            }
            catch (Exception ex)
            {

                throw;
            }
            conn.Close();
            return output;
        }

        private DataTable GetFirsTable()
        {
            DataTable output = new DataTable();
            string connectionString = @"Server=10.32.0.201;userid=mscapp;password=!0arnyOztQPWcuK*Z6J&;Database=mscapp;Convert Zero Datetime=True";
            MySqlConnection conn = null;
            conn = new MySqlConnection(connectionString);
            conn.Open();

            try
            {
                //StringBuilder sql = new StringBuilder();
                string sQuery = txtQuery1.Text.Trim();
                //sql.AppendLine(" SELECT * FROM ( SELECT c.name, f.company_id ");
                //sql.AppendLine(" FROM companies AS c ");
                //sql.AppendLine(" LEFT JOIN foreign_knowledge_workers AS f ON f.company_id = c.id ");
                //sql.AppendLine(" ORDER BY name ASC ");
                //sql.AppendLine(" ) AS tmp_table GROUP BY name ");
                MySqlDataAdapter returnVal = new MySqlDataAdapter(txtQuery1.Text.Trim(), conn);
                returnVal.Fill(output);

                //ConvertToExcel(output);

            }
            catch (Exception ex)
            {

                throw;
            }
            conn.Close();
            return output;
        }

        private void ConvertToExcel(DataTable output)
        {
            string ExcelFileName = "";
            ExcelFileHelper.GenerateExcelFile(output, out ExcelFileName);
            System.Diagnostics.Process.Start(ExcelFileName);
        }

        private string getLatestExcelFile()
        {

            string FileName = "";
            string folderPath = ConfigurationSettings.AppSettings["ExcelLocation"].ToString();
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            if (System.IO.Directory.GetFiles(folderPath).Length != 0)
            {
                string ExcelFile = "";
                var directory = new DirectoryInfo(folderPath);
                var fileName = directory.GetFiles()
                .OrderByDescending(f => f.LastWriteTime)
                .First();
                if (fileName != null)
                    FileName = ConfigurationSettings.AppSettings["ExcelLocation"].ToString() + fileName.ToString();
            }
            else
            {
                FileName = ConfigurationSettings.AppSettings["ExcelLocation"].ToString() + "DBUtilityTool_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
            }

            return FileName;
        }
    }
}
