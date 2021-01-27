using MySql.Data.MySqlClient;
using OfficeOpenXml;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
//using System.Threading;
using System.Timers;
using System.Threading.Tasks;

namespace RoboReestrService
{
    public partial class RoboReestrService : ServiceBase
    {
        public RoboReestrService()
        {
            InitializeComponent();

        }

        TimerStruct timer;

        protected override void OnStart(string[] args)
        {
            Logger.InitLogger();
            Logger.Log.Info("Старт");

            string strPath = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);

            try
            {
                INIManager manager = new INIManager(strPath + @"\settings.ini");
                timer = new TimerStruct(manager.GetPrivateString("Timer", "interval"), manager.GetPrivateString("Timer", "format"),
                manager.GetPrivateString("Timer", "condition"));

                /*// устанавливаем метод обратного вызова
                TimerCallback tm = new TimerCallback(CheckDate);
                // создаем таймер
                Logger.Log.Info(t.interval.ToString());
                Timer T2 = new Timer(tm, t, 0, t.interval);
                */

                //MainFunc();

                Timer T2 = new Timer();
                T2.Interval = timer.interval;
                T2.AutoReset = true;
                T2.Enabled = true;
                T2.Start();
                T2.Elapsed += new ElapsedEventHandler(T2_Elapsed);
            }
            catch (Exception ex)
            {
                Logger.Log.Error(ex.ToString());
            }           

        }

        public void CheckDate(object obj) 
        {

            TimerStruct timer = (TimerStruct)obj;
            try
            {

                string date = DateTime.Now.ToString(timer.format);
                if (date == timer.condition)
                {
                    MainFunc();
                }
            }
            catch (Exception ex)
            {
                Logger.Log.Error(ex.ToString());
            }

        }

        private void T2_Elapsed(object sender, EventArgs e)
        {
            try
            {
                
                string date = DateTime.Now.ToString(timer.format);
                if (date == timer.condition)
                {
                    MainFunc();
                }
            }
            catch (Exception ex)
            {
                Logger.Log.Error(ex.ToString());
            }
            //T2.Interval = timer.interval;
            //T2.Interval = 60000;
        }

        private void MainFunc() 
        {
            string strPath = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            INIManager manager = new INIManager(strPath + @"\settings.ini");

            SQLConnect conn = new SQLConnect(manager.GetPrivateString("Oracle", "host"), manager.GetPrivateString("Oracle", "port"), 
                manager.GetPrivateString("Oracle", "database"), manager.GetPrivateString("Oracle", "user"), manager.GetPrivateString("Oracle", "password"), 
                manager.GetPrivateString("MySQL", "host"), manager.GetPrivateString("MySQL", "port"),
                manager.GetPrivateString("MySQL", "database"), manager.GetPrivateString("MySQL", "user"), manager.GetPrivateString("MySQL", "password"), 
                manager.GetPrivateString("Request", "senderid"), manager.GetPrivateString("Request", "aggregator"));
            
            Mail mail = new Mail(manager.GetPrivateString("Email", "SMTPServer"), manager.GetPrivateString("Email", "from"), 
                manager.GetPrivateString("Email", "to"), manager.GetPrivateString("Email", "port"), manager.GetPrivateString("Email", "login"), 
                manager.GetPrivateString("Email", "password"));
            
            try
            {
                string[] fileNameArr = new string[2];
                fileNameArr[0] = GetRequestOracle(conn.oracleConnection, conn.senderId, strPath);
                fileNameArr[1] = GetRequestMySQL(conn.mysqlConnection, conn.aggregator, strPath);
                mail.SendMail(fileNameArr[0], fileNameArr[1], strPath);
            }
            catch (Exception ex)
            {
                Logger.Log.Error(ex.ToString());
            }
        }

        private string GetRequestOracle(OracleConnection oracleConn, string senderId, string strPath)
        {
            List<object[]> cellData = new List<object[]>();

            object[] objHeader = new object[12];
            objHeader[0] = "T_OBJECTID";
            objHeader[1] = "T_MESSAGEID";
            objHeader[2] = "T_SYSDATE";
            objHeader[3] = "T_LASTNAME";
            objHeader[4] = "T_FIRSTNAME";
            objHeader[5] = "T_MIDDLENAME";
            objHeader[6] = "T_STATE";
            objHeader[7] = "T_STATUS";
            objHeader[8] = "T_SENDERID";
            objHeader[9] = "T_SENDERPOINTID";
            objHeader[10] = "T_ERRORTEXT";
            objHeader[11] = "T_SMEVERRORTEXT";
            
            cellData.Add(objHeader);
            
            oracleConn.Open();

            string date = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");

            string oracleQuery = $@"SELECT T_OBJECTID ,
                                T_MESSAGEID ,
                                TO_CHAR(T_SYSDATE, 'dd.mm.yyyy HH24:MI:SS') as Date_of_request,
                                T_LASTNAME ,
                                T_FIRSTNAME ,
                                T_MIDDLENAME ,
                                T_STATE ,
                                T_STATUS ,
                                T_SENDERID ,
                                T_SENDERPOINTID ,
                                T_ERRORTEXT ,
                                T_SMEVERRORTEXT FROM duprid_dbt
                                WHERE T_SENDERID = '{senderId}'
                                and T_SYSDATE >= to_date('{date} 00:00:00', 'yyyy-mm-dd HH24:MI:SS') and T_SYSDATE <= to_date('{date} 23:59:59', 'yyyy-mm-dd HH24:MI:SS')
                                order by t_sysdate";

            OracleCommand oracleCommand = new OracleCommand(oracleQuery);

            oracleCommand.Connection = oracleConn;
            oracleCommand.CommandType = CommandType.Text;

            OracleDataReader reader = oracleCommand.ExecuteReader();

            while (reader.Read())
            {
                object[] obj = new object[12];
                for (int i = 0; i < 12; i++)
                {
                    obj[i] = reader.GetValue(i).ToString();
                    if (i == 2)
                        date = obj[i].ToString();
                }
                cellData.Add(obj);
            }

            reader.Dispose();
            oracleCommand.Dispose();
            oracleConn.Dispose();

            
            using (ExcelPackage excel = new ExcelPackage())
            {
                string[] dateArr = date.Split(' ');
                date = dateArr[0];

                excel.Workbook.Worksheets.Add(senderId + " " + date);
                ExcelWorksheet worksheet = excel.Workbook.Worksheets[senderId + " " + date];
                worksheet.Cells[1,1].LoadFromArrays(cellData);
                worksheet.Cells.AutoFitColumns();

                DirectoryInfo dirInfo = new DirectoryInfo(strPath + @"\Reestrs");
                if (!dirInfo.Exists)
                {
                    dirInfo.Create();
                }

                FileInfo excelFile = new FileInfo(strPath + @"\Reestrs\Reestr UPRID " + senderId + " " + date + ".xlsx");
                excel.SaveAs(excelFile);
            }
            return(senderId + " " + date);
        }

        private string GetRequestMySQL(MySqlConnection mysqlConn, string aggregator, string strPath) 
        {
            List<object[]> cellData = new List<object[]>();
            
            object[] objHeader = new object[12];
            objHeader[0] = "id";
            objHeader[1] = "aggregator";
            objHeader[2] = "type";
            objHeader[3] = "registered_datetime";
            objHeader[4] = "status";
            objHeader[5] = "aggregator_payee_id";
            objHeader[6] = "name";
            objHeader[7] = "email";
            objHeader[8] = "phone";
           
            cellData.Add(objHeader);

            mysqlConn.Open();

            string date = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
            //string date = "2020-08-03";

            string mysqlQuery = $@"select id, aggregator, type, DATE_FORMAT(registered_datetime, '%d.%m.%Y %T') as registered_datetime, status, aggregator_payee_id, name, email, phone, first_name, last_name, third_name from payees_all_data
                where aggregator='{aggregator}'
                and type='ESP'
                and registered_datetime >= '{date} 00:00:00' and registered_datetime <= '{date} 23:59:59';";

            MySqlCommand mysqlCommand = new MySqlCommand(mysqlQuery);

            mysqlCommand.Connection = mysqlConn;
            mysqlCommand.CommandType = CommandType.Text;

            MySqlDataReader reader = mysqlCommand.ExecuteReader();

            while (reader.Read())
            {
                object[] obj = new object[9];
                for (int i = 0; i < 9; i++)
                {
                    
                    obj[i] = reader.GetValue(i).ToString();
                    if (i == 3)
                        date = obj[i].ToString();

                }
                cellData.Add(obj);
            }

            reader.Dispose();
            mysqlCommand.Dispose();
            mysqlConn.Dispose();

            using (ExcelPackage excel = new ExcelPackage())
            {
                string[] dateArr = date.Split(' ');
                date = dateArr[0];

                excel.Workbook.Worksheets.Add(aggregator + " " + date);
                ExcelWorksheet worksheet = excel.Workbook.Worksheets[aggregator + " " + date];
                worksheet.Cells[1,1].LoadFromArrays(cellData);
                worksheet.Cells.AutoFitColumns();

                DirectoryInfo dirInfo = new DirectoryInfo(strPath + @"\Reestrs");
                if (!dirInfo.Exists)
                {
                    dirInfo.Create();
                }

                FileInfo excelFile = new FileInfo(strPath + @"\Reestrs\Reestr Wallet " + aggregator + " " + date + ".xlsx");
                excel.SaveAs(excelFile);
            }
            return(aggregator + " " + date);
        }

        protected override void OnStop()
        {
            Logger.Log.Info("Стоп");
        }
    }
}
