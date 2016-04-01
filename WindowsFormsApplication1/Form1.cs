using Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        Queue<string> que = new Queue<string>();
        Queue<string> queSql = new Queue<string>();
        private BackgroundWorker worker;
        private BackgroundWorker worker2;

        public Form1()
        {
            InitializeComponent();
            tbFileUrl.Text = @"C:\Users\JSR\Desktop\인터페이스_1\009. 전력량_MODBUS(node91~98)\MODBUS";
            tbNet.Text = "1.1";
            tbCul.Text = "기동시간적산,가스량계,E_Cooling,E_Heating,KWH,적산,누적";
            tbSql.Text = "Data Source=localhost;Initial Catalog=" + "OEMSMST" + ";USER ID=" + "sa" + ";PASSWORD=" + "pass@word!02" + ";Connection Timeout=3600";

            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;
            worker.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);

            worker2 = new BackgroundWorker();
            worker2.WorkerReportsProgress = true;
            worker2.WorkerSupportsCancellation = true;
            worker2.DoWork += new DoWorkEventHandler(backgroundWorker2_DoWork);
            worker2.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker2_RunWorkerCompleted);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            worker.RunWorkerAsync();
            Run();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            worker2.RunWorkerAsync();
        }

        private void Run()
        {
            var filePath = tbFileUrl.Text;

            if (tbNet.Text == string.Empty)
            {
                MessageBox.Show("앞구분자를 입력하세요");
                return;
            }

            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(filePath);

            foreach (System.IO.FileInfo f in di.GetFiles())
            {
                var subFIleFullName = f.FullName;
                FileStream stream = File.Open(subFIleFullName, FileMode.Open, FileAccess.Read);

                IExcelDataReader excelReader = null;
                if (f.Extension.Equals(".xlsx"))
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                else 
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

                excelReader.IsFirstRowAsColumnNames = true;
                DataSet result = excelReader.AsDataSet();

                if (result.Tables == null || result.Tables.Count <= 0) return;

                var resultData = string.Empty;
                for (int i = 0; i < result.Tables.Count; i++)
                {
                    if (result.Tables[i].TableName.Equals("ANALOG") || result.Tables[i].TableName.Equals("BINARY"))
                    {
                        //OBJECT_NAME	DEVICE_INST	OBJECT_TYPE	OBJECT_INST
                        for (int j = 0; j < result.Tables[i].Rows.Count; j++)
                        {
                            var pointCul = "순시값"; 
                            var pointName = result.Tables[i].Rows[j]["OBJECT_NAME"].ToString();
                            var pointAddr = string.Format("{0}.{1}.{2}-{3}", tbNet.Text, result.Tables[i].Rows[j]["DEVICE_INST"].ToString(), result.Tables[i].Rows[j]["OBJECT_TYPE"].ToString(), result.Tables[i].Rows[j]["OBJECT_INST"].ToString());
                            var cul = CheckCul(pointName);
                            resultData = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16}", pointAddr, pointName.Replace(",",""),string.Empty, string.Empty, "Real", "사용중", "사용중",string.Empty,string.Empty,string.Empty, string.Empty, string.Empty, "Analog", "R", cul, string.Empty, pointName.Replace(",", ""));
                            var tbPoint = "INSERT INTO tbPoint VALUES ( " + string.Format(" '{0}', N'{1}', {2}, '{3}', '{4}', {5}, {6}, {7}, GETDATE(), 'setup', GETDATE(), 'setup')", pointAddr, pointName.Replace("'", ""), 9999, string.Empty, "R", 1, 1, 9999);
                            var tbPointProp = "DECLARE @Seq int; select @Seq = PointSeq from tbPoint where pt_addr = '" + pointAddr + "'; INSERT INTO tbPointProp VALUES ( " + string.Format("@Seq, '{0}', '{1}', {2}, {3}, '{4}', '{5}', '{6}', N'{7}',{8})", string.Empty, string.Empty, 0, 0, 'A', string.Empty, cul == "누적값" ? "C" : "I", pointName.Replace("'", ""), -1);
                            que.Enqueue(resultData);
                            queSql.Enqueue(tbPoint);
                            queSql.Enqueue(tbPointProp);
                        }
                    }
                }

                excelReader.Close();
            }
            queSql.Enqueue("insert into FMS_Master.dbo.tbTotal select PointSeq as pt_index, pt_addr, pt_name, 0 as pt_value, 0 as LogTime from vwPoint where pt_addr not in (select pt_addr from FMS_Master.dbo.tbTotal)");
        }


        public string CheckCul(string _str)
        {
            string[] splitStr = tbCul.Text.Split(',');
            string rtnValue = "순시값";
            for (int i = 0; i < splitStr.Length; i++)
            {
                if (_str.ToUpper().Contains(splitStr[i].ToUpper()))
                {
                    rtnValue = "누적값";
                    break;
                }
            }

            return rtnValue;
        }


        public static string GetDateTime()
        {
            DateTime NowDate = DateTime.Now;
            return NowDate.ToString("yyyy-MM-dd HH:mm:ss") + ":" + NowDate.Millisecond.ToString("000");
        }

        public static void Log(string str, string fileName)
        {
            // 삭제 기준 Log 날짜
            var dt = DateTime.Today.AddDays(-7);

            string FilePath = @"C:\BEMS_PointList\" + fileName + DateTime.Now.ToString("yyyyMMddHH") + ".csv";
            string DirPath = @"C:\BEMS_PointList";
            string temp;

            DirectoryInfo di = new DirectoryInfo(DirPath);
            FileInfo fi = new FileInfo(FilePath);

            try
            {
                if (di.Exists != true)
                {
                    Directory.CreateDirectory(DirPath);
                }
                else
                {
                    FileInfo[] dfi = di.GetFiles();
                    foreach (var item in dfi)
                    {
                        if (item.CreationTime < dt)
                            item.Delete();
                    }
                }

                using (StreamWriter sw = new StreamWriter(FilePath, true, Encoding.UTF8))
                {
                    temp = string.Format("{0}", str).Replace("Log", "");
                    sw.WriteLine(temp);
                    sw.Close();
                }
            }
            catch (Exception e)
            {
               // Log(string.Format("{1} {0}", e.ToString(), GetDateTime()));
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (que.Count > 0)
            {
                for (int i = 0; i < que.Count; i++)
                {
                    //tbLog.Text += que.Dequeue() + "\r\n"; 
                    var outputFileName = tbFileUrl.Text.Contains("MODBUS") ? "MODBUS" : "BACNET";
                    Log(que.Dequeue(), "BEMS_" + outputFileName + "_");
                }
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("CSV 파일 생성 완료");
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            if (queSql.Count > 0)
            {
                for (int i = 0; i < queSql.Count; i++)
                {
                    SqlConnection sqlConn = new SqlConnection(tbSql.Text);
                    SqlCommand sqlComm = new SqlCommand();
                    sqlComm = sqlConn.CreateCommand();
                    sqlComm.CommandText = queSql.Dequeue();
                    sqlConn.Open();
                    sqlComm.ExecuteNonQuery();
                    sqlConn.Close();

                }
            }
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("DB에 넣기 완료");
        }
    }
}
