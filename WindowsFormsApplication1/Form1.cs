using Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
        public Form1()
        {
            InitializeComponent();
            tbFileUrl.Text = @"C:\Users\JSR\Desktop\인터페이스_1\009. 전력량_MODBUS(node91~98)\MODBUS";
            tbNet.Text = "1.1";
            tbCul.Text = "기동시간적산,가스량계,E_Cooling,E_Heating,KWH,적산,누적";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var filePath = tbFileUrl.Text;
            var outputFileName = tbFileUrl.Text.Contains("MODBUS") ? "MODBUS" : "BACNET";

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
                            resultData = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16}", pointAddr, pointName.Replace(",",""),string.Empty, string.Empty, "Real", "사용중", "사용중",string.Empty,string.Empty,string.Empty, string.Empty, string.Empty, "Analog", "R", CheckCul(pointName), string.Empty, pointName.Replace(",", ""));
                            Log(resultData, outputFileName);
                        }
                    }
                }

                excelReader.Close();
            }

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

            string FilePath = @"C:\Users\JSR\Desktop\" + @"\PointLists\" + fileName + DateTime.Now.ToString("yyyyMMddHH") + ".csv";
            string DirPath = @"C:\Users\JSR\Desktop\" + @"\PointLists";
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
    }
}
