using System;
using System.Diagnostics;
using System.ServiceProcess;
using System.Configuration;
using System.IO;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Net;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace CryptBotRequestDocNoService
{
    public partial class CryptBotRequestDocNoService : ServiceBase
    {
        string dbu, dbp, dbn, dbs, svr1, eofficeid, cbl, NetWorkCardID, Web_Server = "", ConnectionString = "";
        System.Timers.Timer timer = new System.Timers.Timer();
        System.Timers.Timer timerClock = new System.Timers.Timer();
        System.Diagnostics.EventLog eventLog1;
        DataTable DT = new DataTable();
        SqlConnection conn = new SqlConnection();
        bool ResultTime = true;
        List<string> lines = new List<string>();
        string EventName = ConfigurationManager.AppSettings["EventName"].ToString();
        string EventSource = ConfigurationManager.AppSettings["EventSource"].ToString();
        string sql = "";
        string path = "";
        string filename = "";
        string fullPath = "";
        bool ExistsFile;
        int countTimer = 0;

        public void SelectData(string str)  //for select to DT
        {
            conn = new SqlConnection(ConnectionString);
            conn.Open();
            SqlCommand queryCommand = new SqlCommand(str, conn);
            SqlDataReader queryCommandReader = queryCommand.ExecuteReader();
            DT.Load(queryCommandReader);
            conn.Close();
        }
        public void UpdateData(string str)
        {
            conn = new SqlConnection(ConnectionString);
            conn.Open();
            SqlCommand queryCommand = new SqlCommand(str, conn);
            queryCommand.ExecuteNonQuery();
            Console.WriteLine("Quesry STR : "+str);
            conn.Close();
        }
        internal void TestStartupAndStop(string[] args)
        {
            this.OnStart(args);
            Console.ReadLine();
            this.OnStop();
        }
        public CryptBotRequestDocNoService()
        {
            InitializeComponent();
            eventLog1 = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists(EventSource))
            {
                System.Diagnostics.EventLog.CreateEventSource(EventSource, EventName);
            }
            eventLog1.Source = EventSource;
            eventLog1.Log = EventName;
        }
        public void OnDebug()
        {
            this.OnStart(null);
        }
        protected override void OnStart(string[] args)
        {
            ConvertConfig();
            ConnectionString = "Persist Security Info=False;User ID=" + dbu + ";Password=" + dbp + ";Initial Catalog=" + dbn + ";Data Source=" + dbs;
            //ConnectionString = "Persist Security Info=False;User ID=eOfficeParliament;Password=P#w5NT89;Initial Catalog=eOfficeParliamentTest;Data Source=192.168.220.46";
            Web_Server = svr1 + "/";
            //eventLog1.WriteEntry("Connection String" + ConnectionString +"\r\n"+ NetWorkCardID, EventLogEntryType.Warning, 0);
            if (CheckNetworkCardID(NetWorkCardID))
            {
                timer.Interval = double.Parse(ConfigurationManager.AppSettings["ServiceTimer"]) * 1000; // Timer 3 วินาที
                timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
                timer.Start();
                eventLog1.WriteEntry("Start Service", EventLogEntryType.Information);
                Console.WriteLine("Timer Start");
                timerClock.Interval = 60000; // Timer
                timerClock.Elapsed += new System.Timers.ElapsedEventHandler(this.RequestTime);
                timerClock.Start();
                //SaveTextFileLog();
            }
            else
            {
                //SaveTextFileLog();
                eventLog1.WriteEntry("Invalid License of CryptBot Hi - Secure e - Office!", EventLogEntryType.Error);
            }
        }
        public void OnTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            //ResultTime = true; //Fix รันตลอดเวลา*******************************************เอาออกเวลาก่อน buid 
            if (ResultTime == true)
            {
                Console.WriteLine("Start Service");
                InsertRequestDocNoID();
            }
            else
            {
                Console.WriteLine("Pause Service");
            }

        }
        public void RequestTime(object sender, System.Timers.ElapsedEventArgs args)
        {
            //DateTime Time = Convert.ToDateTime("6/9/2017 00:00:00 AM");
            DateTime Time = DateTime.Now;
            string hour = Time.Hour.ToString();
            string minute = Time.Minute.ToString();
            string type = Time.ToString("tt");
            if (hour.Length == 1)
            {
                hour = "0" + hour;
            }
            if (minute.Length == 1)
            {
                minute = "0" + minute;
            }
            string TextTime = hour + ":" + minute + " " + type;
            Console.WriteLine(TextTime);
            if (TextTime.Equals("00:01 AM"))
            {
                ResultTime = false;
            }
            else if (TextTime.Equals("03:01 AM"))
            {
                ResultTime = true;
            }
        }
        protected override void OnStop()
        {
            eventLog1.WriteEntry("Service Stop", EventLogEntryType.Information);
        }
        protected override void OnPause()
        {
            eventLog1.WriteEntry("Service Pause", EventLogEntryType.Information);
            OnContinue();
        }
        protected override void OnContinue()
        {
            eventLog1.WriteEntry("Service Continue", EventLogEntryType.Information);
        }
        // for Process
        public void InsertRequestDocNoID()
        {
            int Queue_RequestDocNoID = 0;
            int eOfficeTransaction_ID = 0;
            int RequestDocNoType = 0;
            string FileName = "";
            int IndexWord = 0;
            string Result = "";
            string FlgStatus = "";
            DateTime  CreateDate, RevisedDate;
            HttpWebRequest request;
            DT = new DataTable();

            try
            {
                sql = "SELECT TOP (100) [Queue_RequestDocNoID],[eOfficeTransaction_ID],[RequestDocNoType],[FileName],[FlgStatus],[ErrorMsg],[CreateDate],[RevisedDate] FROM [Queue_RequestDocNo] where FlgStatus = 'W' order by Queue_RequestDocNoID asc";
                //sql = "select *  from queue_docroute where  QueueDocRouteID = 87 order by QueueDocRouteID asc ";
                //Console.WriteLine(sql);
                SelectData(sql);
            }
            catch (Exception ex)
            {
                //eventLog1.WriteEntry("Error while select Queue_RequestDocNoID  : " + sql);
                eventLog1.WriteEntry("Error while select Queue_RequestDocNoID  : " + ex.Message, EventLogEntryType.Error);
            }

            // Check แรก 
            if (DT.Rows.Count > 0) // Check ข้อมูลจาก Query มากกว่า 0
            {
                timer.Stop();
                countTimer = 0;
                try
                {
                    SaveTextFileLog();
                    File.AppendAllText(fullPath, "---------------------------------------------------------------------------------------------------------------" + Environment.NewLine);
                    File.AppendAllText(fullPath, "Queue DocRoute = " + DT.Rows.Count + " Record :" + DateTime.Now.ToString() + Environment.NewLine);
                }
                catch (Exception ex)
                {
                    eventLog1.WriteEntry("Save Log to text file Error : " + ex.Message, EventLogEntryType.Error);
                }
            }
            else
            {
                timer.Stop();
                countTimer += 1;
                try
                {
                    SaveTextFileLog();
                }
                catch (Exception ex)
                {
                    eventLog1.WriteEntry("Save Log to text file Error : " + ex.Message, EventLogEntryType.Error);
                }

                if (countTimer == 20)
                {
                    File.AppendAllText(fullPath, " Waiting for Queue :" + DateTime.Now.ToString() + Environment.NewLine);
                    countTimer = 0;
                }
            }

            // Check 2
            if (DT.Rows.Count > 0)
            {
                //eventLog1.WriteEntry("Queue Doc Route = " + DT.Rows.Count, EventLogEntryType.Warning, 0);

                timer.Stop();
                for (int i = 0; i < DT.Rows.Count; i++)
                {
                    int t = i;
                    try
                    {
                        Queue_RequestDocNoID = Convert.ToInt32(DT.Rows[t][0]);
                        eOfficeTransaction_ID = Convert.ToInt32(DT.Rows[t][1]);
                        RequestDocNoType = Convert.ToInt32(DT.Rows[t][2]);
                        FileName = Convert.ToString(DT.Rows[t][3]);
                        FlgStatus = Convert.ToString(DT.Rows[t][4]);
                        //ErrorMsg = Convert.ToString(DT.Rows[0][5]);//null
                        CreateDate = Convert.ToDateTime(DT.Rows[t][6]);
                        //RevisedDate = Convert.ToDateTime(DT.Rows[0][7]);//null
                    }
                    catch (Exception ex)
                    {
                            File.AppendAllText(fullPath, "Convert variable Failed : " + DateTime.Now.ToString()+ Environment.NewLine);
                        eventLog1.WriteEntry("Error while convert variable from queue_docroute  : " + ex.Message, EventLogEntryType.Error);
                            Console.WriteLine(ex.Message);
                    }

                    HttpWebResponse response = null;
                    //File.AppendAllText(fullPath, "Queue_RequestDocNoID :  " + Queue_RequestDocNoID + "\r\n"
                    //                + "eOfficeTransaction_ID :  " + eOfficeTransaction_ID + "\r\n" + "RequestDocNoType : " + RequestDocNoType + "\r\n"
                    //                + "FileName :  " + FileName + "\r\n"
                    //                + "FlgStatus :  " + FlgStatus + "\r\n"
                    //                + DateTime.Now.ToString() + Environment.NewLine);
                    try
                    {
                        //กำหนด Get ตัวแปรที่จะส่งค่าไปในหน้า ASP
                        request = (HttpWebRequest)HttpWebRequest.Create(Web_Server + FileName);
                        Console.WriteLine(request);

                        request.Method = "GET";
                        //กำหนด Time Out
                        request.Timeout = System.Threading.Timeout.Infinite;
                        //eventLog1.WriteEntry("Start HttpWebRequest wait for responding", EventLogEntryType.Information, 2);
                        Console.WriteLine("----------------------------Start HttpWebRequest wait for responding----------------------------");
                        //ทำการ response แล้วรอรับค่ากลับจาก ASP
                        response = (HttpWebResponse)request.GetResponse();
                        Console.WriteLine(response);
                        Console.WriteLine("----------------------------Responsed----------------------------");

                        StreamReader sr = new StreamReader(response.GetResponseStream());
                        //ผลที่ response กลับมา
                        Result = sr.ReadToEnd();
                        File.AppendAllText(fullPath, "Request Result : " + Result + Environment.NewLine);
                        //eventLog1.WriteEntry("Result: \r\n"+Result, EventLogEntryType.Information, 3);

                        //ถ้า File ASP ทำงานสำเร็จ จะส่ง tag กลับมาว่า MaxserviceCompleted

                        IndexWord = Result.IndexOf("RequestDocNoComplete");
                        if (IndexWord == 0) //เจอ ทำการ Update Status 'S'
                        {
                            //File.AppendAllText(fullPath, "Request >0 : " + IndexWord + "\r\n" + Environment.NewLine);
                            File.AppendAllText(fullPath, "Request success : " + DateTime.Now.ToString() + Environment.NewLine);
                            MoveQueueToHistory(Queue_RequestDocNoID);
                            Console.WriteLine("Maxservice Completed");
                        }
                        else if (IndexWord == -1)// ไม่เจอ  ทำการ  Update Status 'F'
                        {
                            //File.AppendAllText(fullPath, "Request == -1 : " + IndexWord + "\r\n" + Environment.NewLine);
                            UpdateFlgStatus(Queue_RequestDocNoID, FileName);
                            Console.WriteLine("Maxservice Incompleted");
                        }
                    }
                    catch (WebException e)
                    {
                        File.AppendAllText(fullPath, "Request Failed : " + e.Message  + DateTime.Now.ToString() + Environment.NewLine);
                        //แสดง Error Code
                        if (e.Status == WebExceptionStatus.ProtocolError)
                        {
                            response = (HttpWebResponse)e.Response;
                            Console.WriteLine("Errorcode: {0}", (int)response.StatusCode);
                        }
                        else
                        {
                            Console.WriteLine("Error: {0}", e.Status);
                        }
                        WebResponse errResp = e.Response;
                        string textError;
                        using (Stream respStream = errResp.GetResponseStream())
                        {
                            StreamReader reader = new StreamReader(respStream);
                            textError = reader.ReadToEnd();
                        }
                        //แสดง Error ที่แสดงออกมาจากการ Respone

                        string ErrorMsg = "Description : " + textError + "\r\n";


                        ErrorMsg = getErrorMsg(TakeLastLines(ErrorMsg, 6));
                        ErrorMsg = ErrorMsg.Replace("\"", "\'");
                        ErrorMsg = " Error Code: " + Convert.ToInt32(response.StatusCode) + "  :  " + response.StatusDescription + " ; " + ErrorMsg;
                        ErrorMsg = ErrorMsg.Replace("<font face='Arial' size=2>", " ");
                        ErrorMsg = ErrorMsg.Replace("</font>", " ; ");
                        ErrorMsg = ErrorMsg.Replace("<p>", " ");
                        ErrorMsg = ErrorMsg.Replace("</p>", " ; ");
                        ErrorMsg = ErrorMsg.Replace("<pre>", " ");
                        ErrorMsg = ErrorMsg.Replace("</pre> ", " ; ");
                        ErrorMsg = ErrorMsg.Replace("\'", "\"");

                        eventLog1.WriteEntry(ErrorMsg, EventLogEntryType.Error);
                        File.AppendAllText(fullPath, "Request Failed : " + e.Message + "\r\n" + DateTime.Now.ToString() + Environment.NewLine);

                        Console.WriteLine("Error: {0}", ErrorMsg);

                        UpdateFlgStatusRequestDocNo(Queue_RequestDocNoID, ErrorMsg);
                    }
                    //เคลียร์ Respone ทุกครั้งเมื่อเสร็จ Loop
                    finally
                    {
                        lines.Clear();

                        if (response != null)
                        {
                            response.Close();
                        }
                    }
                }
                Console.WriteLine("Status: {0}", "Finished");
                //eventLog1.WriteEntry("End Process for QueueDocRouteID :   "+ QueueDocRouteID, EventLogEntryType.Information, 5); 
                File.AppendAllText(fullPath, "Process complete for Queue_RequestDocNoID :  " + Queue_RequestDocNoID + " : " + DateTime.Now.ToString() + Environment.NewLine);
                File.AppendAllText(fullPath, "---------------------------------------------------------------------------------------------------------------" + Environment.NewLine);

            }
            else
            {
                //eventLog1.WriteEntry("Queue Doc Route = " + DT.Rows.Count, EventLogEntryType.Information, 0);
                //File.AppendAllText(fullPath, "Waiting for Queue" + DateTime.Now.ToString() + Environment.NewLine);
                Console.WriteLine("Status: {0}", "Waiting for Queue");
            }
            if (timer.Enabled == false)
            {
                timer.Start();
            }
        }
        public string getErrorMsg(List<string> ArrMsg)
        {
            string strMsg = "";
            for (int i = 0; i < lines.Count; i++)
            {
                strMsg += lines[i];
            }
            return strMsg;
        }
        public List<string> TakeLastLines(string text, int count)
        {
            Match match = Regex.Match(text, "^.*$", RegexOptions.Multiline | RegexOptions.RightToLeft);

            while (match.Success && lines.Count < count)
            {
                lines.Insert(0, match.Value);
                match = match.NextMatch();
            }

            return lines;
        }

        //Move ไปที่ Update Success เมื่อการทำงานเสร็จสมบูรณ์
        public void MoveQueueToHistory(int Queue_RequestDocNoID)
        {
          
            sql = " update [Queue_RequestDocNo] set FlgStatus ='S' ,ErrorMsg = '' , RevisedDate = getdate()where Queue_RequestDocNoID = '" + Queue_RequestDocNoID + "'";
            try
            {
                UpdateData(sql);
                //File.AppendAllText(fullPath, "Query String : " + sql +  Environment.NewLine);
                File.AppendAllText(fullPath, "Update Queue_RequestDocNo Success : " + DateTime.Now.ToString() + Environment.NewLine);
            }
            catch (Exception ex)
            {
                File.AppendAllText(fullPath, "Update Queue_RequestDocNo Failed : " + DateTime.Now.ToString() + Environment.NewLine);
                eventLog1.WriteEntry("Error while Update Queue_RequestDocNo: " + ex.Message, EventLogEntryType.Error);
            }

        }
        public void UpdateFlgStatus(int Queue_RequestDocNoID, string FileName)
        {
            sql = "update[Queue_RequestDocNo] set FlgStatus = 'F', ErrorMsg = 'File " + FileName + " not respone RequestDocNoComplete' , RevisedDate = getdate() where Queue_RequestDocNoID = '" + Queue_RequestDocNoID+"'"; 
            try
            {
                UpdateData(sql);
                //File.AppendAllText(fullPath, "Query String : " + sql  + Environment.NewLine);
                File.AppendAllText(fullPath, "update Queue_RequestDocNo success : " + DateTime.Now.ToString()  + Environment.NewLine);
            }
            catch (Exception ex)
            {
                File.AppendAllText(fullPath, "update Queue_RequestDocNo Failed : " + DateTime.Now.ToString() + Environment.NewLine);
                //File.AppendAllText(fullPath, "Query String : " + sql  + DateTime.Now.ToString() + Environment.NewLine);

                eventLog1.WriteEntry("Error update Queue_RequestDocNo: " + ex.Message, EventLogEntryType.Error);
            }
            //eventLog1.WriteEntry("Insert to DocRoute Failed" + "\r\n" + " Update FlgStatus = 'F' ", EventLogEntryType.Error, 4);
            //SendEmail("File " + FileName + " don't respone MaxserviceCompleted", QueueDocRouteID);
        }
        public void UpdateFlgStatusRequestDocNo(int Queue_RequestDocNoID, string ErrorMsg)
        {
            sql = " update [Queue_RequestDocNo] set FlgStatus ='F' ,ErrorMsg = '" + ErrorMsg + "', RevisedDate = getdate() where Queue_RequestDocNoID = " + Queue_RequestDocNoID + "'";

            try
            {
                UpdateData(sql);
                //File.AppendAllText(fullPath, "Query String : " + sql + Environment.NewLine);
                File.AppendAllText(fullPath, "update Queue_RequestDocNo success : " + DateTime.Now.ToString() + Environment.NewLine);
            }
            catch (Exception ex)
            {
                File.AppendAllText(fullPath, "update Queue_RequestDocNo Failed : " + DateTime.Now.ToString() + Environment.NewLine);
                eventLog1.WriteEntry("Error update Queue_RequestDocNo: " + ex.Message, EventLogEntryType.Error);
            }
            //eventLog1.WriteEntry("Insert to DocRoute Failed"+"\r\n"+" Update FlgStatus = 'F' ", EventLogEntryType.Error, 4);
            //SendEmail(ErrorMsg,QueueDocRouteID);
        }
        public void SaveTextFileLog()
        {
            //สร้างไฟล์เพื่อเก็บ Log ของแต่ละวัน
            path = (ConfigurationManager.AppSettings["PathFileLog"]).ToString();

            filename = DateTime.Now.Year + "_" + DateTime.Now.Month + "_" + DateTime.Now.Day + ".txt";
            fullPath = path + filename;
            ExistsFile = File.Exists(fullPath);

            //เช็คว่ามีโฟล์เดอร์หรือยัง ถ้าไม่มีจะสร้างโฟลเดอร์ขึ้นใหม่
            if (!Directory.Exists(path))
            {
                DirectoryInfo di = Directory.CreateDirectory(path);
            }

            //เช็คว่าวันนั้นสร้างไฟล์แล้วหรือยัง ถ้ายังจะสร้างใหม่
            if (ExistsFile == false)
            {
                using (FileStream fs = new FileStream(fullPath, FileMode.Create))
                {
                    using (StreamWriter w = new StreamWriter(fs, Encoding.UTF8))
                    {
                        w.WriteLine("--------------------------------------------------CreatFile At " + DateTime.Now.ToString() + "--------------------------------------------------");
                    }
                }
            }

            //สร้างไฟล์เพื่อเก็บ Log ของแต่ละวัน
        }
        public void ConvertConfig()
        {
            //var path = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "ipservice_Standard.ceo");

            string path = (ConfigurationManager.AppSettings["PathFileCBLicense"]).ToString();
            string dtext = ReadAndDecryptFile(path);

            //System.Console.WriteLine("Decrypt of ipservice.ceo = \n{0}", dtext);

            //System.Console.WriteLine("=======================================");
            string[] arrsvrconfig = dtext.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            //for (int i = 0; i < arrsvrconfig.Length; i++)
            //{
            //    System.Console.WriteLine(" {0}:{1}", i, arrsvrconfig[i]);
            //}

            if (arrsvrconfig.Length > 1)
            {
                svr1 = arrsvrconfig[0];                                 //web server name
                eofficeid = arrsvrconfig[1];                            //e-Office ID
                cbl = arrsvrconfig[2];                                  //CryptBot License Serial No.
                NetWorkCardID = arrsvrconfig[3];                        //NetWork CardID
                // database docRoute
                dbu = arrsvrconfig[40];                                 //database user name
                dbp = arrsvrconfig[41];                                 //database password 
                dbn = arrsvrconfig[42];                                 //database name
                dbs = arrsvrconfig[43];                                 //database source name

            }
        }
        public string ReadAndDecryptFile(string path)
        {
            string text = System.IO.File.ReadAllText(path);
            return Decrypt(text);
        }
        public string Decrypt(string s)
        {
            string d;
            d = OddEvenAlternateDecode(s);
            //Console.WriteLine("OddEvenAlternateDecode:\n" + d);
            return HexToString(d);
        }
        public string OddEvenAlternateDecode(string s)
        {
            int sHalf;
            string sOut = "";
            string str = s;
            sHalf = str.Length / 2;
            for (int iChr = 0; iChr < sHalf; iChr++)
            {
                sOut = sOut + str.Substring(sHalf - iChr - 1, 1) + str.Substring(sHalf + iChr, 1);
            }
            return sOut;
        }
        public string HexToString(string hx)
        {
            string rturn = "";
            char ch;
            string ansii;
            string hx1;
            try
            {
                if ((hx.Length % 3) != 0)
                {
                    //hx = hx.PadRight(hx.Length - 1);
                    hx = hx.Substring(1, hx.Length - 1);
                }

                for (int index = 0; index < hx.Length; index = index + 3)
                {
                    hx1 = String.Format("{0:X}", hx.Substring(index + 1, 2));
                    //Console.WriteLine("hx1:" + hx1);
                    ansii = hx1;
                    //Console.WriteLine("ansii:" + ansii);
                    int ansii2 = Convert.ToInt32(ansii, 16);
                    //Console.WriteLine("ansii2:" + ansii2);
                    ch = Convert.ToChar(ansii2);
                    //Console.WriteLine("ch:" + ch);
                    rturn = rturn + ch.ToString();
                }
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry("Convert Config is invalid    " + ex.Message, EventLogEntryType.Error);
                Console.WriteLine("Convert Config is invalid  " + ex.Message);
            }
            return rturn;
        }
        public bool CheckNetworkCardID(string NetworkCardID)
        {
            string LicenseInfo = "none";
            bool checkValidCardID;

            //NetworkCardID = "571899437FC3464CE0320E73955D69F7"; //อ่านจาก ipservice.ceo (ถูกต้อง)
            //NetworkCardID = "00000000000000000000000000000000"; //อ่านจาก ipservice.ceo (ไม่ถูกต้อง)

            //ถ้า มีการกำหนด NetworkCardID ใน ipservice.ceo จะนำมาเทียบค่าจาก CryptBotLib.Utility
            //ถ้า กำหนดเป็น Invalid ไม่ต้องทำ ใช้สำหรับกรณีเครื่องนั้นๆไม่ได้ลงโปรแกรม CryptBotLib.Utility
            //if (NetworkCardID.Length > 0)
            try
            {
                if (!NetworkCardID.Equals("00000000000000000000000000000000"))
                {
                    CryptBotLib h = new CryptBotLib();
                    LicenseInfo = h.GetLicenseInfo;
                }
                //Console.WriteLine("LicenseInfo:{0}", LicenseInfo);

                if (NetworkCardID != LicenseInfo)
                {
                    if (NetworkCardID.Equals("00000000000000000000000000000000"))
                    {
                        checkValidCardID = true;
                    }
                    else
                    {
                        checkValidCardID = false;
                        eventLog1.WriteEntry("Invalid License NetworkCardID of CryptBot Hi-Secure e-Office! ", EventLogEntryType.Error);
                        Console.WriteLine("Invalid License of CryptBot Hi-Secure e-Office!");
                    }
                }
                else
                {
                    checkValidCardID = true;
                    eventLog1.WriteEntry("OK! ", EventLogEntryType.Information, 0);
                    Console.WriteLine("OK!");
                }
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry("Invalid License NetworkCardID of CryptBot Hi-Secure e-Office! " + ex.Message, EventLogEntryType.Error);
                checkValidCardID = false;
            }

            return checkValidCardID;
        }
    }
}
