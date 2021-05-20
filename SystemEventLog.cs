using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SystemLogDB
{
    public partial class Form1 : Form
    {
        public string serverIP { get; set; }
        public Dictionary<string, string> servIPDictionary { get; set; } = new Dictionary<string, string>();

        public Form1()
        {
            InitializeComponent();
            IPHostEntry ipHostInfo = Dns.GetHostEntry(Dns.GetHostName()); // `Dns.Resolve()` method is deprecated.
            for (int i = 0; i < ipHostInfo.AddressList.Length; i++)
            {
                IPAddress ipAddress = ipHostInfo.AddressList[i];
                if (ipAddress.ToString().Contains("xx.xx.xxx"))
                {
                    serverIP = ipAddress.ToString();
                    break;
                }
            }
            addServerIP();
            //insertSystemLogs();
            //Form1 form = new Form1();    // causes infinite loop, so use this.Close();
            //form.Close();
            //testlabel.Text = "Finished";
            if (EventLog.Exists("Running Service - BizActor"))
                insertSystemLogs("Running Service - BizActor");
            if (EventLog.Exists("Runn_1 Running Service - BizActor"))
                insertSystemLogs("Runn_1 Running Service - BizActor");

            this.Close();
        }
        public void insertSystemLogs(string eventLogArea)
        {
            string sqlConnectionString = string.Format("Server=xx.xx.xxx.xxx,xxxx; database={0};user id={1};password={2}; Integrated Security=false;", "DBNAME", "ACCOUNT_ID", "PASSWORD");
            SqlConnection conn = conn = new SqlConnection(sqlConnectionString);

            var startTime = Convert.ToDateTime(DateTime.Now.AddMinutes(-30).AddSeconds(1).ToString("yyyy-MM-dd HH:mm:ss"));  // if you want to get past logs, change this date
            var endTime = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            //var startTime = Convert.ToDateTime("2020-12-01 00:00:00");
            //var endTime = Convert.ToDateTime("2020-12-23 00:00:00");

            // @"*[System[TimeCreated[@SystemTime >= '{0}']]] and *[System[TimeCreated[@SystemTime <= '{1}']]] and *[System/EventID=" + eventID + "]"

            var query = string.Format(@"*[System[TimeCreated[@SystemTime >= '{0}']]] and *[System[TimeCreated[@SystemTime <= '{1}']]]",
                startTime.ToUniversalTime().ToString("o"),
                endTime.ToUniversalTime().ToString("o"));      // This way is better than EventLog since EventLog method takes very very long time and process fails

            EventLogQuery logQuery = null;

            logQuery = new EventLogQuery(eventLogArea, PathType.LogName, query);

            var reader = new EventLogReader(logQuery);
            List<EventLogClass> eventLogClasses = new List<EventLogClass>();
            for (EventRecord eventRecord = reader.ReadEvent(); null != eventRecord; eventRecord = reader.ReadEvent())
            {
                try
                {
                    EventLogClass eventLogClass = new EventLogClass
                    {
                        TimeWritten = Convert.ToDateTime(eventRecord.TimeCreated),
                        Index = Convert.ToInt32(eventRecord.RecordId),
                        EventID = eventRecord.Id,
                        Message = eventRecord.FormatDescription(),
                        Category = eventRecord.OpcodeDisplayName,
                        Source = eventRecord.LevelDisplayName,
                        EntryType = eventRecord.ProviderName
                    };
                    eventLogClasses.Add(eventLogClass);
                }
                catch (Exception e)  // Catch unknown err
                {
                    continue;
                }
            }
            if (eventLogClasses.Where(x => x.Source.Equals("error")).Count() == 0 || eventLogClasses.Count == 0)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"D:\SystemBizLogTXT\SystemBizLog.txt", true)) // txt file should exist (create it beforehand)
                {
                    file.WriteLine("RunTime : " + endTime + " -- [" + eventLogArea +  "] No Error Logs at this time");
                    return;
                }
            }
            try
            {
                conn.Open();

                foreach (var item in eventLogClasses)
                {
                    string source = item.Source;
                    if (source != "Error") continue;
                    string timeWritten = item.TimeWritten.ToString("yyyy-MM-dd HH:mm:ss");
                    string index = item.Index.ToString();
                    string eventId = item.EventID.ToString();
                    string message = item.Message;
                    string category = item.Category;
                    string entryType = item.EntryType;
                    string firstLine = "";
                    try
                    {
                        firstLine = item.Message.IndexOf("\r\n") > 0 ? item.Message.Substring(0, item.Message.IndexOf("\r\n")) : item.Message;
                    }
                    catch (Exception e1)
                    {
                        firstLine = "";
                    }
                    string bizName = "";
                    try
                    {
                        if (entryType == "S_ExecuteService")
                        {
                            if (firstLine.IndexOf("Is Not Serviced In this System") < 0)
                                bizName = message.Substring(message.IndexOf(":") + 1, (message.IndexOf(")") - message.IndexOf(":") - 1));
                            else
                                bizName = message.Substring(message.IndexOf(":") + 1, (message.IndexOf(" Is Not Serviced In this System") - message.IndexOf(":") - 1));
                        }
                    }
                    catch (Exception e2)
                    {
                        bizName = "";
                    }
                    string sqlQuery = "";
                    if(eventLogArea == "Running Service - BizActor")
                    {
                        sqlQuery = string.Format("INSERT INTO MY_TABLE " +
                        "(x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12) " +
                        "VALUES (N'{0}', N'{1}', N'{2}', N'{3}', {4}, N'{5}', N'{6}', N'{7}', N'{8}', N'{9}', N'{10}', 'N') ",
                        servIPDictionary[serverIP], serverIP, eventLogArea, timeWritten, "GETDATE()", eventId, message.Replace("'", ""), source.Replace("'", ""),
                        entryType.Replace("'", ""), firstLine.Replace("'", ""), bizName.Replace("'", ""));
                    }
                    else if(eventLogArea == "Runn_1 Running Service - BizActor" && servIPDictionary.ContainsKey(serverIP))
                    {
                        sqlQuery = string.Format("INSERT INTO MY_TABLE " +
                        "(x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12) " +
                        "VALUES (N'{0}', N'{1}', N'{2}', N'{3}', {4}, N'{5}', N'{6}', N'{7}', N'{8}', N'{9}', N'{10}', 'N') ",
                        servIPDictionary[serverIP] + " EIF", serverIP + " EIF", eventLogArea, timeWritten, "GETDATE()", eventId, message.Replace("'", ""), source.Replace("'", ""),
                        entryType.Replace("'", ""), firstLine.Replace("'", ""), bizName.Replace("'", ""));
                    }
                    
                    SqlCommand cmd = new SqlCommand(sqlQuery, conn);
                    cmd.ExecuteNonQuery();
                }
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"D:\SystemBizLogTXT\SystemBizLog.txt", true)) // txt file should exist (create it beforehand)
                {
                    conn.Close();
                    file.WriteLine("RunTime : " + endTime + " -- [" + eventLogArea + "] Server EventLog Gathering Success");
                }
            }
            catch (Exception e)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"D:\SystemBizLogTXT\SystemBizLog.txt", true))
                {
                    file.WriteLine("RunTime : " + endTime + " -- [" + eventLogArea + "]  ErrMessage: " + e.Message.Substring(0, 150) + " ...");
                }
            }
        }

        public void addServerIP()
        {
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#1");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#2");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#3");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#4");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#5");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#6");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#7");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#8");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#9");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#10");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#11");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#12");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#13");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#14");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#15");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#16");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#17");
            servIPDictionary.Add("xx.xx.xxx.xxx", "Server#18");
        }

        public class EventLogClass
        {
            public DateTime TimeWritten { get; set; }
            public int Index { get; set; }
            public int EventID { get; set; }
            public string Message { get; set; }
            public string Category { get; set; }
            public string Source { get; set; }
            public string EntryType { get; set; }

            public string FirstLine { get; set; }
        }
    }
}
