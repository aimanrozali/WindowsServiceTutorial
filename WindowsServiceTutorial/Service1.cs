using System;
using System.Timers;
using System.ServiceProcess;
using System.IO;
using System.Management;
using System.Data.SqlClient;
using System.Threading;
using Timer = System.Timers.Timer;

namespace WindowsServiceTutorial
{
    public partial class Service1 : ServiceBase
    {
        private const string ConnectionString = "Data Source=ARZ\\SQLEXPRESS;Initial Catalog=USB;Integrated Security=True;";
        private const string AuthorizationQuery = "SELECT COUNT(*) FROM [USB].[dbo].[UsbDevices] WHERE DeviceID = @DeviceID";
        Timer timer = new Timer();
        Thread t;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = 5000;
            timer.Enabled = true;

            t = new Thread(new ThreadStart(new ThreadStart(backgroundWorker1_DoWork)));
            t.Start();

        }

        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
            if((t != null) && (t.IsAlive)) 
            {
                Thread.Sleep(1000);
                t.Abort();
            }
            
        }
        
        private void OnElapsedTime(object sender, ElapsedEventArgs e)
        {
            WriteToFile("Service was ran at " + DateTime.Now);
        }

        private void disableUSB(string devicename)
        {
            WriteToFile("Reach Here");
            try
            {

                //"SELECT * FROM Win32_PnPEntity WHERE DeviceID LIKE '%{devicename}%'"
                ManagementObjectSearcher myDevices = new ManagementObjectSearcher("root\\CIMV2", @"SELECT * FROM Win32_PnPEntity WHERE DeviceID LIKE '%" + devicename + "%'");

                foreach (ManagementObject item in myDevices.Get())
                {
                    WriteToFile("DeviceID: " + (string)item["DeviceID"]);
                    //WriteToFile("Name: " + (string)item["Name"]);
                    ManagementBaseObject UWFEnable = item.InvokeMethod("Disable", null, null);
                    WriteToFile(devicename + " ::: " + UWFEnable.ToString() + "::: Disabled");
                }
                /*string cmdarg = "/disable-device \"" + devicename + "\""; 
                initiatePnP(cmdarg);*/
            }
            catch(Exception ex)
            {
                WriteToFile("Error: " + ex.Message);
            }
        }

        /*private void initiatePnP (string cmd)
        {
            try
            {
                WriteToFile("Starting CMD..");
                Process p = new Process();
                p.StartInfo = new ProcessStartInfo
                {
                    CreateNoWindow = false,
                    FileName = "cmd.exe",
                    RedirectStandardError = false,
                    //RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    Arguments = $"/C pnputil.exe {cmd}"
                };
                WriteToFile(p.StartInfo.Arguments);
                p.Start();
                WriteToFile("Executed 1 !");
                WriteToFile("OUTPUT CMD: "+ p.StandardOutput.ReadToEnd());
                p.WaitForExit();
                WriteToFile("ExitCode: " + p.ExitCode);
                *//*var newProcessInfo = new System.Diagnostics.ProcessStartInfo();
                newProcessInfo.FileName = @"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe";
                newProcessInfo.Verb = "runas";
                newProcessInfo.Arguments = string.Format("pnputil {0}", cmd);
                WriteToFile("Command: " +  newProcessInfo.Arguments);
                System.Diagnostics.Process.Start(newProcessInfo);*//*
            }
            catch (Exception ex)
            {
                WriteToFile ("Exception CMD: "+ex.ToString());
            }
            
        }*/
        private void enableUSB(string devicename)
        {
            try
            {
                ManagementObjectSearcher myDevices = new ManagementObjectSearcher("root\\CIMV2", @"SELECT * FROM Win32_PnPEntity WHERE DeviceID LIKE '%" + devicename + "%'");

                foreach (ManagementObject item in myDevices.Get())
                {
                    WriteToFile("Enable DeviceID: " + (string)item["DeviceID"]);
                    ManagementBaseObject UWFEnable = item.InvokeMethod("Enable", null, null);
                    WriteToFile(devicename + " ::: " + UWFEnable.ToString() + "::: Enabled");
                }
            }
            catch(Exception ex)
            {
                WriteToFile("Error: " + ex.Message);
            }
        }

        private void DeviceInsertedEvent(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
            var devname = "";
            
            //var property = instance.Properties;
                foreach (var property in instance.Properties)
                {
                if (property.Name == "DeviceID"){
                    
                    devname = property.Value.ToString();
                    if(CheckUSB(devname))
                    {
                        WriteToFile("Authorized");
                        string[] deviceid = devname.Split('\\');
                        enableUSB(deviceid[2]);
                    }
                    else
                    {
                        WriteToFile("Unauthorized");
                        string[] deviceid = devname.Split('\\');
                        disableUSB(deviceid[2]);
                    }
                    
                    //WriteToFile(deviceid[2]);
                    //disableUSB(deviceid[2]);
                    //enableUSB(deviceid[2]);

                    WriteToFile("------------------ START -------------------");
                    WriteToFile(property.Name + " = " + property.Value);
                }
                }
            
        }
        private bool CheckUSB (string device)
        {
            bool isAuthorize = false;
            
            try
            {
                using (var connection = new SqlConnection(ConnectionString))
                {
                    connection.Open();
                    using (var command = new SqlCommand(AuthorizationQuery,connection))
                    {
                        command.Parameters.AddWithValue("@DeviceID", device);
                        var count = (int)command.ExecuteScalar();
                        if (count > 0)
                            isAuthorize = true;
                        else
                            isAuthorize = false;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile("Database Exception: "+ ex.Message);
            }
            return isAuthorize;
        }

       /* private void DeviceRemovedEvent(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
            foreach (var property in instance.Properties)
            {
                
                WriteToFile(property.Name + " = " + property.Value);
            }
        }*/

        private void backgroundWorker1_DoWork()
        {
            WqlEventQuery insertQuery = new WqlEventQuery("SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_USBHub'");
            insertQuery.WithinInterval = new TimeSpan(0, 0, 2);

            ManagementEventWatcher insertWatcher = new ManagementEventWatcher(insertQuery);
            
            insertWatcher.EventArrived += new EventArrivedEventHandler(DeviceInsertedEvent);
            insertWatcher.Start();

            // Do something while waiting for events
            System.Threading.Thread.Sleep(20000000);
        }

        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if(!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/','_')+".txt";
            if(!File.Exists(filepath))
            {
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                    //sw.Close();
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                    //sw.Close();
                }
            }
        }
    }
}
