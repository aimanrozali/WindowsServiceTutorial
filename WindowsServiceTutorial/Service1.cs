using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.IO;
using System.Management;

namespace WindowsServiceTutorial
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer();
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
            backgroundWorker1_DoWork();
        }

        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
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
            }
            catch(Exception ex)
            {
                WriteToFile("Error: " + ex.Message);
            }
        }
        private void enableUSB(string devicename)
        {
            ManagementObjectSearcher myDevices = new ManagementObjectSearcher("root\\CIMV2", @"SELECT * FROM Win32_PnPEntity where Name Like " + '"' + devicename + '"');

            foreach (ManagementObject item in myDevices.Get())
            {

                ManagementBaseObject UWFEnable = item.InvokeMethod("Enable", null, null);
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
                    string[] deviceid = devname.Split('\\');
                    disableUSB(deviceid[2]);

                    WriteToFile("------------------ START -------------------");
                    WriteToFile(property.Name + " = " + property.Value);
                }
                }
            
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
