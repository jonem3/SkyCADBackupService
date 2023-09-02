using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Globalization;

namespace SkyCADBackup
{
    public partial class SkyCADBackups : ServiceBase
    {

        public enum ServiceState
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ServiceStatus
        {
            public int dwServiceType;
            public ServiceState dwCurrentState;
            public int dwControlsAccepted;
            public int dwWin32ExitCode;
            public int dwServiceSpecificExitCode;
            public int dwCheckPoint;
            public int dwWaitHint;
        };

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(System.IntPtr handle, ref ServiceStatus serviceStatus);

        private int eventID = 1;
        private String SkyCADEnvironmentLoc = @"G:\Shared drives\SkyCAD\SkyCAD Environments";
        //private String SkyCADEnvironmentLocBackup = @"I:\Shared drives\SkyCAD\SkyCAD Environments";
        private string BackupLoc = @"D:\SkyCAD_Backups";

        public SkyCADBackups()
        {
            InitializeComponent();
            eventLog1 = new EventLog();
            if (!EventLog.SourceExists("SkyCADBackupSource"))
            {
                EventLog.CreateEventSource(
                    "SkyCADBackupSource", "SkyCADBackupLog");
            }
            eventLog1.Source = "SkyCADBackupSource";
            eventLog1.Log = "SkyCADBackupLog";


        }

        protected override void OnStart(string[] args)
        {
            // Update the service state to Start Pending.
            if (!Directory.Exists(BackupLoc))
            {
                Directory.CreateDirectory(BackupLoc);
            }
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            Timer timer = new Timer();
            eventLog1.WriteEntry("Starting SkyCAD Backup Service.");

            timer.Interval = 1000 * 60 * 60 * 24;
            timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            timer.Start();

            // Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

        protected override void OnStop()
        {
            // Update the service state to Stop Pending.
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOP_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            eventLog1.WriteEntry("Stopping SkyCAD Backup Service.");

            // Update the service state to Stopped.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOPPED;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            eventLog1.WriteEntry("Initiating SkyCAD Backup.", EventLogEntryType.Information, eventID++);

            DirectoryInfo ExistingBackups = new DirectoryInfo(BackupLoc);
            var Backups = ExistingBackups.GetFiles();
            CultureInfo cultureInfo = CultureInfo.GetCultureInfo("en-GB");
            foreach (FileInfo i in Backups)
            {
                eventLog1.WriteEntry($"Checking {i.Name}");
                string dateString = i.Name.Substring(13, 6);
                eventLog1.WriteEntry($"Date String {dateString}");
                DateTime FileDate = DateTime.ParseExact(dateString, "yyMMdd", cultureInfo);
                int fileAge = (DateTime.Today - FileDate).Days;
                eventLog1.WriteEntry($"File Age: {fileAge}");
                if(fileAge > 93) 
                {
                    eventLog1.WriteEntry($"Deleting {i.FullName}");
                    File.Delete(@i.FullName);
                }
            }
            
            string zipName = BackupLoc + @"\SkyCadBackup_" + DateTime.Today.ToString("yyMMdd") + ".zip";
            if (File.Exists(zipName))
            {
                eventLog1.WriteEntry("Backup Exists for Today");
            }
            else
            {
                eventLog1.WriteEntry($"Zipping Backup To {zipName}.");
                ZipFile.CreateFromDirectory(SkyCADEnvironmentLoc, zipName);
            }

        }

    }
}
