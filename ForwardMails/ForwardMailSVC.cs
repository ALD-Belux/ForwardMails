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
using ForwardMailsLibrary;
using Serilog;
using InfluxDB.Collector;

namespace ForwardMailsService
{
    public partial class ForwardMailSVC : ServiceBase
    {
        private string mailboxSMTP = Properties.Settings.Default.MailboxSMTP;
        private bool deleteWhenForwarded = Properties.Settings.Default.DeleteWhenForwarded;
        private Microsoft.Exchange.WebServices.Data.ExchangeVersion requestedServerVersion = Properties.Settings.Default.ExchangeVersion;
        private string forwardAddress = Properties.Settings.Default.ForwardAddress;
        private int processIntervalMs = Properties.Settings.Default.ProcessIntervalMs;
        private string srcMailsFolderName = Properties.Settings.Default.SourceInboxMailsSubFolderName;
        private string dstFolderName = Properties.Settings.Default.DestinationInboxMailsSubFolderName;
        private bool impersonate = Properties.Settings.Default.Impersonate;


        private bool methodRunning = false;
        private Timer timer = new Timer(Properties.Settings.Default.ProcessIntervalMs);
        private ForwardMails fwMails;

        public ForwardMailSVC()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            bool run = true;
            while (run)
            {
                run = !(Start());
            }
        }

        public bool Start()
        {
            try
            {
                fwMails = new ForwardMails(mailboxSMTP, srcMailsFolderName, dstFolderName, impersonate, requestedServerVersion);
            }
            catch (Exception)
            {
                Log.Error("An error occured, please review previous logs");
                return false;
            }

            timer.Elapsed += Timer_Elapsed;
            timer.Enabled = true;
            return true;
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            timer.Enabled = false;
            methodRunning = true;
            try
            {
                fwMails.ForwardMailsInFolder(forwardAddress, deleteWhenForwarded);
            }
            catch (Exception)
            {
                Log.Error("An error occured, please review previous logs");
                Metrics.Increment("ErrorOccured");
            }
            methodRunning = false;
            timer.Enabled = true;
        }

        protected override void OnStop()
        {
            timer.Enabled = false;

            while (methodRunning)
            {
                Log.Information("Waiting method to finish");
                System.Threading.Thread.Sleep(1000);
            }
            Log.Information("Stopping Service");
        }
    }
}
