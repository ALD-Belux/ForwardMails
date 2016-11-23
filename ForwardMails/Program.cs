using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using Serilog;
using InfluxDB.Collector;

namespace ForwardMailsService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            bool TimeSeriesLogging = Properties.Settings.Default.TimeSeriesLogging;
            string TimeSeriesDBAddress = Properties.Settings.Default.TimeSeriesDBAddress;
            string TimeSeriesDBName = Properties.Settings.Default.TimeSeriesDBName;
            string TimeSeriesDBUser = Properties.Settings.Default.TimeSeriesDBUser;
            string TimeSeriesDBPassword = Properties.Settings.Default.TimeSeriesDBPassword;


            try
            {
                Log.Logger = new LoggerConfiguration()
                           .ReadFrom.AppSettings()
                           .CreateLogger();

                Log.Debug("Hello Serilog!");
                Log.Information("Starting Process");
            }
            catch (Exception)
            {
                throw;
            }


            if (TimeSeriesLogging && (string.IsNullOrEmpty(TimeSeriesDBAddress) || string.IsNullOrEmpty(TimeSeriesDBName)))
            {
                Log.Error("Time Series Logging enable but parameter(s) missing. Exiting...");
                return;
            }

            if (string.IsNullOrEmpty(TimeSeriesDBUser))
            {
                TimeSeriesDBUser = null;
            }
            if (string.IsNullOrEmpty(TimeSeriesDBPassword))
            {
                TimeSeriesDBPassword = null;
            }


            if (TimeSeriesLogging)
            {
                try
                {
                    Metrics.Collector = new CollectorConfiguration()
                      .Tag.With("host", Environment.GetEnvironmentVariable("COMPUTERNAME"))
                      .Batch.AtInterval(TimeSpan.FromSeconds(2))
                      .WriteTo.InfluxDB(TimeSeriesDBAddress,TimeSeriesDBName,TimeSeriesDBUser,TimeSeriesDBPassword)
                      .CreateCollector();

                    Log.Information("Metrics Collector Created");
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "Unable to create Metrics Collector. Exiting.");
                    return;
                }
            }
            

#if (!DEBUG)
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new ForwardMailSVC()
            };
            ServiceBase.Run(ServicesToRun);
#else
            //Debug code: this allows the process to run
            // as a non-service. It will kick off the
            // service start point, and then run the
            // sleep loop below.
            ForwardMailSVC service = new ForwardMailSVC();
            bool done = !(service.Start());
            // Break execution and set done to true to run Stop()
            //bool done = false;
            while (!done)
                System.Threading.Thread.Sleep(10000);
            service.Stop();
#endif
        }
    }
}
