using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NYDOE.PMO.UserProfileSync.Services;
using System.Configuration;
using NYDOE.PMO.UserProfileSync.Entities;
using NYDOE.PMO.DAL;
using System.Collections;

namespace NYDOE.PMO.UserProfileSync
{
    partial class UserProfileSync : ServiceBase
    {
        private Timer Schedular;

        public UserProfileSync()
        {
            InitializeComponent();
            this.ServiceName = "PMOUserProfileSync";
            this.EventLog.Log = "Application";
            
            this.CanHandlePowerEvent = true;
            this.CanHandleSessionChangeEvent = true;
            this.CanPauseAndContinue = true;
            this.CanShutdown = true;
            this.CanStop = true;
            this.AutoLog = true;

        }
        public void ScheduleService()
        {
            try
            {
                Schedular = new Timer(new TimerCallback(SchedularCallback));
                string mode = ConfigurationManager.AppSettings["Mode"].ToUpper();
                SPLibrary.WriteToFile("{0} : " + "Service Mode: " + mode );

                //Set the Default Time.
                DateTime scheduledTime = DateTime.MinValue;

                if (mode == "DAILY")
                {
                    //Get the Scheduled Time from AppSettings.
                    scheduledTime = DateTime.Parse(System.Configuration.ConfigurationManager.AppSettings["ScheduledTime"]);
                    if (DateTime.Now > scheduledTime)
                    {
                        //If Scheduled Time is passed set Schedule for the next day.
                        scheduledTime = scheduledTime.AddDays(1);
                    }
                }

                if (mode.ToUpper() == "INTERVAL")
                {
                    //Get the Interval in Minutes from AppSettings.
                    int intervalMinutes = Convert.ToInt32(ConfigurationManager.AppSettings["IntervalMinutes"]);

                    //Set the Scheduled Time by adding the Interval to Current Time.
                    scheduledTime = DateTime.Now.AddMinutes(intervalMinutes);
                    if (DateTime.Now > scheduledTime)
                    {
                        //If Scheduled Time is passed set Schedule for the next Interval.
                        scheduledTime = scheduledTime.AddMinutes(intervalMinutes);
                    }
                }

                TimeSpan timeSpan = scheduledTime.Subtract(DateTime.Now);
                string schedule = string.Format("{0} day(s) {1} hour(s) {2} minute(s) {3} seconds(s)", timeSpan.Days, timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);

                //Write log file
                SPLibrary.WriteToFile("{0} : " + "Service scheduled to run after: " + schedule);
                ////////////////////////////////////////////////////////////////////////
                //Call userprofile sync
                try
                {
                    //PMO site collection
                    string site = ConfigurationManager.AppSettings["rootSite"].ToString();
                    SPUserProfile up = new SPUserProfile(site);
                    List<NYDOEUser> userColls = up.UserProfiles();
                    //Console.Write(userColls.Count);
                    // insert tasks to UserTask table
                    DataLayer dal = new DataLayer();
                    dal.ExecuteSclrBySPDelete("DeleteUsers");
                    foreach (NYDOEUser item in userColls)
                    {
                        if (item != null)
                        {
                            SortedList sl = new SortedList();
                            sl.Add("@UID", item.UID);
                            sl.Add("@FirstName", item.FirstName);
                            sl.Add("@LastName", item.LastName);
                            sl.Add("@email", item.Email);
                            sl.Add("@Title", item.JobTitle);
                            sl.Add("@Phone", item.Workphone);
                            sl.Add("@Department", item.Department);
                            sl.Add("@loginame", item.LoginName);
                            dal.ExecuteSclrBySP("InsertUserProfile", sl);

                        }

                    }
                    SPLibrary.WriteToFile("{0} : SYNC SUCCESSFUL ");
                }
                catch (Exception ex)
                {
                    SPLibrary.WriteToFile("{0} : SYNC ERROR: " + ex.Message + ex.StackTrace);
                }
                ////////////////////////////////////////////////////////////////////////

                //Get the difference in Minutes between the Scheduled and Current Time.
                int dueTime = Convert.ToInt32(timeSpan.TotalMilliseconds);

                //Change the Timer's Due Time.
                Schedular.Change(dueTime, Timeout.Infinite);
            }
            catch (Exception ex)
            {
                SPLibrary.WriteToFile("ERROR: {0} " + ex.Message + ex.StackTrace);

                //Stop the Windows Service.
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("PMOUserProfileSyncService"))
                {
                    serviceController.Stop();
                }
            }
        }

        private void SchedularCallback(object e)
        {
            SPLibrary.WriteToFile("{0} : Service Log");            
            this.ScheduleService();
        }
        protected override void OnStart(string[] args)
        {
            // TODO: Add code here to start your service.
            SPLibrary.WriteToFile("{0} : Service started");
            this.ScheduleService();
        }

        protected override void OnStop()
        {
            // TODO: Add code here to perform any tear-down necessary to stop your service.
            SPLibrary.WriteToFile("{0} : Service stopped");
            this.Schedular.Dispose();
        }

        protected override void OnShutdown()
        {
            base.OnShutdown();
        }
    }
}
