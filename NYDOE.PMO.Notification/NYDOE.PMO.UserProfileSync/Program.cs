using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using NYDOE.PMO.DAL;
using NYDOE.PMO.UserProfileSync.Entities;
using NYDOE.PMO.UserProfileSync.Services;
using System.Collections;
using System.ServiceProcess;


namespace NYDOE.PMO.UserProfileSync
{
    class Program
    {
        static void Main(string[] args)
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new UserProfileSync()
            };
            ServiceBase.Run(ServicesToRun);

            //try
            //{
            //    //PMO site collection
            //    string site = ConfigurationManager.AppSettings["rootSite"].ToString();
            //    SPUserProfile up = new SPUserProfile(site);
            //    List<NYDOEUser> userColls = up.UserProfiles();
            //    //Console.Write(userColls.Count);
            //    // insert tasks to UserTask table
            //    DataLayer dal = new DataLayer();
            //    dal.ExecuteSclrBySPDelete("DeleteUsers");                
            //    foreach (NYDOEUser item in userColls)
            //    {
            //        if (item != null)
            //        {
            //            SortedList sl = new SortedList();
            //            sl.Add("@UID ", item.UID);
            //            sl.Add("@FirstName", item.FirstName);
            //            sl.Add("@LastName", item.LastName);
            //            sl.Add("@email ", item.Email);
            //            sl.Add("@Title", item.JobTitle);
            //            sl.Add("@Phone", item.Workphone);
            //            sl.Add("@Department", item.Department);
            //            sl.Add("@loginame", item.LoginName);
            //            dal.ExecuteSclrBySP("InsertUserProfile", sl);

            //        }

            //    }
            //}
            //catch (Exception e)
            //{
            //    Console.Write(e.Message);
            //}
            

        }
    }
}
