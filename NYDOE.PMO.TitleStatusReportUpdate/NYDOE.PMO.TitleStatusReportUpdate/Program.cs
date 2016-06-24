using Microsoft.SharePoint.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NYDOE.PMO.TitleStatusReportUpdate
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //Dashboard site
                string site = ConfigurationManager.AppSettings["sites"].ToString();
                using (ClientContext ctx = new ClientContext(site))
                {
                    //Get all sites under Dashboard site (recursvie)
                    Web oWeb = ctx.Web;
                    WebCollection Webs = oWeb.Webs;
                    ctx.Load(Webs);
                    ctx.ExecuteQuery();
                    var websFromLooping = new ArrayList();
                    GetListOfWebs(ctx, Webs, websFromLooping);

                    //Go through each web, check Weekly meeting of each project by get information from Project Statement list
                    //and insert status report for each project accordingly
                    foreach (string web in websFromLooping)
                    {
                        ClientContext subContext = new ClientContext(web);
                        List StatusReport = GetListByTitle(subContext, ConfigurationManager.AppSettings["statusreport"].ToString());
                        if (StatusReport != null)
                        {
                            CamlQuery oQuery = new CamlQuery();
                            oQuery.ViewXml = @"<View>" +
                                             "<OrderBy><FieldRef Name='ID' /></OrderBy></Query>" +
                                            "<ViewFields>" +
                                            "<FieldRef Name='Status_x0020_Start_x0020_Date' IncludeTimeValue='FALSE' /><Value Type='DateTime'>" +
                                            "<FieldRef Name='Status_x0020_End_x0020_Date' IncludeTimeValue='FALSE' /><Value Type='DateTime'>" + 
                                            "</ViewFields>" +
                                            "</View>";
                            ListItemCollection items = StatusReport.GetItems(oQuery);
                            subContext.Load(items);
                            subContext.ExecuteQuery();

                            if (items.Count > 0)
                            {
                                foreach(ListItem item in items)
                                {
                                    var startDate = item["Status_x0020_Start_x0020_Date"]!=null? DateTime.Parse(item["Status_x0020_Start_x0020_Date"].ToString()).ToShortDateString() : "N/A";
                                    var endDate = item["Status_x0020_End_x0020_Date"] != null ? DateTime.Parse(item["Status_x0020_End_x0020_Date"].ToString()).ToShortDateString() : "N/A";
                                    item["Title"] = String.Format("Report from {0} to {1}", startDate, endDate);
                                    item.Update();                                    
                                }
                                subContext.ExecuteQuery();
                            }

                        }

                    }
                }
                Console.WriteLine("Press any key to continue ....");
                Console.Read();
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
        }
        /// <summary>
        /// Check list exsited or not in site
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="listTitle"></param>
        /// <returns>List</returns>
        private static List GetListByTitle(ClientContext clientContext, String listTitle)
        {
            List existingList;

            Web web = clientContext.Web;
            ListCollection lists = web.Lists;

            IEnumerable<List> existingLists = clientContext.LoadQuery(
                    lists.Where(
                    list => list.Title == listTitle)
                    );
            clientContext.ExecuteQuery();

            existingList = existingLists.FirstOrDefault();

            return existingList;
        }
        /// <summary>
        /// Get all siites under given site URL and add site URL to webs list
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="webCollection"></param>
        /// <param name="webs"></param>
        private static void GetListOfWebs(ClientContext ctx, IEnumerable<Web> webCollection, ArrayList webs)
        {
            foreach (Web web in webCollection)
            {
                if (web.WebTemplate.Equals(ConfigurationManager.AppSettings["sitetemplate"].ToString()))
                {
                    webs.Add(web.Url);
                    WebCollection Webs = web.Webs;
                    ctx.Load(Webs);
                    ctx.ExecuteQuery();
                    if (Webs.Count() > 0)
                    {
                        GetListOfWebs(ctx, Webs, webs);
                    }
                }

            }
        }
    }
}
