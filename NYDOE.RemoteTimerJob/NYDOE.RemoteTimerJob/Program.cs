using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Collections;

namespace NYDOE.RemoteTimerJob
{
    class Program
    {
        public class StatusReport
        {                        
            public DateTime StatusStartDate { get; set; }
            public DateTime StatusEndDate { get; set; }
            public string Schedule { get; set; }
            public string Budget { get; set; }           
            public string OpenTasks { get; set; }
            public string ReportedIssues { get; set; }
            public string ReportedRisks { get; set; }
            public string CompletedTasksOver7Days { get; set; }

           

        }
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
                        List ProjectInfo = GetListByTitle(subContext, "Project Statement");
                        if (ProjectInfo != null)
                        {
                            CamlQuery oQuery = new CamlQuery();
                            oQuery.ViewXml = @"<View><Query><OrderBy><FieldRef Name='ID' /></OrderBy></Query>" +
                                            "<ViewFields><FieldRef Name='ID'/><FieldRef Name='Status_x0020_Meeting_x0020_Weekd'/></ViewFields></View>";
                            ListItemCollection items = ProjectInfo.GetItems(oQuery);
                            subContext.Load(items);
                            subContext.ExecuteQuery();
                            if (items.Count > 0)
                            {
                                ListItem item = items.FirstOrDefault();
                                subContext.Load(item);
                                subContext.ExecuteQuery();

                                // If today is weekly meeting day then run status report fot this project from last 7 days to the day before
                                if (NullorEmpty(item["Status_x0020_Meeting_x0020_Weekd"]) != String.Empty)
                                {
                                    string today = DateTime.Today.DayOfWeek.ToString();
                                    if (item["Status_x0020_Meeting_x0020_Weekd"].Equals(today))
                                    {
                                        DateTime startDate = DateTime.Today.AddDays(-7);
                                        DateTime endDate = DateTime.Today.AddDays(-1);
                                        if (today.Equals("Monday"))
                                            endDate = DateTime.Today.AddDays(-3);
                                        string issues = getIssues(ref subContext, web);
                                        string risks = getRisks(ref subContext, web);
                                        string openTasks = getOpenTasks(ref subContext, web);
                                        string completedTasksLast7days = getCompletedTasksLast7days(ref subContext, ref startDate, web);                                        
                                        // check item existed or not before inserting into Status Report 
                                        List statusReportList = subContext.Web.Lists.GetByTitle("Status Report");
                                        CamlQuery camlQuery = new CamlQuery();
                                        camlQuery.ViewXml = "<View>" +
                                                            "<Query><Where><And>" +
                                                            "<Eq><FieldRef Name='Status_x0020_Start_x0020_Date' IncludeTimeValue='FALSE' /><Value Type='DateTime'>" + startDate.ToString("yyyy-MM-ddTHH:mm:ssZ") + "</Value></Eq>" +
                                                            "<Eq><FieldRef Name='Status_x0020_End_x0020_Date' IncludeTimeValue='FALSE' /><Value Type='DateTime'>" + endDate.ToString("yyyy-MM-ddTHH:mm:ssZ") + "</Value></Eq></And></Where></Query>" +
                                                            "<ViewFields><FieldRef Name='ID'/>" +                                                            
                                                            "</ViewFields>" +
                                                            "</View>";


                                        ListItemCollection statusReport = statusReportList.GetItems(camlQuery);
                                        subContext.Load(statusReport);
                                        subContext.ExecuteQuery();
                                        bool itemIsFound = (statusReport.Count == 1);
                                        // insert an item to Status Report list if it's not existed                       
                                        if (!itemIsFound)
                                        {
                                            ListItemCreationInformation newItemFormation = new ListItemCreationInformation();
                                            ListItem newItem = statusReportList.AddItem(newItemFormation);                                            
                                            newItem["Status_x0020_Start_x0020_Date"] = startDate;
                                            newItem["Status_x0020_End_x0020_Date"] = endDate;
                                            newItem["Title"] = String.Format("Report from {0} to {1}", startDate.ToShortDateString(), endDate.ToShortDateString());                                            
                                            newItem["Current_x0020_Stage"] = item["Current_x0020_Stage"] != null ? item["Current_x0020_Stage"].ToString() : string.Empty;                                            
                                            newItem["Budget"] = "(1) On Schedule";
                                            newItem["Health"] = "(1) On Schedule";
                                            newItem["Upcoming_x0020_Tasks"] = openTasks;
                                            newItem["Tasks_x0020_Completed"] = completedTasksLast7days;
                                            newItem["Reported_x0020_Issues"] = issues;
                                            newItem["Reported_x0020_Risks"] = risks;
                                            newItem.Update();
                                            subContext.ExecuteQuery();
                                       }
                                        
                                    }
                                }
                            }

                        }


                    }
                }
                //Console.WriteLine("Press any key to continue ....");
                //Console.Read();
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
        /// Get all risks has "Exclude_x0020_from_x0020_Reports" is false and top 3
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="webURL"></param>
        /// <returns></returns>
        private static string getRisks(ref ClientContext ctx, string webURL)
        {            
            string risks = string.Empty;

            try
            {                
                List list = ctx.Web.Lists.GetByTitle("Risks");
                CamlQuery query = new CamlQuery();
                query.ViewXml = "<View>"+
                                "<Query><Where>" +
                                "<Or><IsNull><FieldRef Name='Exclude_x0020_from_x0020_Reports'/></IsNull>" +
                                "<Eq><FieldRef Name='Exclude_x0020_from_x0020_Reports'/><Value Type='Boolean'>0</Value></Eq>" +
                                "</Or></Where>" +
                                "<GroupBy Collapse = 'TRUE'><FieldRef Name='Reporting_x0020_Order' /></GroupBy>" +
                                "<OrderBy><FieldRef Name='Reporting_x0020_Order' Ascending='True' /></OrderBy></Query>" +
                                "<ViewFields><FieldRef Name='ID'/><FieldRef Name='Title'/><FieldRef Name='DueDate'/><FieldRef Name='WorkstreamsId'/><FieldRef Name='Reporting_x0020_Order'/></ViewFields>"+
                                "<RowLimit>3</RowLimit></View>";               
                ListItemCollection items = list.GetItems(query);
                ctx.Load(items);
                ctx.ExecuteQuery();
                if (items.Count > 0)
                {
                    risks = "<ul>";
                    foreach (ListItem item in items)
                    {
                        var link = "<a href='" + webURL + "/Lists/Risks/DispForm.aspx?ID=" + item["ID"] + "' target='_blank'>" + item["Title"] + "</a>";
                        risks += "<li>" + link + "</li>";

                    }
                    risks+= "</ul>";
                }

            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
            return risks;
        }
        /// <summary>
        /// Get all active issues has "Exclude_x0020_from_x0020_Reports" is false
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="webURL"></param>
        /// <returns></returns>
        private static string getIssues(ref ClientContext ctx, string webURL)
        {
            string issues = string.Empty;
            try
            {
                List list = ctx.Web.Lists.GetByTitle("Issues");                
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View>"+
                                    "<Query><Where><And><Eq><FieldRef Name='Status'/><Value Type='Text'>Active</Value></Eq>" +
                                    "<Or><IsNull><FieldRef Name='Exclude_x0020_from_x0020_Reports'/></IsNull>" +
                                    "<Eq><FieldRef Name='Exclude_x0020_from_x0020_Reports'/><Value Type='Boolean'>0</Value></Eq>"+
                                    "</Or>" +
                                    "</And></Where></Query>" + 
                                    "<ViewFields><FieldRef Name='ID'/><FieldRef Name='Title'/><FieldRef Name='DueDate'/><FieldRef Name='WorkstreamsId'/>" +
                                    "</ViewFields>" + 
                                    "</View>";

                ListItemCollection items = list.GetItems(camlQuery);
                ctx.Load(items);
                ctx.ExecuteQuery();
                if (items.Count > 0)
                {
                    issues = "<ul>";
                    foreach (ListItem item in items)
                    {
                        var link = "<a href='" + webURL + "/Lists/Issues/DispForm.aspx?ID=" + item["ID"] + "' target='_blank'>" + item["Title"] + "</a>";
                        issues += "<li>" + link + "</li>";

                    }
                    issues += "</ul>";
                }

            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
            return issues;
        }
        /// <summary>
        /// Get open tasks has "Exclude_x0020_from_x0020_Reports" is false
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="webURL"></param>
        /// <returns></returns>
        private static string getOpenTasks(ref ClientContext ctx, string webURL)
        {
            string openTasks = string.Empty;
            try
            {
                List list = ctx.Web.Lists.GetByTitle("Tasks");
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View>" + 
                                    "<Query><Where><And><Neq><FieldRef Name='PercentComplete'/><Value Type='Number'>1.00</Value></Neq>" +
                                    "<Or><IsNull><FieldRef Name='Exclude_x0020_from_x0020_Reports'/></IsNull>" +
                                    "<Eq><FieldRef Name='Exclude_x0020_from_x0020_Reports'/><Value Type='Boolean'>0</Value></Eq>"+
                                    "</Or>" +
                                    "</And></Where></Query>" +
                                    "<ViewFields><FieldRef Name='ID'/>" +
                                    "<FieldRef Name='Title'/><FieldRef Name='Status'/><FieldRef Name='Created'/><FieldRef Name='DueDate'/><FieldRef Name='WorkstreamsId'/>" +
                                    "</ViewFields>" + 
                                    "</View>";

                ListItemCollection items = list.GetItems(camlQuery);
                ctx.Load(items);
                ctx.ExecuteQuery();
                if (items.Count > 0)
                {
                    openTasks = "<ul>";
                    foreach (ListItem item in items)
                    {
                        var link = "<a href='" + webURL + "/Lists/Tasks/DispForm.aspx?ID=" + item["ID"] + "' target='_blank'>" + item["Title"] + "</a>";
                        openTasks += "<li>" + link + "</li>";

                    }
                    openTasks += "</ul>";
                }

            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
            return openTasks;
        }
        /// <summary>
        /// Get all tasks comppleted last 7 days
        /// "PercentCompete" = 1 and "Exclude_x0020_from_x0020_Reports" is empty or 'No' and "Project_x0020_End_x0020_Date" greater equal than today - 7 and less than today - [if "Project_x0020_End_x0020_Date" is empty then using "Modified" field]
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="endDate"></param>
        /// <param name="webURL"></param>
        /// <returns></returns>
        private static string getCompletedTasksLast7days(ref ClientContext ctx, ref DateTime startDate, string webURL)
        {
            string completedTasksLast7days = string.Empty;
            try
            {
                List list = ctx.Web.Lists.GetByTitle("Tasks");
                CamlQuery camlQuery = new CamlQuery();
               

                string viewQueryTask = @"                                        
                                       <View>
                                       <Query><Where><And> 
                                                <And>
                                                <Eq>
                                                     <FieldRef Name='PercentComplete' /><Value Type='Number'>1.00</Value>
                                                </Eq>                                                
                                                <Or>
                                                    <IsNull><FieldRef Name='Exclude_x0020_from_x0020_Reports'/></IsNull>
                                                    <Eq><FieldRef Name='Exclude_x0020_from_x0020_Reports'/><Value Type='Boolean'>0</Value></Eq>
                                                </Or>
                                                </And>                                               
                                                <Or>
                                                <And>
                                                    <IsNull><FieldRef Name='Project_x0020_End_x0020_Date'/></IsNull>
                                                    <And>
                                                        <Geq>
                                                            <FieldRef Name='Modified' IncludeTimeValue='FALSE' /><Value Type='DateTime' IncludeTimeValue='FALSE'><Today  OffsetDays='-7' /></Value>
                                                        </Geq>
                                                        <Lt>
                                                            <FieldRef Name='Modified' IncludeTimeValue='FALSE' /><Value Type='DateTime' IncludeTimeValue='FALSE'><Today  OffsetDays='0' /></Value>
                                                        </Lt>
                                                    </And>
                                                </And>
                                                <And>
                                                    <Geq>
                                                        <FieldRef Name='Project_x0020_End_x0020_Date' IncludeTimeValue='FALSE' /><Value Type='DateTime' IncludeTimeValue='FALSE'><Today  OffsetDays='-7' /></Value>
                                                    </Geq>
                                                    <Lt>
                                                        <FieldRef Name='Project_x0020_End_x0020_Date' IncludeTimeValue='FALSE' /><Value Type='DateTime' IncludeTimeValue='FALSE'><Today  OffsetDays='0' /></Value>
                                                    </Lt>                                                           
                                                </And>
                                                 </Or>
                                               </And></Where></Query>                                      
                                        <OrderBy>
                                                <FieldRef Name='Modified' Ascending='False' />
                                        </OrderBy>
                                        <ViewFields><FieldRef Name='ID'/><FieldRef Name='Title'/><FieldRef Name='Status'/><FieldRef Name='Created'/><FieldRef Name='DueDate'/><FieldRef Name='WorkstreamsId'/>
                                        </ViewFields></View>";
                camlQuery.ViewXml = viewQueryTask;
                ListItemCollection items = list.GetItems(camlQuery);
                ctx.Load(items);
                ctx.ExecuteQuery();
                if (items.Count > 0)
                {
                    completedTasksLast7days = "<ul>";
                    foreach (ListItem item in items)
                    {
                        var link = "<a href='" + webURL + "/Lists/Tasks/DispForm.aspx?ID=" + item["ID"] + "' target='_blank'>" + item["Title"] + "</a>";
                        completedTasksLast7days += "<li>" + link + "</li>";

                    }
                    completedTasksLast7days += "</ul>";
                }

            }
            catch (Exception e)
            {
                Console.Write(e);
            }
            return completedTasksLast7days;
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
        public static string NullorEmpty(object s)
        {
            string ret = String.Empty;
            if (s != null)
            {
                try
                {
                    ret = s.ToString();
                }
                catch
                {

                }
            }
            return ret;
        }

    }

}
