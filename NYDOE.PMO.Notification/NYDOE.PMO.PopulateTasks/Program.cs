using Microsoft.SharePoint.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NYDOE.PMO.PopulateTasks.Entities;
using NYDOE.PMO.DAL;
using System.Data;

namespace NYDOE.PMO.PopulateTasks
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //empty UserTask table
                DataLayer dal = new DataLayer();
                dal.ExecuteSclrBySPDelete("deleteTask");

                //Dashboard site
                string site = ConfigurationManager.AppSettings["sites"].ToString();
                List<UserTask> incompleteTasks = GetIncompleteTaskAll(site);
                // insert tasks to UserTask table
                foreach (UserTask item in incompleteTasks)
                {
                    if (item != null && item.UID!=-1 && item.TaskStatus!= "Completed")                   
                    {
                        SortedList sl = new SortedList();                        
                        sl.Add("@UID", item.UID);
                        sl.Add("@TID", item.TaskID);
                        sl.Add("@DueDate",item.DueDate);
                        sl.Add("@emailtemplateID", 1);
                        sl.Add("@Completed",item.IsComplete);
                        sl.Add("@TaskName", item.TaskName);
                        sl.Add("@TaskDescription",item.TaskDescription);
                        sl.Add("@TaskURL", item.TaskUrl);
                        sl.Add("@TaskStatus", item.TaskStatus);
                       

                        // DataTable dtSaveEmailStatus = new DataTable();
                        string dtSaveEmailStatus = dal.ExecuteSclrBySP("InsertTask", sl);

                    }

                }                
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
            Console.WriteLine("Press any key to continue ....");
            Console.Read();
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
        /// <summary>
        /// check list exsited or not
        /// </summary>
        /// <param name="oWeb"></param>
        /// <param name="listname"></param>
        /// <returns></returns>
        private static bool ListExists(Web oWeb, string listname)
        {
            ClientContext ctx = (ClientContext)oWeb.Context;
            try
            {
                List oList = oWeb.Lists.GetByTitle(listname);
                ctx.Load(oList);
                ctx.ExecuteQuery();
                return true;
            }
            catch (Exception ex)
            {
                
            }
            return false;
        }
        private static List<UserTask> GetIncompleteTaskAll(string siteUrl)
        {
            List<UserTask> incompleteTasks = new List<UserTask>();
            using (ClientContext ctx = new ClientContext(siteUrl))
            {
                //Get all sites under Dashboard site (recursvie)
                Web oWeb = ctx.Web;
                WebCollection Webs = oWeb.Webs;
                ctx.Load(Webs);
                ctx.ExecuteQuery();
                var websFromLooping = new ArrayList();
                GetListOfWebs(ctx, Webs, websFromLooping);

                foreach (string webUrl in websFromLooping)
                {
                    ClientContext subContext = new ClientContext(webUrl);
                    // check task list existed
                    Web osubWeb = subContext.Web;
                    if (ListExists(osubWeb, ConfigurationManager.AppSettings["tasklist"].ToString()))
                    {
                        try
                        {
                            List list = subContext.Web.Lists.GetByTitle(ConfigurationManager.AppSettings["tasklist"].ToString());
                            CamlQuery camlQuery = new CamlQuery();

                            string viewQueryTask = @"                                        
                                        <View> 
                                        <Query><Where>
                                            <Or>
                                                <Or>
                                                <And>
                                                    <Or>
                                                        <And>
                                                            <IsNotNull>
                                                                <FieldRef Name='AssignedTo' />
                                                            </IsNotNull>
                                                            <Neq>
                                                                <FieldRef Name='PercentComplete' /><Value Type='Number'>1.00</Value>
                                                            </Neq>
                                                        </And>
                                                        <Neq>
                                                            <FieldRef Name='Status' /><Value Type='Choice'>Completed</Value>
                                                        </Neq>
                                                    </Or>
                                                    <Eq>
                                                          <FieldRef Name='DueDate' IncludeTimeValue='FALSE' /><Value Type='DateTime' IncludeTimeValue='FALSE'><Today  OffsetDays='-3' /></Value>
                                                    </Eq>
                                                </And>
                                                    <Eq>
                                                            <FieldRef Name='DueDate' IncludeTimeValue='FALSE' /><Value Type='DateTime' IncludeTimeValue='FALSE'><Today  OffsetDays='3' /></Value>
                                                    </Eq>
                                                </Or>
                                                    <Eq>
                                                        <FieldRef Name='DueDate' IncludeTimeValue='FALSE' /><Value Type='DateTime' IncludeTimeValue='FALSE'><Today  OffsetDays='0' /></Value>
                                                    </Eq>
                                            </Or>
                                            </Where> </Query>                                                            
                                         <OrderBy>
                                                <FieldRef Name='Modified' Ascending='False' />
                                        </OrderBy>
                                        <ViewFields><FieldRef Name='ID'/><FieldRef Name='Title'/><FieldRef Name='Body' /><FieldRef Name='Status'/><FieldRef Name='Created'/><FieldRef Name='DueDate'/><FieldRef Name='AssignedTo' />
                                        </ViewFields></View>";




                            camlQuery.ViewXml = viewQueryTask;
                            ListItemCollection items = list.GetItems(camlQuery);
                            subContext.Load(items);
                            subContext.ExecuteQuery();
                            if (items.Count > 0)
                            {
                                foreach (ListItem item in items)
                                {                                    
                                    FieldUserValue[] assignedTo = (FieldUserValue[])item["AssignedTo"];                                    
                                    UserTask userTask = new UserTask
                                    {

                                        UID = item["AssignedTo"] !=null ? assignedTo[0].LookupId : -1,
                                        TaskID = item["ID"] != null ? item["ID"].ToString() : string.Empty,
                                        TaskName = item["Title"] !=null ? item["Title"].ToString(): string.Empty,
                                        TaskDescription = item["Body"] != null ? item["Body"].ToString() : string.Empty,
                                        TaskUrl = String.Format("{0}/Lists/Tasks/DispForm.aspx?ID={1}", webUrl, item["ID"]),
                                        TaskStatus = item["Status"] != null ? item["Status"].ToString() : string.Empty,                                        
                                        DueDate = item["DueDate"] != null ? Convert.ToDateTime(item["DueDate"].ToString()) : DateTime.MinValue
                                    };
                                    incompleteTasks.Add(userTask);
                                }

                            }

                        }
                        catch (Exception e)
                        {
                            Console.Write(e);
                        }

                    }

                }
            }
            return incompleteTasks;
            
           
        }
        private static List<UserTask> GetIncompleteTasksbyProject(ref ClientContext ctx, string webURL)
        {
            List<UserTask> incompleteTasks = new List<UserTask>();
            try
            {
                List list = ctx.Web.Lists.GetByTitle(ConfigurationManager.AppSettings["tasklist"].ToString());
                CamlQuery camlQuery = new CamlQuery();

                string viewQueryTask = @"                                        
                                         <Where>                                                                                                                               
                                             <Or>
                                                 <Or>
                                                    <And>
                                                       <Neq>
                                                          <FieldRef Name='PercentComplete' /><Value Type='Number'>1.00</Value>
                                                       </Neq>
                                                       <Eq>
                                                          <FieldRef Name='DueDate' IncludeTimeValue='FALSE' /><Value Type='DateTime' IncludeTimeValue='FALSE'><Today  OffsetDays='-3' /></Value>
                                                       </Eq>
                                                    </And>
                                                    <Eq>
                                                       <FieldRef Name='DueDate' IncludeTimeValue='FALSE' /><Value Type='DateTime' IncludeTimeValue='FALSE'><Today  OffsetDays='3' /></Value>
                                                    </Eq>
                                                 </Or>
                                                 <Eq>
                                                    <FieldRef Name='DueDate' IncludeTimeValue='FALSE' /><Value Type='DateTime' IncludeTimeValue='FALSE'><Today  OffsetDays='0' /></Value>
                                                 </Eq>
                                              </Or>
                                         </Where>                                       
                                          <OrderBy>
                                                <FieldRef Name='Modified' Ascending='False' />
                                        </OrderBy>
                                        <ViewFields><FieldRef Name='ID'/><FieldRef Name='Title'/><FieldRef Name='Status'/><FieldRef Name='Created'/><FieldRef Name='DueDate'/>
                                        </ViewFields></View>";



                
                camlQuery.ViewXml = viewQueryTask;
                ListItemCollection items = list.GetItems(camlQuery);
                ctx.Load(items);
                ctx.ExecuteQuery();
                if (items.Count > 0)
                {
                    
                    foreach (ListItem item in items)
                    {
                        //var link = "<a href='" + webURL + "/Lists/Tasks/DispForm.aspx?ID=" + item["ID"] + "' target='_blank'>" + item["Title"] + "</a>";
                        UserTask userTask = new UserTask
                        {
                            //AssignedTo = item.FieldValues["AssignedTo"].ToString(),
                            TaskName = item.FieldValues["Title"].ToString()
                        };
                        
                       
                        incompleteTasks.Add(userTask);
                    }
                    
                }

            }
            catch (Exception e)
            {
                Console.Write(e);
            }
            return incompleteTasks;
        }
       
    }
}
