using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace NYDOE.PMO.ClientSideRendering
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string url = ConfigurationManager.AppSettings["sites"].ToString();
                //UpdateWebpartToAll(url);
                //AddWebpartToAll(url);
                UpdateViewForAll(url);
                //clientContext.ExecuteQuery();
                //Register JS files via JSLink properties
                //RegisterJStoWebPart(web, list.DefaultViewUrl, "~sitecollection/Style%20Library/PMOCustom/js/TasksDueDate.js");
                //RegisterJStoWebPart(web, "/DashBoard/PSALPO/PSALFastTrackPO/FinSystemsBackLog1/Lists/Tasks/AllItems.aspx", ConfigurationManager.AppSettings["JSFile"].ToString()); 
                Console.WriteLine("Done");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }
        /// <summary>
        /// Add script editor to page
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="url"></param>
        /// <param name="jsPath"></param>
        private static void AddWebpartToPage(ClientContext ctx, string url, string title, string jsPath)
        {  
            //<script type=\"text/javascript\" src=\"/Style%20Library/PMOCustom/JSLink/TasksDueDate.js\"></script>     
            string xmlWebPart = "<webParts>" + 
                        "<webPart xmlns=\"http://schemas.microsoft.com/WebPart/v3\">" +
                        "<metaData>" +
                            "<type name=\"Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\" />" +
                            "<importErrorMessage>Cannot import this Web Part.</importErrorMessage>" +
                        "</metaData>" +
                        "<data>" +
                            "<properties>" +
                            "<property name=\"HelpMode\" type=\"helpmode\">Navigate</property>" +
                            "<property name=\"ExportMode\" type=\"exportmode\">All</property>" +
                            "<property name=\"HelpUrl\" type=\"string\" />" +
                            "<property name=\"Hidden\" type=\"bool\">False</property>" +
                            "<property name=\"Description\" type=\"string\">Allows authors to insert HTML snippets or scripts.</property>" +
                            "<property name=\"Content\" type=\"string\">&lt;script type=\"text/javascript\" src=\"" + jsPath + "\"&gt;&lt;/script&gt;</property>" +
                            "<property name=\"CatalogIconImageUrl\" type=\"string\" />" +
                            "<property name=\"Title\" type=\"string\">" + title + "</property>" +
                            "<property name=\"AllowHide\" type=\"bool\">True</property>" +
                            "<property name=\"AllowMinimize\" type=\"bool\">True</property>" +
                            "<property name=\"AllowZoneChange\" type=\"bool\">True</property>" +
                            "<property name=\"ChromeType\" type=\"chrometype\">None</property>" +
                            "<property name=\"AllowConnect\" type=\"bool\">True</property>" +
                            "<property name=\"Width\" type=\"unit\" />" +
                            "<property name=\"Height\" type=\"unit\" />" +
                            "<property name=\"TitleUrl\" type=\"string\" />" +
                            "<property name=\"AllowEdit\" type=\"bool\">True</property>" +
                            "<property name=\"TitleIconImageUrl\" type=\"string\" />" +
                            "<property name=\"Direction\" type=\"direction\">NotSet</property>" +
                            "<property name=\"AllowClose\" type=\"bool\">True</property>" +
                            "<property name=\"ChromeState\" type=\"chromestate\">Normal</property>" +
                            "</properties>" +                            
                        "</data>" +
                        "</webPart>" +
                    "</webParts>";
            File oFile = ctx.Web.GetFileByServerRelativeUrl(url);
            LimitedWebPartManager limitedWebPartManager = oFile.GetLimitedWebPartManager(PersonalizationScope.Shared);

            WebPartDefinition oWebPartDefinition = limitedWebPartManager.ImportWebPart(xmlWebPart);
            limitedWebPartManager.AddWebPart(oWebPartDefinition.WebPart, "Left", 1);
            ctx.ExecuteQuery();
        }
        private static void UpdateWebpart(ClientContext ctx, string url, string title, string jspath)
        {
            //Script Editor
            File oFile = ctx.Web.GetFileByServerRelativeUrl(url);
            LimitedWebPartManager limitedWebPartManager = oFile.GetLimitedWebPartManager(PersonalizationScope.Shared);
            ctx.Load(limitedWebPartManager.WebParts,
               wps => wps.Include(
               wp => wp.WebPart.Title,
               wp => wp.WebPart.Properties));

            ctx.Load(limitedWebPartManager);
            ctx.ExecuteQuery();

            if (limitedWebPartManager.WebParts.Count == 0)
            {
                throw new Exception("No Web Parts on this page.");
            }
            foreach(WebPartDefinition wp in limitedWebPartManager.WebParts)
            {
                WebPart oWebPart = wp.WebPart;
                if (oWebPart.Title == title)
                {
                    //oWebPart.Title = "Script Editor";

                    oWebPart.Properties["Content"] = "<script type=\"text/javascript\" src=\"/Style%20Library/PMOCustom/js/TaskUpdate.js\"></script>";
                    wp.SaveWebPartChanges();
                    //oWebPart.Properties["Content"] = "<script type=\"text/javascript\"" + jspath +  "</script>";
                    ctx.ExecuteQuery();
                }

            }

            //Project Overview: content editor webpart
            /*foreach (WebPartDefinition wp in limitedWebPartManager.WebParts)
            {
                WebPart oWebPart = wp.WebPart;
                if (oWebPart.Title == "Project Overview")
                {
                    oWebPart.Title = "Project Overview 2";

                    //oWebPart.Properties["Content"] = "<script type=\"text/javascript\" src=\"/Style%20Library/PMOCustom/JSLink/IssuesDueDate.js\"></script>"; wp.SaveWebPartChanges();
                    ctx.ExecuteQuery();
                }

            }*/

        }
        /// <summary>
        /// Add script editor to page to all projects
        /// </summary>
        /// <param name="url"></param>
        private static void AddWebpartToAll(string url)
        {
            try
            {                
                using (ClientContext ctx = new ClientContext(url))
                {
                    //Get all sites under Dashboard site (recursvie)
                    Web oWeb = ctx.Web;
                    WebCollection Webs = oWeb.Webs;
                    ctx.Load(Webs);
                    ctx.ExecuteQuery();
                    var websFromLooping = new ArrayList();
                    GetListOfWebs(ctx, Webs, websFromLooping);

                    foreach (string web in websFromLooping)
                    {
                        ClientContext subContext = new ClientContext(web);
                        if (subContext != null)
                        {
                            Web oWebsite = subContext.Web;
                            subContext.Load(oWebsite);
                            subContext.ExecuteQuery();

                            List ProjectInfo = GetListByTitle(subContext, ConfigurationManager.AppSettings["projectstatement"].ToString());
                            if (ProjectInfo != null)
                            {
                                //string taskURL = oWebsite.ServerRelativeUrl + "/Lists/Tasks/" + "DispForm.aspx";
                                //string issueURL = oWebsite.ServerRelativeUrl + "/Lists/Issues/" + "DispForm.aspx";
                                //UpdateWebpart(subContext, issueURL);
                                //AddWebpartToPage(subContext, taskURL, ConfigurationManager.AppSettings["JSTask"].ToString());
                                //AddWebpartToPage(subContext, issueURL, ConfigurationManager.AppSettings["JSIssue"].ToString());
                                //DeleteWebpart(subContext, issueURL);

                                string taskURL = oWebsite.ServerRelativeUrl + "/Lists/Tasks/" + "EditForm.aspx";
                                AddWebpartToPage(subContext, taskURL, "Task Update", "/Style%20Library/PMOCustom/js/TaskUpdate.js");
                            }
                        }

                    }

                }

            }
            catch (Exception e)
            {
                Console.Write(e.Message + url);

            }
        }
        private static void UpdateWebpartToAll(string url)
        {
            try
            {
                using (ClientContext ctx = new ClientContext(url))
                {
                    //Get all sites under Dashboard site (recursvie)
                    Web oWeb = ctx.Web;
                    WebCollection Webs = oWeb.Webs;
                    ctx.Load(Webs);
                    ctx.ExecuteQuery();
                    var websFromLooping = new ArrayList();
                    GetListOfWebs(ctx, Webs, websFromLooping);

                    foreach (string web in websFromLooping)
                    {
                        ClientContext subContext = new ClientContext(web);
                        if (subContext != null)
                        {
                            Web oWebsite = subContext.Web;
                            subContext.Load(oWebsite);
                            subContext.ExecuteQuery();

                            List ProjectInfo = GetListByTitle(subContext, ConfigurationManager.AppSettings["projectstatement"].ToString());
                            if (ProjectInfo != null)
                            {
                                //string taskURL = oWebsite.ServerRelativeUrl + "/Lists/Tasks/" + "DispForm.aspx";
                                //http://devpmo.opt-osfns.org/DashBoard/SFPO/SFMPO/SFMProjectGovernance/default.aspx
                                //string homepageURL = oWebsite.ServerRelativeUrl + "/default.aspx";
                                //UpdateWebpart(subContext, homepageURL);
                                //AddWebpartToPage(subContext, taskURL, ConfigurationManager.AppSettings["JSTask"].ToString());
                                //AddWebpartToPage(subContext, issueURL, ConfigurationManager.AppSettings["JSIssue"].ToString());
                                //DeleteWebpart(subContext, issueURL);

                                string taskURL = oWebsite.ServerRelativeUrl + "/Lists/Tasks/" + "EditForm.aspx";
                                UpdateWebpart(subContext, taskURL, "Task Update", "Style%20Library/PMOCustom/js/TaskUpdate.js");
                            }
                        }

                    }

                }

            }
            catch (Exception e)
            {
                Console.Write(e.Message + url);

            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="url"></param>
        /// <param name="titleList"></param>
        /// <param name="titleView"></param>
        /// <param name="viewFields"></param>
        /// <param name="viewQuery"></param>
        /// <param name="jslink"></param>
        private static void UpdateView(string url, string titleList, string titleView, string[] viewFields, string viewQuery, string jslink = "" )
        {
            try
            {

                using (ClientContext clientContext = new ClientContext(url))
                {

                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    clientContext.ExecuteQuery();
                    List list = web.Lists.GetByTitle(titleList);
                    ViewCollection viewColl = list.Views;
                    clientContext.Load(viewColl);
                    clientContext.ExecuteQuery();
                    View taskView = GetViewByTitle(list, titleView);
                    
                    if (taskView != null)
                    {
                        taskView.ViewFields.RemoveAll();
                        foreach (string i in viewFields)
                        {
                            taskView.ViewFields.Add(i);
                            
                        }
                        taskView.ViewQuery = viewQuery;
                        if (jslink != "")
                            taskView.JSLink = jslink;
                        taskView.Update();
                        list.Update();
                       
                        clientContext.ExecuteQuery();
                    }
                    else
                    {
                        ViewCreationInformation creationInfo = new ViewCreationInformation();
                        creationInfo.Title = titleView;
                        creationInfo.RowLimit = 100;                        
                        creationInfo.ViewFields = viewFields;                        
                        creationInfo.Query = viewQuery;
                        creationInfo.ViewTypeKind = ViewType.None;
                        creationInfo.SetAsDefaultView = false;
                        
                        viewColl.Add(creationInfo);
                        
                        list.Update();
                        clientContext.ExecuteQuery();
                        taskView = GetViewByTitle(list, titleView);
                        if (jslink != "")
                            taskView.JSLink = jslink;
                        taskView.Update();
                        list.Update();
                        clientContext.ExecuteQuery();
                    }
                    //tasks
                    //string viewURL = web.ServerRelativeUrl + "/Lists/Tasks/" + titleView + ".aspx";
                    //string displayURL = web.ServerRelativeUrl + "/Lists/Tasks/DispForm.aspx";
                    //Issues
                    
                    if (jslink != "")
                    {
                        string viewURL = web.ServerRelativeUrl + "/Lists/Issues/" + titleView + ".aspx";
                        string displayURL = web.ServerRelativeUrl + "/Lists/Issues/DispForm.aspx";
                        RegisterJStoWebPart(web, viewURL, ConfigurationManager.AppSettings["JSFile"].ToString());
                        RegisterJStoWebPart(web, displayURL, ConfigurationManager.AppSettings["JSFile"].ToString());
                    }
                    
                    //RegisterJStoWebPart(web, "/DashBoard/PSALPO/PSALFastTrackPO/FinSystemsBackLog1/Lists/Tasks/AllItems.aspx", ConfigurationManager.AppSettings["JSFile"].ToString()); 
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message + url);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="list"></param>
        /// <param name="viewName"></param>
        /// <returns></returns>
        private static View GetViewByTitle(List list, string viewName)
        {
            View existingView = null;
            ViewCollection viewColl = list.Views;
            IEnumerable<View> existingViews = viewColl.Context.LoadQuery(viewColl.Where(view => view.Title == viewName));
            //IEnumerable<View> existingViews = ctx.LoadQuery(viewColl.Where(view => view.Title == viewName));
            viewColl.Context.ExecuteQuery();
            existingView = existingViews.FirstOrDefault();
            return existingView;
           
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="web"></param>
        /// <param name="url"></param>
        /// <param name="jsPath"></param>
        private static void RegisterJStoWebPart(Web web, string url, string jsPath)
        {
            Microsoft.SharePoint.Client.File newFormPageFile = web.GetFileByServerRelativeUrl(url);
            //LimitedWebPartManager limitedWebPartManager = newFormPageFile.GetLimitedWebPartManager(PersonalizationScope.Shared);
            LimitedWebPartManager limitedWebPartManager = newFormPageFile.GetLimitedWebPartManager(PersonalizationScope.Shared);
            web.Context.Load(limitedWebPartManager.WebParts);
            web.Context.ExecuteQuery();
            if (limitedWebPartManager.WebParts.Count > 0)
            {
                WebPartDefinition webPartDef = limitedWebPartManager.WebParts.FirstOrDefault();
                webPartDef.WebPart.Properties["JSLink"] = jsPath;
                webPartDef.SaveWebPartChanges();
                web.Context.ExecuteQuery();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="url"></param>
        private static void UpdateViewForAll(string url)
        {
            try
            {                
                using (ClientContext ctx = new ClientContext(url))
                {
                    //Get all sites under Dashboard site (recursvie)
                    Web oWeb = ctx.Web;
                    WebCollection Webs = oWeb.Webs;
                    ctx.Load(Webs);
                    ctx.ExecuteQuery();
                    var websFromLooping = new ArrayList();
                    GetListOfWebs(ctx, Webs, websFromLooping);
                    
                    foreach (string web in websFromLooping)
                    {
                        ClientContext subContext = new ClientContext(web);
                        if (subContext != null)
                        {
                            Web oWebsite = subContext.Web;
                            subContext.Load(oWebsite);
                            subContext.ExecuteQuery();

                            List ProjectInfo = GetListByTitle(subContext, ConfigurationManager.AppSettings["projectstatement"].ToString());
                            if (ProjectInfo != null)
                            {
                                 //tasks
                                 string tasksList = ConfigurationManager.AppSettings["tasks"].ToString();
                                 string alltasksView = ConfigurationManager.AppSettings["alltasksView"].ToString();
                                 string opentaskView = ConfigurationManager.AppSettings["opentaskView"].ToString();
                                 string latetaskView = ConfigurationManager.AppSettings["latetaskView"].ToString();
                                 string completedtaskView = ConfigurationManager.AppSettings["completedtaskView"].ToString();
                                 string mytaskView = ConfigurationManager.AppSettings["mytaskView"].ToString();
                                 // create custome View for Task list
                                 string[] taskviewFields = { "LinkTitle", "AssignedTo", "Status", "DueDate", "Current_x0020_Stage", "Exclude_x0020_from_x0020_Reports", "Workstreams" };
                                 string viewQueryOpenTasks = @"                                        
                                         <Where>                                                                                                                               
                                             <Or>
                                                 <Neq>
                                                     <FieldRef Name='Status' /><Value Type='Choice'>Completed</Value>
                                                 </Neq>
                                                 <Neq>
                                                     <FieldRef Name='PercentComplete' /><Value Type='Number'>1.00</Value>
                                                 </Neq>
                                             </Or>                                       
                                         </Where>                                       
                                         <OrderBy>
                                                 <FieldRef Name='Modified' Ascending='False' />
                                         </OrderBy>";
                                 string viewQueryLateTasks = @"                                        
                                         <Where>  
                                             <And>                                                                                   
                                             <Or>
                                                 <Neq>
                                                     <FieldRef Name='Status' /><Value Type='Choice'>Completed</Value>
                                                 </Neq>
                                                 <Neq>
                                                     <FieldRef Name='PercentComplete' /><Value Type='Number'>1.00</Value>
                                                 </Neq>
                                             </Or>
                                                 <Lt>
                                                     <FieldRef Name='DueDate' /><Value Type='DateTime' IncludeTimeValue='FALSE'><Today  Offset='0' /></Value>
                                                 </Lt>  
                                             </And>                                        
                                         </Where>                                       
                                         <OrderBy>
                                                 <FieldRef Name='DueDate' Ascending='False' />
                                                 <FieldRef Name='Priority' />
                                         </OrderBy>";
                                 string viewQueryCompletedTasks = @"                                        
                                         <Where>  
                                             <Or>
                                                 <Eq>
                                                     <FieldRef Name='Status' /><Value Type='Choice'>Completed</Value>
                                                 </Eq>
                                                 <Eq>
                                                     <FieldRef Name='PercentComplete' /><Value Type='Number'>1.00</Value>
                                                 </Eq>
                                             </Or>                                     
                                         </Where>                                       
                                         <OrderBy>
                                                 <FieldRef Name='Modified' Ascending='False' />                                                
                                         </OrderBy>";
                                 string viewQueryMyTasks = @"                                        
                                         <Where> 
                                                 <Eq>
                                                     <FieldRef Name='AssignedTo' /><Value Type='Integer'><UserID Type='Integer' /></Value>
                                                 </Eq>                                                                              
                                         </Where>                                       
                                         <OrderBy>
                                                 <FieldRef Name='Modified' Ascending='False' />                                                
                                         </OrderBy>";
                                 UpdateView(web, tasksList, alltasksView, taskviewFields, "", "");
                                 UpdateView(web, tasksList, opentaskView, taskviewFields, viewQueryOpenTasks, "");
                                 UpdateView(web, tasksList, latetaskView, taskviewFields, viewQueryLateTasks, "");
                                 UpdateView(web, tasksList, completedtaskView, taskviewFields, viewQueryCompletedTasks, "");
                                 UpdateView(web, tasksList, mytaskView, taskviewFields, viewQueryMyTasks, "");

                                 //Issues
                                 string issueList = ConfigurationManager.AppSettings["issues"].ToString();
                                 string allissuesView = ConfigurationManager.AppSettings["allissuesView"].ToString();
                                 string myissuesView = ConfigurationManager.AppSettings["myissuesView"].ToString();
                                 string activeissuesView = ConfigurationManager.AppSettings["activeissuesView"].ToString();
                                 //string overdueissuesView = ConfigurationManager.AppSettings["overdueissuesView"].ToString();
                                 // create custom view for Issues list
                                 string[] issueviewFields = { "LinkTitle", "AssignedTo", "Status", "Priority","DueDate", "Current_x0020_Stage", "Exclude_x0020_from_x0020_Reports", "Workstreams" };
                                 string viewQueryMyIssues = @"                                        
                                         <Where>                                                                                                                               
                                            <Eq>
                                                <FieldRef Name='AssignedTo' /><Value Type='Integer'><UserID Type='Integer' /></Value>
                                           </Eq>                                    
                                         </Where>                                       
                                         <OrderBy>
                                                 <FieldRef Name='Priority'  />
                                         </OrderBy>";
                                 string viewQueryActiveIssues = @"                                        
                                         <Where>  
                                             <Eq>
                                                  <FieldRef Name='Status' /><Value Type='Choice'>Active</Value>
                                             </Eq>                                       
                                         </Where>                                       
                                         <OrderBy>                                                
                                                 <FieldRef Name='Priority' />
                                         </OrderBy>";

                                 UpdateView(web, issueList, allissuesView, issueviewFields, "", "");
                                 UpdateView(web, issueList, myissuesView, issueviewFields, viewQueryMyIssues, "");
                                 UpdateView(web, issueList, activeissuesView, issueviewFields, viewQueryActiveIssues, "");

                                 //Risks
                                 string risksList = ConfigurationManager.AppSettings["risks"].ToString();
                                 string allrisksView = ConfigurationManager.AppSettings["allrisksView"].ToString();
                                 string myrisksView = ConfigurationManager.AppSettings["myrisksView"].ToString();
                                 string highrisksView = ConfigurationManager.AppSettings["highrisksView"].ToString();
                                 string byprojectphaseView = ConfigurationManager.AppSettings["byprojectphaseView"].ToString();

                                 // create custom view for Risks list
                                 string[] riskviewFields = { "LinkTitle", "AssignedTo", "Status", "Priority", "DueDate", "Current_x0020_Stage", "Exclude_x0020_from_x0020_Reports", "Workstreams" };
                                 string viewQueryAllRisks = @"     
                                         <OrderBy>
                                                 <FieldRef Name='Priority'  />
                                         </OrderBy>";
                                 string viewQueryMyRisks = @"   
                                         <Where>                                                                                                                               
                                            <Eq>
                                                <FieldRef Name='AssignedTo' /><Value Type='Integer'><UserID Type='Integer' /></Value>
                                           </Eq>                                    
                                         </Where>                                                                                                                   
                                         <OrderBy>
                                                 <FieldRef Name='Priority'  />
                                         </OrderBy>";
                                 string viewQueryHighRisks = @"                                                                                                                    
                                         <Where>  
                                             <Eq>
                                                  <FieldRef Name='Priority' /><Value Type='Choice'>High</Value>
                                             </Eq>                                       
                                         </Where>  
                                         <OrderBy>                                                
                                                 <FieldRef Name='Priority' />
                                         </OrderBy>";
                                 string viewQueryByProjectPhase = @"                                        
                                         <GroupBy Collapse = 'TRUE'><FieldRef Name='Current_x0020_Stage' /></GroupBy>
                                         <OrderBy>                                                
                                                 <FieldRef Name='Priority' />
                                         </OrderBy>";
                                 UpdateView(web, risksList, allrisksView, riskviewFields, viewQueryAllRisks, "");
                                 UpdateView(web, risksList, myrisksView, riskviewFields, viewQueryMyRisks, "");
                                 UpdateView(web, risksList, highrisksView, riskviewFields, viewQueryHighRisks, "");
                                 UpdateView(web, risksList, byprojectphaseView, riskviewFields, viewQueryByProjectPhase, "");
                                 
                                //Contacts list
                                /*string contactList = ConfigurationManager.AppSettings["contacts"].ToString();
                                string contactsView = ConfigurationManager.AppSettings["contactsView"].ToString();
                                // create custom view for Contacts list
                                string[] contactviewFields = { "Edit", "LinkTitle", "FullName", "JobTitle", "Company", "Email", "WorkPhone", "CellPhone" };
                                string viewQueryAllContact = @"     
                                        <OrderBy>
                                                <FieldRef Name='Title'  />
                                        </OrderBy>";                                
                                UpdateView(web, contactList, contactsView, contactviewFields, viewQueryAllContact, "");*/


                                //string jslink = ConfigurationManager.AppSettings["JSFile"].ToString();

                                //status Report
                                /*string titleList = ConfigurationManager.AppSettings["statusreport"].ToString();
                                string titleView = ConfigurationManager.AppSettings["taskView"].ToString();
                                string jslink = "";
                                string[] viewFields = { "Edit", "LinkTitle", "Current_x0020_Stage", "PercentComplete", "Health", "Budget", "Status_x0020_Start_x0020_Date", "Status_x0020_End_x0020_Date" };
                                string viewQuery = @"                                                                                                                   
                                        <OrderBy>
                                                <FieldRef Name='Status_x0020_Start_x0020_Date' Ascending='False' />
                                        </OrderBy>";*/

                                // Issues View Overdue
                                /*string overdueissuesView = ConfigurationManager.AppSettings["overdueissuesView"].ToString();
                                string issueList = ConfigurationManager.AppSettings["issues"].ToString();
                                string[] viewIssueFields = { "LinkTitle", "AssignedTo", "Status", "Priority", "DueDate", "Current_x0020_Stage", "Exclude_x0020_from_x0020_Reports", "Workstreams", "OverDueDays" };
                                string viewQueryIssue = @"                                        
                                        <Where>  
                                            <And>                                                                                   
                                                <Eq>
                                                    <FieldRef Name='Status' /><Value Type='Text'>Active</Value>
                                                </Eq>
                                                <Lt>
                                                    <FieldRef Name='DueDate' /><Value Type='DateTime' IncludeTimeValue='FALSE'><Today  Offset='0' /></Value>
                                                </Lt>  
                                            </And>                                        
                                        </Where>                                       
                                        <OrderBy>
                                                <FieldRef Name='Priority'  />
                                        </OrderBy>";
                                //Tasks View Overdue
                                string[] viewTaskFields = { "LinkTitle", "AssignedTo", "Status", "DueDate", "Current_x0020_Stage", "Exclude_x0020_from_x0020_Reports", "Workstreams", "OverDueDays" };
                                string tasksList = ConfigurationManager.AppSettings["tasks"].ToString();
                                string overduetaskView = ConfigurationManager.AppSettings["overduetaskView"].ToString();
                                string viewQueryTask = @"                                        
                                        <Where>  
                                            <And>                                                                                   
                                            <Or>
                                                <Neq>
                                                    <FieldRef Name='Status' /><Value Type='Choice'>Completed</Value>
                                                </Neq>
                                                <Neq>
                                                    <FieldRef Name='PercentComplete' /><Value Type='Number'>1.00</Value>
                                                </Neq>
                                            </Or>
                                                <Lt>
                                                    <FieldRef Name='DueDate' /><Value Type='DateTime' IncludeTimeValue='FALSE'><Today  Offset='0' /></Value>
                                                </Lt>  
                                            </And>                                        
                                        </Where>                                       
                                        <OrderBy>
                                                <FieldRef Name='Modified' Ascending='False' />
                                        </OrderBy>";
                                
                                
                                UpdateView(web, issueList, overdueissuesView, viewIssueFields, viewQueryIssue, "");
                                UpdateView(web, tasksList, overduetaskView, viewTaskFields, viewQueryTask, "");
                                //
                                */
                            }
                        }

                    }

                }

            }
            catch (Exception e)
            {
                Console.Write(e.Message + url);

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
