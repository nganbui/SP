using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NYDOE.PMO.GroupProvisionTimerJob
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
                        if (subContext != null)
                        {
                            Web oWebsite = subContext.Web;
                            subContext.Load(oWebsite);
                            subContext.ExecuteQuery();

                            List ProjectInfo = GetListByTitle(subContext, ConfigurationManager.AppSettings["projectstatement"].ToString());
                            // run on projects not run on program
                            if (ProjectInfo != null && ProjectInfo.ItemCount > 0)
                            {
                                CamlQuery oQuery = new CamlQuery();
                                oQuery.ViewXml = @"<View><Query>" +
                                                 "<Where>" +
                                                 "<Eq><FieldRef Name='Project_x0020_Status'/><Value Type='Choice'>Active</Value></Eq>" +
                                                 "</Where>" +
                                                 "<OrderBy><FieldRef Name='ID' /></OrderBy></Query>" +
                                                "<ViewFields><FieldRef Name='Project_x0020_Manager' LookupId='TRUE' /><Value Type='User'></Value>" +
                                                "<FieldRef Name='Project_x0020_Team' LookupId='TRUE' /><Value Type='User'></Value></ViewFields>" +
                                                "<RowLimit>1</RowLimit>" +
                                                "</View>";
                                ListItemCollection items = ProjectInfo.GetItems(oQuery);
                                subContext.Load(items);
                                subContext.ExecuteQuery();
                                if (items.Count > 0)
                                {
                                    ListItem item = items.FirstOrDefault();
                                    subContext.Load(item);
                                    subContext.ExecuteQuery();
                                    // check last modified of Project Information
                                    DateTime lastModifiedDate = DateTime.Parse(item["Modified"].ToString());
                                    //DateTime.TryParse(item["Modified"].ToString(), out lastModifiedDate);
                                    DateTime today = DateTime.Today;
                                    TimeSpan dif = today - lastModifiedDate;
                                    int difference = DateTime.Compare(today, lastModifiedDate);
                                    // just run if project information has been updated within one day
                                    if (dif.Days==0)
                                    {                                        
                                        // project manager   
                                        FieldUserValue manager = item["Project_x0020_Manager"]!=null? item["Project_x0020_Manager"] as FieldUserValue : null;
                                        List<FieldUserValue> uManagers = new List<FieldUserValue>();
                                        if (manager != null)
                                            uManagers.Add(manager);

                                        // project members
                                        FieldUserValue[] members = item["Project_x0020_Team"]!=null? item["Project_x0020_Team"] as FieldUserValue[]: null;
                                        List<FieldUserValue> uMembers = new List<FieldUserValue>();
                                        if (members != null)
                                        {
                                            foreach (FieldUserValue user in members as FieldUserValue[])
                                            {
                                                if (user != null)
                                                    uMembers.Add(user);
                                            }
                                        }
                                        if (uManagers.Count > 0 || uMembers.Count > 0)
                                        {                                            
                                            // add users to corresponding sharepoint group
                                            string groupOwners = oWebsite.Title + " " + ConfigurationManager.AppSettings["Owners"].ToString();
                                            string groupMembers = oWebsite.Title + " " + ConfigurationManager.AppSettings["Members"].ToString();

                                            GroupCollection collGroup = oWebsite.SiteGroups;
                                            subContext.Load(collGroup);
                                            subContext.ExecuteQuery();

                                            // remove users from Contact list and sharepoint group
                                            RemoveUsers(collGroup, groupOwners, subContext);
                                            RemoveUsers(collGroup, groupMembers, subContext);
                                            // add users to Contact List and sharepoint group
                                            // Create owner group
                                            CreateGroup(collGroup, groupOwners, oWebsite, subContext, RoleType.Administrator, uManagers);
                                            // Create contribute group
                                            CreateGroup(collGroup, groupMembers, oWebsite, subContext, RoleType.Contributor, uMembers);

                                        }
                                    }
                                }
                            }
                        }

                    }

                }

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
        /// <summary>
        /// Create group if it's not existed
        /// </summary>
        /// <param name="collGroup"></param>
        /// <param name="groupName"></param>
        /// <param name="oWebsite"></param>
        /// <param name="clientContext"></param>
        /// <param name="roleType"></param>
        /// <param name="users"></param>
        private static void CreateGroup(GroupCollection collGroup, string groupName, Web oWebsite, ClientContext clientContext, RoleType roleType, List<FieldUserValue> users)
        {
            try
            {
                Group grp = collGroup.Where(g => g.Title == groupName).FirstOrDefault();
                oWebsite.BreakRoleInheritance(true, false);
                if (grp == null)
                {
                    GroupCreationInformation groupCreationInfo = new GroupCreationInformation();
                    groupCreationInfo.Title = groupName;
                    groupCreationInfo.Description = "Use this group to grant people " + roleType.ToString() + " permissions to the SharePoint site: " + oWebsite.Title;
                    grp = oWebsite.SiteGroups.Add(groupCreationInfo);
                    clientContext.ExecuteQuery();
                }
                // grant role to group
                RoleDefinitionBindingCollection collRoleDefinitionBinding = new RoleDefinitionBindingCollection(clientContext);
                RoleDefinition oRoleDefinition = oWebsite.RoleDefinitions.GetByType(roleType);
                collRoleDefinitionBinding.Add(oRoleDefinition);
                oWebsite.RoleAssignments.Add(grp, collRoleDefinitionBinding);
                clientContext.Load(grp, group => group.Title);
                clientContext.Load(oRoleDefinition, role => role.Name);
                clientContext.ExecuteQuery();

                // Add users to newly created group or existed group
                AddUsertoGroup(grp, clientContext, users);
            }
            catch (Exception e)
            {
                Console.Write(e.Message);

            }

        }
        /// <summary>
        /// Add users to existed group
        /// </summary>
        /// <param name="grp"></param>
        /// <param name="clientContext"></param>
        /// <param name="users"></param>
        private static void AddUsertoGroup(Group grp, ClientContext clientContext, List<FieldUserValue> users)
        {
            try
            {
                if (grp != null)
                {    
                    // add user to this group
                    foreach (FieldUserValue member in users)
                    {
                        User user = null;
                        user = clientContext.Web.GetUserById(member.LookupId);
                        clientContext.Load(user);
                        clientContext.ExecuteQuery();
                        if (user != null)
                        {
                            // add user to corresponding group
                            UserCreationInformation userCreationInfo = new UserCreationInformation();
                            userCreationInfo.Email = user.Email;
                            userCreationInfo.LoginName = user.LoginName;
                            userCreationInfo.Title = user.Title;

                            User oUser = grp.Users.Add(userCreationInfo);
                            clientContext.ExecuteQuery();

                            // add user to Contact list
                            List contactlist = clientContext.Web.Lists.GetByTitle(ConfigurationManager.AppSettings["contactlist"].ToString());
                            if (contactlist != null)
                            {
                                var loginname = user.LoginName;
                                // get user Information from User Profile
                                PersonProperties userProfile = GetUserInformation(clientContext, loginname);
                                // in case not able to get user information from User Profile -- > get user Info from SP Information list (userCreationInfo) 
                                var accountName = userProfile.IsPropertyAvailable("AccountName")==true ? userProfile.UserProfileProperties["AccountName"].ToString() : user.LoginName;                                
                                ListItemCollection items = null;
                                bool isExist = ItemExists(contactlist, accountName, out items);
                                if (!isExist)
                                {
                                    var lastname = user.Title.IndexOf(" ") > 0 ? user.Title.Substring(user.Title.IndexOf(" ")) : user.Title;
                                    ListItemCreationInformation contactInfo = new ListItemCreationInformation();
                                    ListItem newItem = contactlist.AddItem(contactInfo);
                                    newItem["AccountName"] = accountName;
                                    newItem["Title"] = userProfile.IsPropertyAvailable("AccountName") == true ? userProfile.UserProfileProperties["LastName"].ToString() : lastname;
                                    newItem["FullName"] = userProfile.IsPropertyAvailable("AccountName") == true ? userProfile.UserProfileProperties["PreferredName"].ToString() : user.Title;
                                    newItem["Email"] = userProfile.IsPropertyAvailable("AccountName") == true ? user.Email : user.Email;
                                    newItem["Company"] = userProfile.IsPropertyAvailable("AccountName") == true ? userProfile.UserProfileProperties["Department"].ToString() : string.Empty;
                                    newItem["JobTitle"] = userProfile.IsPropertyAvailable("AccountName") == true ? userProfile.Title : string.Empty;
                                    newItem["WorkPhone"] = userProfile.IsPropertyAvailable("AccountName") == true ? userProfile.UserProfileProperties["WorkPhone"].ToString() : string.Empty;
                                    newItem["WorkFax"] = userProfile.IsPropertyAvailable("AccountName") == true ? userProfile.UserProfileProperties["Fax"].ToString() : string.Empty;
                                    newItem["WorkAddress"] = userProfile.IsPropertyAvailable("AccountName") == true ? userProfile.UserProfileProperties["SPS-Location"].ToString() : string.Empty;
                                    newItem["CellPhone"] = userProfile.IsPropertyAvailable("AccountName") == true  ? userProfile.UserProfileProperties["CellPhone"].ToString() : string.Empty;
                                    newItem["WebPage"] = userProfile.IsPropertyAvailable("AccountName") == true ? userProfile.UserUrl : string.Empty;
                                    newItem.Update();
                                    clientContext.ExecuteQuery();
                                }
                            }

                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.Write(e.Message);

            }

        }
        /// <summary>
        /// Get user profile to add to contact list
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="accountName"></param>
        private static PersonProperties GetUserInformation(ClientContext ctx, string accountName)
        {           
            PersonProperties personProperties = null;
            
            try
            {
                PeopleManager peopleManager = new PeopleManager(ctx);
                personProperties = peopleManager.GetPropertiesFor(accountName);                
                ctx.Load(personProperties, p=>p.AccountName, p => p.Email, p => p.PictureUrl, p => p.Title, p => p.UserUrl, p => p.UserProfileProperties);
                ctx.ExecuteQuery();
                
            }
            catch(Exception e)
            {                
                Console.Write(e.Message);
            }
            return personProperties;

        }       
        /// <summary>
        /// Check if user in Contact list
        /// </summary>
        /// <param name="list"></param>
        /// <param name="accountName"></param>
        /// <returns></returns>
        private static bool ItemExists(List list, string accountName, out ListItemCollection items)
        {
            var ctx = list.Context;
            var query = new CamlQuery();
            query.ViewXml = @"<View>
                             <Query>
                             <Where>
                             <Eq>
                             <FieldRef Name='AccountName' />
                             <Value Type='Text'>" + accountName + @"</Value>
                             </Eq>
                             </Where>
                             </Query>
                             </View>";            
            items = list.GetItems(query);
            ctx.Load(items);
            ctx.ExecuteQuery();
            return items.Count > 0;
        }
        /// <summary>
        /// Remove users from Contact list and Sharepoint Group
        /// </summary>
        /// <param name="collGroup"></param>
        /// <param name="groupName"></param>
        /// <param name="clientContext"></param>
       
        private static void RemoveUsers(GroupCollection collGroup, string groupName, ClientContext clientContext)
        {
            try
            {
                Group grp = collGroup.Where(g => g.Title == groupName).FirstOrDefault();
                if (grp != null)
                {
                    UserCollection listofUsers = grp.Users;
                    clientContext.Load(listofUsers);
                    clientContext.ExecuteQuery();
                    if (listofUsers.Count > 0)
                    {
                        //get contact list
                        List contactlist = GetListByTitle(clientContext, ConfigurationManager.AppSettings["contactlist"].ToString());                       

                        foreach (User existedUser in listofUsers)
                        {
                            if (existedUser != null)
                            {
                                // remove existedUser from contact list if found                              
                                var existedAccountName = existedUser.LoginName.IndexOf("|") > 0 ? existedUser.LoginName.Substring(existedUser.LoginName.IndexOf("|") + 1) : existedUser.LoginName;
                                existedAccountName = existedAccountName.ToLower();                                
                                ListItemCollection items = null;
                                bool isExist = ItemExists(contactlist, existedAccountName, out items);
                                if (isExist)
                                {
                                    foreach (ListItem item in items)
                                    {
                                        item.DeleteObject();
                                        clientContext.ExecuteQuery();
                                    }

                                }
                                // remove existedUser from sharepoint group 'groupName'
                                listofUsers.RemoveById(existedUser.Id);
                                clientContext.ExecuteQuery();
                            }                            
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.Write(e.Message);

            }

        }
    }
}
