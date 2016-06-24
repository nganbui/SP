using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using System.Configuration;

namespace NYDOE.RER.PMO.SPGroupProvisionWeb.Services
{
    public class RERProjectStatement : IRemoteEventService
    {
        /// <summary>
        /// Handles events that occur before an action occurs, such as when a user adds or deletes a list item.
        /// </summary>
        /// <param name="properties">Holds information about the remote event.</param>
        /// <returns>Holds information returned from the remote event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            SPRemoteEventResult result = new SPRemoteEventResult();

            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null)
                {
                    clientContext.Load(clientContext.Web);
                    clientContext.ExecuteQuery();
                }
            }

            return result;
        }

        /// <summary>
        /// Handles events that occur after an action occurs, such as after a user adds an item to a list or deletes an item from a list.
        /// </summary>
        /// <param name="properties">Holds information about the remote event.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            Uri sharepointUrl = new Uri(properties.ItemEventProperties.WebUrl);
            ClientContext clientContext = new ClientContext(sharepointUrl);

            //using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
           // {
                if (clientContext != null)
                {
                    Web oWebsite = clientContext.Web;
                    clientContext.Load(oWebsite);
                    clientContext.ExecuteQuery();

                    try
                    {
                        /*string rerListname = ConfigurationManager.AppSettings["RERListName"].ToString();
                        List DemoList = clientContext.Web.Lists.GetByTitle(rerListname);
                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem newItem = DemoList.AddItem(itemCreateInfo);*/

                        string groupOwners = oWebsite.Title + " " + ConfigurationManager.AppSettings["Owners"].ToString();
                        string groupMembers = oWebsite.Title + " " + ConfigurationManager.AppSettings["Members"].ToString();

                        GroupCollection collGroup = clientContext.Web.SiteGroups;
                        clientContext.Load(collGroup);
                        clientContext.ExecuteQuery();

                        List lstExternalEvents = clientContext.Web.Lists.GetByTitle(properties.ItemEventProperties.ListTitle);
                        ListItem itemEvent = lstExternalEvents.GetItemById(properties.ItemEventProperties.ListItemId);
                        clientContext.Load(itemEvent);
                        clientContext.ExecuteQuery();

                        // project manager   
                        FieldUserValue manager = itemEvent["Project_x0020_Manager"] as FieldUserValue;
                        List<FieldUserValue> uManagers = new List<FieldUserValue>();
                        if (manager != null)
                            uManagers.Add(manager);
                        // project members
                        List<FieldUserValue> uMembers = new List<FieldUserValue>();
                        foreach (FieldUserValue user in itemEvent["Project_x0020_Team"] as FieldUserValue[])
                        {
                            if (user != null)
                                uMembers.Add(user);
                        }
                        //
                        switch (properties.EventType)
                        {
                            case SPRemoteEventType.ItemAdded:
                                ///newItem["Title"] = "Updated by RER an item Project Statement " + itemEvent["Title"];
                                //newItem.Update();
                                //clientContext.ExecuteQuery();
                                // Create owner groups
                                CreateGroup(collGroup, groupOwners, oWebsite, clientContext, RoleType.Administrator, uManagers);
                                // Create contribute groups
                                CreateGroup(collGroup, groupMembers, oWebsite, clientContext, RoleType.Contributor, uMembers);
                                break;
                            case SPRemoteEventType.ItemUpdated:
                                //newItem["Title"] = "Updated by RER updated an item Project Statement : " + itemEvent["Title"] + " " + DateTime.Now.ToShortDateString();
                                //newItem.Update();
                                //clientContext.ExecuteQuery();
                                // Create owner groups
                                CreateGroup(collGroup, groupOwners, oWebsite, clientContext, RoleType.Administrator, uManagers);
                                // add project manager to this group
                                // Create contribute groups
                                CreateGroup(collGroup, groupMembers, oWebsite, clientContext, RoleType.Contributor, uMembers);
                                break;
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.Write(ex.Message);
                    }
                }
           // }
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
                    //clientContext.Load(grp);
                    //clientContext.ExecuteQuery();
                }
                // grant role to group
                RoleDefinitionBindingCollection collRoleDefinitionBinding = new RoleDefinitionBindingCollection(clientContext);
                RoleDefinition oRoleDefinition = oWebsite.RoleDefinitions.GetByType(roleType);
                collRoleDefinitionBinding.Add(oRoleDefinition);
                oWebsite.RoleAssignments.Add(grp, collRoleDefinitionBinding);
                clientContext.Load(grp, group => group.Title);
                clientContext.Load(oRoleDefinition, role => role.Name);
                clientContext.ExecuteQuery();

                // Add users to newly created group or existing group
                AddUsertoGroup(grp, clientContext, users);
            }
            catch (Exception e)
            {
                Console.Write(e.Message);

            }

        }

        /// <summary>
        /// Add users to existing group
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
                    // remove all users from existed group
                    UserCollection listofUsers = grp.Users;
                    clientContext.Load(listofUsers);
                    clientContext.ExecuteQuery();
                    if (listofUsers.Count > 0)
                    {
                        foreach (User existedUser in listofUsers)
                        {
                            if (existedUser != null)
                                listofUsers.RemoveById(existedUser.Id);

                        }

                        clientContext.ExecuteQuery();

                    }
                    // readd to this group
                    foreach (FieldUserValue member in users)
                    {
                        User user = null;
                        user = clientContext.Web.GetUserById(member.LookupId);
                        clientContext.Load(user);
                        clientContext.ExecuteQuery();
                        if (user != null)
                        {

                            UserCreationInformation userCreationInfo = new UserCreationInformation();
                            userCreationInfo.Email = user.Email;
                            userCreationInfo.LoginName = user.LoginName;
                            userCreationInfo.Title = user.Title;

                            User oUser = grp.Users.Add(userCreationInfo);

                            clientContext.ExecuteQuery();
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
