using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using System.Configuration;

namespace NYDOE.RER.PSALWeb.Services
{
    public class RERCalendar : IRemoteEventService
    {
        /// <summary>
        /// Handles events that occur before an action occurs, such as when a user adds or deletes a list item.
        /// </summary>
        /// <param name="properties">Holds information about the remote event.</param>
        /// <returns>Holds information returned from the remote event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            SPRemoteEventResult result = new SPRemoteEventResult();

            switch (properties.EventType)
            {
                case SPRemoteEventType.ItemUpdating:
                    result.ErrorMessage = "You cannot add this list item";
                    result.Status = SPRemoteEventServiceStatus.CancelNoError;
                    break;
                case SPRemoteEventType.ItemAdding:
                    result.ErrorMessage = "You cannot add this list item";
                    result.Status = SPRemoteEventServiceStatus.CancelNoError;
                    break;
            }

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
            
            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null)
                {
                    List lstExternalEvents = clientContext.Web.Lists.GetByTitle(properties.ItemEventProperties.ListTitle);
                    ListItem itemEvent = lstExternalEvents.GetItemById(properties.ItemEventProperties.ListItemId);
                    
                    
                    clientContext.Load(itemEvent);
                    clientContext.ExecuteQuery();

                    clientContext.Load(clientContext.Web);
                    clientContext.ExecuteQuery();

                    try
                    {
                        string rerListname = ConfigurationManager.AppSettings["RERListName"].ToString();
                        List DemoList = clientContext.Web.Lists.GetByTitle(rerListname);
                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem newItem = DemoList.AddItem(itemCreateInfo);
                        switch (properties.EventType)
                        {
                            case SPRemoteEventType.ItemAdded:
                                
                                newItem["Title"] = "Updated by RER added an item calendar : " + itemEvent["Title"];
                                newItem.Update();
                                clientContext.ExecuteQuery();
                                break;
                            case SPRemoteEventType.ItemUpdated:
                                newItem["Title"] = "Updated by RER updated an item calendar : " + itemEvent["Title"] + " " + DateTime.Now.ToShortDateString();
                                newItem.Update();
                                clientContext.ExecuteQuery();
                                break;
                        }
                        
                    }
                    catch(Exception ex)
                    {
                        Console.Write(ex.Message);
                    }
                }
            }
        }
    }
}
