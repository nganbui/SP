using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using System.ServiceModel;
using System.Configuration;

namespace NYDOE.RER.PMO.SPGroupProvisionWeb.Services
{
    public class AppEventReceiver : IRemoteEventService
    {
        /// <summary>
        /// Handles app events that occur after the app is installed or upgraded, or when app is being uninstalled.
        /// </summary>
        /// <param name="properties">Holds information about the app event.</param>
        /// <returns>Holds information returned from the app event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            SPRemoteEventResult result = new SPRemoteEventResult();

            using (ClientContext clientContext = TokenHelper.CreateAppEventClientContext(properties, useAppWeb: false))
            {
                if (clientContext != null)
                {
                    clientContext.Load(clientContext.Web);
                    clientContext.ExecuteQuery();

                    string listTitle = ConfigurationManager.AppSettings["ReceiverList"].ToString();
                    string remoteEventReceiverSvcTitle = ConfigurationManager.AppSettings["ReceiverName"].ToString();
                    string remoteEventReceiverName = ConfigurationManager.AppSettings["ReceiverName"].ToString();
                    List ProjectInfo = GetListByTitle(clientContext, listTitle);

                    if (ProjectInfo != null)
                    {
                        clientContext.Load(ProjectInfo);
                        clientContext.ExecuteQuery();

                        if (properties.EventType == SPRemoteEventType.AppInstalled)
                        {

                            string opContext = OperationContext.Current.Channel.LocalAddress.Uri.AbsoluteUri.Substring(0, OperationContext.Current.Channel.LocalAddress.Uri.AbsoluteUri.LastIndexOf("/"));
                            string remoteEventReceiverSvcUrl = string.Format("{0}/{1}.svc", opContext, remoteEventReceiverSvcTitle);
                            RegisterEventReceiver(clientContext, ProjectInfo, remoteEventReceiverName, remoteEventReceiverSvcUrl, EventReceiverType.ItemAdded, 15010);
                            RegisterEventReceiver(clientContext, ProjectInfo, remoteEventReceiverName, remoteEventReceiverSvcUrl, EventReceiverType.ItemUpdated, 15011);
                        }
                        else if (properties.EventType == SPRemoteEventType.AppUninstalling)
                        {
                            UnregisterAllEventReceivers(clientContext, ProjectInfo, remoteEventReceiverName);
                        }
                    }

                }
            }

            return result;
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
        /// Register the remote event receiver
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="list"></param>
        /// <param name="name"></param>
        /// <param name="serviceUrl"></param>
        /// <param name="eventType"></param>
        /// <param name="sequence"></param>
        private void RegisterEventReceiver(ClientContext clientContext, List list, string name, string serviceUrl, EventReceiverType eventType, int sequence)
        {
            EventReceiverDefinitionCreationInformation newEventReceiver = new EventReceiverDefinitionCreationInformation()
            {
                EventType = eventType,
                ReceiverName = name,
                ReceiverUrl = serviceUrl,
                SequenceNumber = sequence
            };

            list.EventReceivers.Add(newEventReceiver);
            clientContext.ExecuteQuery();
        }

        /// <summary>
        /// Unregister the Remote Event Receiver
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="list"></param>
        /// <param name="name"></param>
        private void UnregisterAllEventReceivers(ClientContext clientContext, List list, string name)
        {
            EventReceiverDefinitionCollection erdc = list.EventReceivers;
            clientContext.Load(erdc);
            clientContext.ExecuteQuery();

            List<EventReceiverDefinition> toDelete = new List<EventReceiverDefinition>();
            foreach (EventReceiverDefinition erd in erdc)
            {
                if (erd.ReceiverName == name)
                {
                    toDelete.Add(erd);
                }
            }

            foreach (EventReceiverDefinition item in toDelete)
            {
                item.DeleteObject();
                clientContext.ExecuteQuery();
            }

        }
        /// <summary>
        /// This method is a required placeholder, but is not used by app events.
        /// </summary>
        /// <param name="properties">Unused.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            throw new NotImplementedException();
        }
    }
}
