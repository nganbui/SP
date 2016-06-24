using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using System.Configuration;
using System.ServiceModel;
using System.Collections;

namespace NYDOE.RER.PMO.GroupProvisionWeb.Services
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
                            subContext.Load(subContext.Web);
                            subContext.ExecuteQuery();
                            string listTitle = ConfigurationManager.AppSettings["ReceiverList"].ToString();
                            string remoteEventReceiverSvcTitle = ConfigurationManager.AppSettings["ReceiverName"].ToString();
                            string remoteEventReceiverName = ConfigurationManager.AppSettings["ReceiverName"].ToString();                            
                            List ProjectInfo = GetListByTitle(subContext, listTitle);
                            if (ProjectInfo != null)
                            {
                                 subContext.Load(ProjectInfo);
                                 subContext.ExecuteQuery();

                                if (properties.EventType == SPRemoteEventType.AppInstalled)
                                {

                                    string opContext = OperationContext.Current.Channel.LocalAddress.Uri.AbsoluteUri.Substring(0, OperationContext.Current.Channel.LocalAddress.Uri.AbsoluteUri.LastIndexOf("/"));
                                    string remoteEventReceiverSvcUrl = string.Format("{0}/{1}.svc", opContext, remoteEventReceiverSvcTitle);
                                    RegisterEventReceiver(subContext, ProjectInfo, remoteEventReceiverName, remoteEventReceiverSvcUrl, EventReceiverType.ItemAdded, 15010);
                                    RegisterEventReceiver(subContext, ProjectInfo, remoteEventReceiverName, remoteEventReceiverSvcUrl, EventReceiverType.ItemUpdated, 15011);
                                }
                                else if (properties.EventType == SPRemoteEventType.AppUninstalling)
                                {
                                    UnregisterAllEventReceivers(subContext, ProjectInfo, remoteEventReceiverName);
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
            
            return result;
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
