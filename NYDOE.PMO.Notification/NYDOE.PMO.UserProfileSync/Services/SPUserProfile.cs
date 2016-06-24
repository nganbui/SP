using System;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using NYDOE.PMO.UserProfileSync.Entities;


namespace NYDOE.PMO.UserProfileSync.Services
{
    public class SPUserProfile
    {
        #region Private Fields
        /// <summary>
        /// Domain name of the AD Store.
        /// </summary>
        private string _domain = string.Empty;

        private string _forest = string.Empty;

        private string _url = string.Empty;
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new AD Adapter to the specified doamin.
        /// </summary>
        /// <param name="domain"></param>
        public SPUserProfile(string url)
        {
            _url = url;
        }
        public SPUserProfile(string forest, string domain)
        {
            _forest = forest;
            _domain = domain;
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Gets all user profiles from one site collection.
        /// </summary>
        /// <returns></returns>
        public IEnumerable<NYDOEUser> GetUserProfiles()
        {
            return UserProfilesfromAD();
        }
        #endregion

        #region Protected Methods
        public List<NYDOEUser> UserProfiles()
        {
            List<NYDOEUser> userProfiles = new List<NYDOEUser>();
            try
            {                
                using (var context = new ClientContext(_url))
                {
                    UserCollection users = context.Web.SiteUsers;
                    context.Load(users);
                    context.ExecuteQuery();
                    foreach (User usr in users)
                    {
                        var loginname = usr.LoginName;
                        // get user Information from User Profile
                        PersonProperties userProfile = GetUserInformation(context, loginname);
                        // in case not able to get user information from User Profile -- > get user Info from SP Information list (userCreationInfo) 
                        var accountName = userProfile.IsPropertyAvailable("AccountName") == true ? userProfile.UserProfileProperties["AccountName"].ToString() : usr.LoginName;
                        var lastname = usr.Title.IndexOf(" ") > 0 ? usr.Title.Substring(usr.Title.IndexOf(" ")) : usr.Title;
                        var firstname = usr.Title.IndexOf(" ") > 0 ? usr.Title.Substring(0,usr.Title.IndexOf(" ")) : usr.Title;

                        var up = new NYDOEUser()
                        {                           
                            UID = usr.Id,
                            LoginName = accountName,
                            Email = userProfile.IsPropertyAvailable("WorkEmail") == true ? userProfile.UserProfileProperties["WorkEmail"].ToString() : usr.Email,
                            FirstName = userProfile.IsPropertyAvailable("FirstName") == true ? userProfile.UserProfileProperties["FirstName"].ToString() : firstname,
                            LastName = userProfile.IsPropertyAvailable("LastName") == true ? userProfile.UserProfileProperties["LastName"].ToString() : lastname,
                            DisplayName = userProfile.IsPropertyAvailable("PreferredName") == true ? userProfile.UserProfileProperties["PreferredName"].ToString() : string.Empty,
                            JobTitle = userProfile.IsPropertyAvailable("Title") == true ? userProfile.UserProfileProperties["Title"].ToString() : string.Empty,
                            Workphone = userProfile.IsPropertyAvailable("AccountName") == true ? userProfile.UserProfileProperties["WorkPhone"].ToString() : string.Empty,
                            Department = userProfile.IsPropertyAvailable("AccountName") == true ? userProfile.UserProfileProperties["Department"].ToString() : string.Empty

                        };
                        userProfiles.Add(up);

                    }
                }                
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return userProfiles;

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
                ctx.Load(personProperties, p => p.AccountName, p => p.Email, p => p.PictureUrl, p => p.Title, p => p.UserUrl, p => p.UserProfileProperties);
                ctx.ExecuteQuery();

            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
            return personProperties;

        }
        /// <summary>
        /// When overridden, retrieves a collection of user profiles from
        /// the AD domain.
        /// </summary>
        /// <returns></returns>
        protected virtual IEnumerable<NYDOEUser> UserProfilesfromAD()
        {
            using (var context = new PrincipalContext(ContextType.Domain, _forest, _domain))
            {
                using (var searcher = new PrincipalSearcher(new UserPrincipal(context)))
                {
                    var managers = new Dictionary<string, string>();
                    foreach (var result in searcher.FindAll())
                    {
                        DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry;
                        var fName = (de.Properties["givenName"] != null && de.Properties["givenName"].Value != null)
                            ? de.Properties["givenName"].Value.ToString() : string.Empty;

                        var lName = (de.Properties["sn"] != null && de.Properties["sn"].Value != null)
                            ? de.Properties["sn"].Value.ToString() : string.Empty;

                        var dName = result.DisplayName ?? string.Empty;
                        var logName = result.SamAccountName ?? string.Empty;
                        var em = de.Properties["EmailAddress"] != null && de.Properties["EmailAddress"].Value != null
                            ? de.Properties["EmailAddress"].Value.ToString() : string.Empty;

                        var mgrSid = string.Empty;
                        

                        var distName = result.DistinguishedName;

                        var up = new NYDOEUser()
                        {
                            FirstName = fName,
                            LastName = lName,
                            DisplayName = dName,
                            LoginName = logName,
                            Email = em,
                           // SID = result.Sid.ToString(),
                            Domain = distName
                            
                        };
                        //up.Memberships = result.GetGroups().Select(g => new BPGroup() { Id = g.Sid.ToString(), Name = g.Name, SourceSystem = SourceSystems.AD.ToString() }).ToList();
                        yield return up;
                    }
                }
            }
        }
        #endregion
    }
}
