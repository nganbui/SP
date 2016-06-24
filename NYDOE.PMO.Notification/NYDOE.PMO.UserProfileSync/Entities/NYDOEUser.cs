using System;

namespace NYDOE.PMO.UserProfileSync.Entities
{
    public class NYDOEUser
    {
        public int UID { get; set; }        
        /// <summary>
        /// User's first name.
        /// </summary>
        public string FirstName { get; set; }

        /// <summary>
        /// User's last name.
        /// </summary>
        public string LastName { get; set; }

        /// <summary>
        /// User's preferred name.
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Title
        /// </summary>
        public string JobTitle { get; set; }

        /// <summary>
        /// User's login name (without domain prefix).
        /// </summary>
        public string LoginName { get; set; }

        /// <summary>
        /// User's email address
        /// </summary>
        public string Email { get; set; }
        /// <summary>
        /// User's workphone
        /// </summary>
        public string Workphone { get; set; }
        /// <summary>
        /// User's department
        /// </summary>
        public string Department { get; set; }
        /// <summary>
        /// The AD domain the user belongs to.
        /// </summary>
        public string Domain { get; set; }
    }
}
