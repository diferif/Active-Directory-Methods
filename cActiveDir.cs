using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Collections;
using System.DirectoryServices.ActiveDirectory;
using System.Reflection;

namespace AD.Classes
{
    public class KeyValue
    {
        public string key { get; set; }
        public string value { get; set; }
    }
    public static class cActiveDir
    {
        //set these two==================================================================================
        static string domainGet()  {  return "ACME";  }
        static string emailGet() { return "@acme.com"; }

        public static void adGetUserBySAMAccountName(string sAMAccountName, ref Dictionary<string,string> dictUser)
        {
            // set up domain context
            using (PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domainGet()))
            {
                UserPrincipal upFound = UserPrincipal.FindByIdentity(ctx, sAMAccountName);
                if (upFound != null) {
                    if (upFound.IsAccountLockedOut()) { dictUser.Add("Locked?", "YES"); }
                    else { dictUser.Add("Locked?", "No"); }
                    foreach (var prop in upFound.GetType().GetProperties()) {
                        try
                        {
                            if (prop.GetValue(upFound, null) != null) { dictUser.Add(prop.Name, prop.GetValue(upFound, null).ToString());  }
                            else { dictUser.Add(prop.Name, ""); }
                        }
                        catch (Exception)
                        {
                           //do nothing...keep going
                            // string sErrorMessage = ex.Message.ToString() + "<hr>";
                            //if (ex.InnerException != null) { sErrorMessage += "<hr>Inner Exception: " + ex.InnerException.ToString(); }
                        }
                    }
                }
            }
            //the properties:
            //AccountExpirationDate = 8 / 4 / 2019 3:59:59 AM
            //AccountLockoutTime =
            //AdvancedSearchFilter = System.DirectoryServices.AccountManagement.AdvancedFilters
            //AllowReversiblePasswordEncryption = False
            //BadLogonCount = 0
            //Certificates = System.Security.Cryptography.X509Certificates.X509Certificate2Collection
            //Context = System.DirectoryServices.AccountManagement.PrincipalContext
            //ContextType = Domain
            //Current = THOMAS, SCOTT
            //DelegationPermitted = True
            //Description =   OPERATIONS - - EMPLOYEE
            //DisplayName = THOMAS, SCOTT
            //DistinguishedName = CN = THOMAS\\, SCOTT, OU = People, OU = Accounts, DC = CORP, DC = ACME, DC = com
            //EmailAddress = SCOTT.THOMAS@acme.com
            //EmployeeId =
            //Enabled = True
            //GivenName = SCOTT
            //Guid = da47f368 - 3952 - 4b78 - bb8c - 7b048cecfdf5
            //HomeDirectory =
            //HomeDrive =
            //LastBadPasswordAttempt = 8 / 17 / 2016 12:59:42 PM
            //LastLogon = 11 / 22 / 2016 12:28:40 PM
            //LastPasswordSet = 11 / 14 / 2016 1:25:47 PM
            //MiddleName =
            //Name = THOMAS, SCOTT
            //PasswordNeverExpires = False
            //PasswordNotRequired = False
            //PermittedLogonTimes =
            //PermittedWorkstations = System.DirectoryServices.AccountManagement.PrincipalValueCollection`1[System.String]
            //SamAccountName = SM9970
            //ScriptPath = 0000.bat
            //Sid = S - 1 - 5 - 21 - 2130729834 - 1480738125 - 1508530778 - 356888
            //SmartcardLogonRequired = False
            //StructuralObjectClass = user
            //Surname = MURRAY
            //UserCannotChangePassword = False
            //UserPrincipalName = SCOTT.THOMAS@ACME.COM
            //VoiceTelephoneNumber = (555)555 - 9959
        }

        public static UserPrincipal adGetUserPrincipalBySAMAccountName(string sAMAccountName)
        {
            // set up domain context
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domainGet());
            UserPrincipal upFound = UserPrincipal.FindByIdentity(ctx, sAMAccountName);
            if (upFound != null) return upFound;
            return null;
        }

        public static string adGetUserPrincipalMethodsHtml()
        {
            string sR = "UserPrincipal Methods:<br> ";

            foreach (System.Reflection.MethodInfo mi in typeof(System.DirectoryServices.AccountManagement.UserPrincipal).GetMethods())
            {
                var parameterDescriptions = string.Join
                (", ", mi.GetParameters()
                             .Select(x => x.ParameterType + " " + x.Name)
                             .ToArray());
                sR += mi.ReturnType + mi.Name + parameterDescriptions + "<br>";
            }
            return sR;
        }

        //public static UserPrincipal adGetUserPrincipalByName(string sName)
        //{
        //    // set up domain context
        //    PrincipalContext ctx = new PrincipalContext(ContextType.Domain, sDomain);
        //    UserPrincipal upFound = UserPrincipal.FindByIdentity(ctx, sName);
        //    if (upFound != null) return upFound;
        //    return null;
        //}

        public static string adGetEmailAddressByFullName(string sFullName)
        {
            // set up domain context

            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domainGet());
            UserPrincipal upFound = UserPrincipal.FindByIdentity(ctx, "MURRAY, SCOTT");
            if (upFound == null) return "";
            return upFound.EmailAddress;
        }

        public static string adGetUserPropertyListHtml(string sUserId)
        {
            System.DirectoryServices.AccountManagement.UserPrincipal ctx = adGetUserPrincipalBySAMAccountName(sUserId);
            if (ctx == null) return "";
            string sR = "Locked = ";
            if (ctx.IsAccountLockedOut()) { sR += "<font color=red>****YES***</font><br>"; }
            else { sR += "Locked = NO<br>"; }
            foreach (var prop in ctx.GetType().GetProperties()) { sR += prop.Name + " = " + prop.GetValue(ctx, null) + "<br>"; }
            return sR;
        }

        
        public static string adUserInfo(string sUserId)
        {
            Dictionary<string, string> dictUser = new Dictionary<string, string>();
            adGetUserBySAMAccountName(sUserId, ref dictUser);
            return new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(dictUser).ToString();
        }

        public static string adUserInfoByEmail(string sEmailWithoutOrgDotCom)
        {

            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domainGet());
            UserPrincipal user = new UserPrincipal(ctx);
            user.EmailAddress = sEmailWithoutOrgDotCom + emailGet();
            // create a principal searcher for running a search operation
            PrincipalSearcher pS = new PrincipalSearcher(user);
            PrincipalSearchResult<Principal> results = pS.FindAll();
            string sReturn = "";
            foreach (Principal result in results)
            {
                sReturn = result.SamAccountName;
                break;
            }
            return adUserInfo(sReturn);
        }

　
        public static string adALLUserGroupsReturnListJson()
        {
            //get list groups in List
            List<KeyValue> lstGroups = new List<KeyValue>();
            PrincipalContext yourOU = new PrincipalContext(ContextType.Domain, domainGet());
            GroupPrincipal findAllGroups = new GroupPrincipal(yourOU, "*");
            findAllGroups.Description = "*VCS*";
            PrincipalSearcher ps = new PrincipalSearcher(findAllGroups);
            foreach (var group in ps.FindAll().OrderBy(x => x.Name))
            {
                lstGroups.Add(new KeyValue { key = "group", value = group.Name });
            }
            return new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(lstGroups).ToString();
        }

        public static bool adIsUserInAdGroupsBool(string sUserId, string sGroupOrUserListSemiColon)
        {
            //check if user has SAM acct
            System.DirectoryServices.AccountManagement.UserPrincipal ctx = adGetUserPrincipalBySAMAccountName(sUserId);
            if (ctx == null) return false;

            //check if sUserId is in sGroupOrUserListSemiColon
            string[] lstGroupsAndUsers = sGroupOrUserListSemiColon.Split(new char[] { ';' });
            for ( int i = 0; i < lstGroupsAndUsers.Length; i++)  {    if (sUserId.IndexOf(lstGroupsAndUsers[i]) > -1) return true;     }

            //check if any groups is in sGroup 
            foreach (GroupPrincipal group in ctx.GetGroups().OrderBy(x => x.Name)) {
                for (int i = 0; i < lstGroupsAndUsers.Length; i++) { if (group.Name == lstGroupsAndUsers[i]  ) return true;  }
            }
            return false;
        }

        public static string returnJsonListOfAdGroupsForUserId(string sUserId)
        {
            //get list groups in List
            List<KeyValue> lstGroups = new List<KeyValue>();
            System.DirectoryServices.AccountManagement.UserPrincipal ctx = adGetUserPrincipalBySAMAccountName(sUserId);
            if (ctx == null) return "";
            foreach (GroupPrincipal group in ctx.GetGroups().OrderBy(x => x.Name)) { lstGroups.Add(new KeyValue { key = "group", value = group.Name } );   }
            return new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(lstGroups).ToString();
        }

        public static string adGroupGetAllUsers(string sGroupName)
        {
            List<KeyValue> lstGroups = new List<KeyValue>();
            using (var context = new PrincipalContext(ContextType.Domain, domainGet()))
            {
                using (var group = GroupPrincipal.FindByIdentity(context, sGroupName))
                {
                    if (group == null) {    return "";  }
                    else {
                        var users = group.GetMembers(true);
                        foreach (UserPrincipal user in users.OrderBy(x=>x.Name)  ) {    lstGroups.Add(new KeyValue { key = "user", value = user.Name + "  ...  " + user.SamAccountName });  }
                    }
                    //return new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(lstGroups).ToString();
                    return listKeyValueSerialize(lstGroups);
                }
            }
        }//function

        public static string listKeyValueSerialize(List<KeyValue> lst)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            foreach (KeyValue item in lst)
            {
                if (sb.Length > 0) sb.Append(",");
                if (sb.Length == 0) sb.Append("[");
                sb.Append("{\"key\":\"" + item.key + "\",\"value\":\"" + item.value + "\"}");
            }
            if (sb.Length > 0) sb.Append("]");
            return sb.ToString();
        }

        public static string adGroupGetAllUsersIdInSqlWithQuotes(string sGroupName)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            using (var context = new PrincipalContext(ContextType.Domain, domainGet()))
            {
                using (var group = GroupPrincipal.FindByIdentity(context, sGroupName))
                {
                    if (group == null) { return ""; }
                    else
                    {
                        var users = group.GetMembers(true);
                        foreach (UserPrincipal user in users.OrderBy(x => x.Name)) {
                            if (sb.Length > 0) sb.Append(",");
                            sb.Append("'" + user.SamAccountName + "'");                        }
                    }
                }//using
            }
            return sb.ToString();
        }//function

        //public static string adGetAllUsersJson(string sUserId)
        //{

        //    //get the group name from db
        //    string sAdUserGroup = "";
        //    using (var db = new Classes.cDatabase()) { 
        //      db.getString("select sAdminConfigValue from tAdminConfig where sAdminConfigName = 'sAdUserGroup'", ref sAdUserGroup); 
        //    }

        //    return adGroupGetAllUsers(sAdUserGroup);
        //}

        private static string GetCurrentDomainPath()
        {
            DirectoryEntry de = new DirectoryEntry("LDAP://RootDSE");
            string ss =  "LDAP://" + de.Properties["defaultNamingContext"][0].ToString();
            return ss;
        }

　
        public static string adGetUserWhenChanged(string sUserId)
        {
            DirectoryEntry searchRoot = new DirectoryEntry(GetCurrentDomainPath(), null, null, AuthenticationTypes.Secure );
            string sReturn = "";
            using (searchRoot)
            {
                DirectorySearcher ds = new DirectorySearcher(searchRoot, "(sAMAccountName=" + sUserId + ")" );
                sReturn = ds.FindAll()[0].Properties["whenChanged"][0].ToString();
            }
            if (sReturn == null) return "";
            return sReturn;
        }

        public static string adGetUserAttributeListHtml(string sUserId)
        {

            DirectoryEntry searchRoot = new DirectoryEntry(GetCurrentDomainPath(), null, null, AuthenticationTypes.Secure);

            System.Text.StringBuilder sb = new System.Text.StringBuilder();

            using (searchRoot)
            {
                DirectorySearcher ds = new DirectorySearcher(searchRoot, "(sAMAccountName=" + sUserId + ")");

　
                ds.SizeLimit = 1;

                SearchResult sr = null;

                using (SearchResultCollection src = ds.FindAll())
                {

                    foreach (SearchResult searchResult in src)
                    {
                        foreach (string propName in searchResult.Properties.PropertyNames)
                        {
                            ResultPropertyValueCollection valueCollection = searchResult.Properties[propName];
                            foreach (Object propertyValue in valueCollection)
                            {
                                sb.Append("Property: " + propName + ": " + propertyValue.ToString());
                                sb.Append("<br>");

                            }//foreach
                        }//foreach
                    }//foreach
                }//using

                return sb.ToString();
            }//using (searchRoot)

　
　
        }// end getAttributesAll()

　
    }// end cActiveDir===============================================================================================================================================

}//end namespace
