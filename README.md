# Active-Directory-Methods
Various C# AD Methods



        public static string adALLUserGroupsReturnListJson()

        public static string adGetEmailAddressByFullName(string sFullName)

        public static string adGetUserAttributeListHtml(string sUserId)

        public static string adGetUserPrincipalMethodsHtml()

        public static string adGetUserPropertyListHtml(string sUserId)

        public static string adGetUserWhenChanged(string sUserId)

        public static string adGroupGetAllUsers(string sGroupName)

        public static string adGroupGetAllUsersIdInSqlWithQuotes(string sGroupName)

        public static string adUserInfo(string sUserId)

        public static string adUserInfoByEmail(string sEmailWithoutOrgDotCom)

        public static string listKeyValueSerialize(List<KeyValue> lst)

        public static string returnJsonListOfAdGroupsForUserId(string sUserId)

        public static UserPrincipal adGetUserPrincipalBySAMAccountName(string sAMAccountName)

        public static void adGetUserBySAMAccountName(string sAMAccountName, ref Dictionary<string,string> dictUser)

