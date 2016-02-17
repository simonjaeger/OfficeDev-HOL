using System.Collections.Generic;

namespace Microsoft_Graph_Mail_Console_App
{
    public class UserModel
    {
        public string Id { get; set; }
        public string UserPrincipalName { get; set; }
        public List<string> BusinessPhones { get; set; }
        public string DisplayName { get; set; }
        public string GivenName { get; set; }
        public object JobTitle { get; set; }
        public string Mail { get; set; }
        public string MobilePhone { get; set; }
        public object OfficeLocation { get; set; }
        public string PreferredLanguage { get; set; }
        public string Surname { get; set; }
    }
}
