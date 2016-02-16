namespace Single_Sign_On_Outlook_Add_inWeb.Models
{
    public class UserRequestModel
    {
        // The raw identity token given by the mail add-in
        public string IdentityToken { get; set; }

        // The source location of the add-in
        public string HostUri { get; set; }

        // Credentials for a user in the web service - may be null.
        // Used when mapping web service users and Office 365 users.
        public CredentialsModel Credentials { get; set; }
    }
}
