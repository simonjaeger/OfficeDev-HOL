    namespace Single_Sign_On_Outlook_Add_inWeb.Models
    {
        public class UserModel
        {
            // The display name of the user
            public string DisplayName { get; set; }

            // The credentials used to verify someone as the  user
            public CredentialsModel Credentials { get; set; }
        }
    }
