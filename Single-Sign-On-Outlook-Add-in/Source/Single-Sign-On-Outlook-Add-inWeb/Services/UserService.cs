    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Single_Sign_On_Outlook_Add_inWeb.Models;

    namespace Single_Sign_On_Outlook_Add_inWeb.Services
    {
        public class UserService : IUserService
        {
            private readonly List<UserModel> _users;
            private readonly Dictionary<string, UserModel> _ssoMap;
            private readonly Dictionary<string, byte[]> _saltMap;
            private readonly Random _random;

            public UserService()
            {
                _users = new List<UserModel>();
                _ssoMap = new Dictionary<string, UserModel>();
                _saltMap = new Dictionary<string, byte[]>();
                _random = new Random();

                // Add a new user that we will use in the mail add-in
                _users.Add(new UserModel
                {
                    DisplayName = "Office 365 Developer",
                    Credentials = new CredentialsModel
                    {
                        Username = "#Office365Dev",
                        Password = "#Office365Dev"
                    }
                });
            }

            // Try to get a user service user mapped to an Office 365 user 
            // with the UUID - created by the Exchange Web Services Managed API. 
            public async Task<UserModel> GetUserAsync(string uuid)
            {
                if (_ssoMap.ContainsKey(uuid))
                    return _ssoMap[uuid];
                return null;
            }

            // Get a unique salt for the UUID - which is used when hashing 
            // it for user mapping. If it doesn't exist in the store - it
            // will be randomly generated. 
            public async Task<byte[]> GetUUIDSaltAsync(string uuid)
            {
                // Check if a salt for the UUID already exists
                if (_saltMap.ContainsKey(uuid))
                    return _saltMap[uuid];

                // Generate a new unique salt and add it
                var salt = new byte[16];
                _random.NextBytes(salt);

                _saltMap.Add(uuid, salt);
                return salt;
            }

            // Map a user service user to an Office 365 user with the 
            // UUID. If the mapping is successful, the user is returned.
            public async Task<UserModel> MapUserAsync(string uuid, CredentialsModel credentials)
            {
                var user = _users.FirstOrDefault(u => 
                           u.Credentials.Username.Equals(credentials.Username) &&
                           u.Credentials.Password.Equals(credentials.Password));

                if (user == null)
                    return null;

                _ssoMap.Add(uuid, user);
                return user;
            }
        }
    }