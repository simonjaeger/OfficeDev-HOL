using Single_Sign_On_Outlook_Add_inWeb.Models;
using System.Threading.Tasks;

namespace Single_Sign_On_Outlook_Add_inWeb.Services
{
    interface IUserService
    {
        // Try to get a user service user mapped to an Office 365 user 
        // with the UUID - created by the Exchange Web Services Managed API. 
        Task<UserModel> GetUserAsync(string uuid);

        // Get a unique salt for the UUID - which is used when hashing 
        // it for user mapping. If it doesn't exist in the store - it
        // will be randomly generated. 
        Task<byte[]> GetUUIDSaltAsync(string uuid);

        // Map a user service user to an Office 365 user with the 
        // UUID. If the mapping is successful, the user is returned.
        Task<UserModel> MapUserAsync(string uuid, CredentialsModel credentials);
    }
}
