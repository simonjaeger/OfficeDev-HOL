using Microsoft.Exchange.WebServices.Auth.Validation;
using Single_Sign_On_Outlook_Add_inWeb.Models;
using Single_Sign_On_Outlook_Add_inWeb.Services;
using System;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

namespace Single_Sign_On_Outlook_Add_inWeb.Controllers
{
    public class SSOController : ApiController
    {
        private static readonly IUserService _userService = new UserService();

        public async Task<UserModel> Post([FromBody]UserRequestModel request)
        {
            try
            {
                // Parse and validate the raw identity token
                var identityToken = (AppIdentityToken)AuthToken.Parse(request.IdentityToken);
                identityToken.Validate(new Uri(request.HostUri));

                // Get the UUID and the salt from the web service
                var uuid = identityToken.UniqueUserIdentification;
                var salt = await _userService.GetUUIDSaltAsync(uuid);

                // Combine and hash it
                var hashedUuid = GetHashedUUID(uuid, salt);

                // Try to find a user mapping with the hashed UUID
                var user = await _userService.GetUserAsync(hashedUuid);

                // If a user mapping is found for the unique user
                // identification - return the user.
                if (user != null)
                {
                    return user;
                }

                // End here if there are no credentials in the request
                if (request.Credentials == null)
                {
                    return null;
                }

                // Credentials are present, try to map and return a user
                // if the credentials are valid for the web service.
                user = await _userService.MapUserAsync(hashedUuid, request.Credentials);

                if (user == null)
                {
                    // Mapping failed.
                    return null;
                }
                return user;
            }
            catch (TokenValidationException ex)
            {
                //throw new ApplicationException("An error occurred while validating the identity token.", ex);
                return null;
            }
        }

        private string GetHashedUUID(string uuid, byte[] salt)
        {
            var input = Encoding.UTF8.GetBytes(uuid);
            var saltedInput = new byte[input.Length + salt.Length];

            // Combine the UUID and the salt
            salt.CopyTo(saltedInput, 0);
            input.CopyTo(saltedInput, salt.Length);

            // Compute the hash
            byte[] hash;
            using (var sha265 = new SHA256Managed())
            {
                hash = sha265.ComputeHash(saltedInput);
            }

            return BitConverter.ToString(hash);
        }
    }
}