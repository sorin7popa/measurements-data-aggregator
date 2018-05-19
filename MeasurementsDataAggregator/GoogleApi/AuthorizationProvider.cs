using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;
using System;
using System.IO;
using System.Threading;

namespace MeasurementsDataAggregator.GoogleApi
{
    internal class AuthorizationProvider
    {
        public UserCredential Authorize(string[] scopes)
        {
            UserCredential credential;
            var credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                credPath = Path.Combine(credPath, ".credentials/drive-sheets-api.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            return credential;
        }
    }
}