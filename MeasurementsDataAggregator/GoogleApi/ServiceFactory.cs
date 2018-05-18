using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;

namespace MeasurementsDataAggregator.GoogleApi
{
    internal class ServiceFactory
    {
        private readonly string _applicationName;

        public ServiceFactory(string applicationName)
        {
            _applicationName = applicationName;
        }
        public DriveService GetDriveService(UserCredential credential)
        {
            // Create Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = _applicationName,
            });
            return service;
        }

        public SheetsService GetSheetsService(UserCredential credential)
        {
            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = _applicationName,
            });
        }
    }
}