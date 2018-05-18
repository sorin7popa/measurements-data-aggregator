using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Sheets.v4;
using MeasurementsDataAggregator.DataProcessing;
using MeasurementsDataAggregator.GoogleApi;
using System;
using System.Linq;

namespace MeasurementsDataAggregator
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        private static readonly string ApplicationName = "Measurements Data Aggregator";
        private static readonly string SummaryFilePrefix = "Summary";
        private static readonly string MeasurementsFolderId = "1CxDEkbYweo0vffewWhDDCLF4uVVR6c2i";
        private static readonly TimeSpan DateAssociationTreshold = TimeSpan.FromDays(5);

        private static readonly AuthorizationProvider AuthorizationProvider = new AuthorizationProvider();
        private static readonly ServiceFactory ServiceFactory = new ServiceFactory(ApplicationName);
        private static readonly Repository Repository = new Repository(SummaryFilePrefix);
        private static readonly SpreadsheetReader SpreadsheetReader = new SpreadsheetReader();
        private static readonly Processor Processor = new Processor(SpreadsheetReader, DateAssociationTreshold);

        static void Main(string[] args)
        {
            UserCredential driveCredential = AuthorizationProvider.AuthorizeDrive();
            DriveService driveService = ServiceFactory.GetDriveService(driveCredential);

            UserCredential sheetCredential = AuthorizationProvider.AuthorizeSheets();
            SheetsService sheetsService = ServiceFactory.GetSheetsService(sheetCredential);

            var allFiles = Repository.GetFiles(driveService, MeasurementsFolderId);
            var inputFiles = allFiles.Where(f => !f.Name.StartsWith(SummaryFilePrefix)).ToList();

            var data = Processor.AggregateData(sheetsService, inputFiles);
            Processor.NormalizeDates(data);
            var measurementAverages = Processor.CalculateAverages(data).ToList();

            Repository.WriteData(measurementAverages);

            Console.WriteLine();
            Console.WriteLine("Done! Press any key to exit...");
            Console.Read();
        }
    }
}