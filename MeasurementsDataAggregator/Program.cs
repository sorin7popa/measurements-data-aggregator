using Google.Apis.Drive.v3;
using Google.Apis.Sheets.v4;
using MeasurementsDataAggregator.DataProcessing;
using MeasurementsDataAggregator.GoogleApi;
using System;
using System.Linq;

namespace MeasurementsDataAggregator
{
    internal class Program
    {
        private static readonly string[] AuthScopes = { DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets };
        private const string ApplicationName = "Measurements Data Aggregator";
        private const string SummaryFilePrefix = "Summary";
        private const string MeasurementsFolderId = "0B15Usj0yt4YOTDk3NmNpQkxMZUE";
        private static readonly TimeSpan DateAssociationTreshold = TimeSpan.FromDays(30);

        private static readonly AuthorizationProvider AuthorizationProvider = new AuthorizationProvider();
        private static readonly ServiceFactory ServiceFactory = new ServiceFactory(ApplicationName);
        private static readonly SpreadsheetParser SpreadsheetParser = new SpreadsheetParser();
        private static readonly Repository Repository = new Repository(SpreadsheetParser, SummaryFilePrefix);
        private static readonly Processor Processor = new Processor(DateAssociationTreshold);

        private static void Main()
        {
            var credential = AuthorizationProvider.Authorize(AuthScopes);
            var driveService = ServiceFactory.GetDriveService(credential);
            var sheetsService = ServiceFactory.GetSheetsService(credential);

            var files = Repository.GetFiles(driveService, MeasurementsFolderId);

            var data = Repository.ReadAllSheets(sheetsService, files);
            Processor.NormalizeDates(data);
            var measurementAverages = Processor.CalculateAverages(data).ToList();

            Repository.WriteData(measurementAverages);

            Console.WriteLine();
            Console.WriteLine("Done! Press any key to exit...");
            Console.Read();
        }
    }
}