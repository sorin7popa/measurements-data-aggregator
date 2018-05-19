using System;
using System.Linq;
using Google.Apis.Drive.v3;
using Google.Apis.Sheets.v4;
using MeasurementsDataAggregator.DataProcessing;
using MeasurementsDataAggregator.GoogleApi;

namespace MeasurementsDataAggregator
{
    internal class Program
    {
        private static readonly string[] AuthScopes = { DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets };
        private const string ApplicationName = "Measurements Data Aggregator";
        private const string SummaryFilePrefix = "Summary";
        private const string MeasurementsFolderId = "1CxDEkbYweo0vffewWhDDCLF4uVVR6c2i";
        private static readonly TimeSpan DateAssociationTreshold = TimeSpan.FromDays(5);

        private static readonly AuthorizationProvider AuthorizationProvider = new AuthorizationProvider();
        private static readonly ServiceFactory ServiceFactory = new ServiceFactory(ApplicationName);
        private static readonly Repository Repository = new Repository(SummaryFilePrefix);
        private static readonly SpreadsheetReader SpreadsheetReader = new SpreadsheetReader();
        private static readonly Processor Processor = new Processor(SpreadsheetReader, DateAssociationTreshold);

        private static void Main()
        {
            var credential = AuthorizationProvider.Authorize(AuthScopes);
            var driveService = ServiceFactory.GetDriveService(credential);
            var sheetsService = ServiceFactory.GetSheetsService(credential);

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