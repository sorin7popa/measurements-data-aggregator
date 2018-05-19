using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Sheets.v4;
using MeasurementsDataAggregator.DataProcessing;
using MeasurementsDataAggregator.Model;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace MeasurementsDataAggregator.GoogleApi
{
    internal class Repository
    {
        private readonly SpreadsheetParser _spreadsheetParser;
        private readonly string _summaryFilePrefix;

        public Repository(SpreadsheetParser spreadsheetParser, string summaryFilePrefix)
        {
            _spreadsheetParser = spreadsheetParser;
            _summaryFilePrefix = summaryFilePrefix;
        }

        public List<File> GetFiles(DriveService driveService, string folderId)
        {
            var listRequest = driveService.Files.List();
            listRequest.Q = $"'{folderId}' in parents";
            listRequest.PageSize = 1000;
            listRequest.Fields = "nextPageToken, files(id, name)";

            var files = listRequest.Execute().Files.ToList();
            return files;
        }

        public List<Measurement> ReadAllSheets(SheetsService sheetsService, List<File> inputFiles)
        {
            var measurements = new List<Measurement>();
            for (var i = 0; i < inputFiles.Count; i++)
            {
                var file = inputFiles[i];
                var sheetData = ReadSheetData(sheetsService, file);
                var parsedRecords = _spreadsheetParser.Parse(sheetData, file.Name).ToList();
                Console.WriteLine(
                    $"Succesfully parsed {parsedRecords.Count} measurements from file '{file.Name}' (file {i + 1} of {inputFiles.Count})");
                measurements.AddRange(parsedRecords);
            }

            return measurements;
        }

        private IList<IList<object>> ReadSheetData(SheetsService sheetsService, File file)
        {
            Thread.Sleep(1000);
            const string range = "A8:H";
            var request = sheetsService.Spreadsheets.Values.Get(file.Id, range);
            try
            {
                var response = request.Execute();
                var values = response.Values;
                return values;
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Exception reading range {range} in file {file.Name}\n{e}");
                Console.ForegroundColor = ConsoleColor.Gray;
                return new List<IList<object>>();
            }
        }

        public void WriteData(List<MeasurementAverage> measurementAverages)
        {
            var dataOutputFileName = $"{_summaryFilePrefix}_{DateTime.Now.Ticks}.csv";

            var headerLine = string.Join(",", new List<string>
            {
                nameof(MeasurementAverage.Date),
                nameof(MeasurementAverage.Weight),
                nameof(MeasurementAverage.MuscleMass),
                nameof(MeasurementAverage.BMI),
                nameof(MeasurementAverage.VisceralFatIndex),
                nameof(MeasurementAverage.FatMassPercentage),
                nameof(MeasurementAverage.MeasurementsCount),
            });

            var dataLines = measurementAverages.Select(average => string.Join(",", new List<string>
            {
                average.Date.ToShortDateString(),
                average.Weight.ToString(CultureInfo.InvariantCulture),
                average.MuscleMass.ToString(CultureInfo.InvariantCulture),
                average.BMI.ToString(CultureInfo.InvariantCulture),
                average.VisceralFatIndex.ToString(CultureInfo.InvariantCulture),
                average.FatMassPercentage.ToString(CultureInfo.InvariantCulture),
                average.MeasurementsCount.ToString(CultureInfo.InvariantCulture),
            })).ToList();

            dataLines.Insert(0, headerLine);

            System.IO.File.WriteAllLines(dataOutputFileName, dataLines);
        }
    }
}
