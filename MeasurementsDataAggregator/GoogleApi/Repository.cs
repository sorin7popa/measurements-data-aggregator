using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using MeasurementsDataAggregator.Model;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace MeasurementsDataAggregator.GoogleApi
{
    class Repository
    {
        private readonly string _summaryFilePrefix;

        public Repository(string summaryFilePrefix)
        {
            _summaryFilePrefix = summaryFilePrefix;
        }

        public IList<File> GetFiles(DriveService driveService, string folderId)
        {
            FilesResource.ListRequest listRequest = driveService.Files.List();
            listRequest.Q = $"'{folderId}' in parents";
            listRequest.PageSize = 1000;
            listRequest.Fields = "nextPageToken, files(id, name)";

            IList<File> files = listRequest.Execute()
                .Files;
            return files;
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
