using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Sheets.v4;
using MeasurementsDataAggregator.Model;

namespace MeasurementsDataAggregator.DataProcessing
{
    class Processor
    {
        private readonly SpreadsheetReader _spreadsheetReader;
        private readonly TimeSpan _dateAssociationTreshold;

        public Processor(SpreadsheetReader spreadsheetReader, TimeSpan dateAssociationTreshold)
        {
            _spreadsheetReader = spreadsheetReader;
            _dateAssociationTreshold = dateAssociationTreshold;
        }

        public List<Measurement> AggregateData(SheetsService sheetsService, List<File> inputFiles)
        {
            var measurements = new List<Measurement>();
            for (var i = 0; i < inputFiles.Count; i++)
            {
                var file = inputFiles[i];
                var parsedRecords = _spreadsheetReader.Parse(sheetsService, file).ToList();
                Console.WriteLine($"Succesfully parsed {parsedRecords.Count} measurements from file '{file.Name}' ({i} of {inputFiles.Count})");
                measurements.AddRange(parsedRecords);
            }

            return measurements;
        }
        public void NormalizeDates(List<Measurement> data)
        {
            var dateMappings = new Dictionary<DateTime, DateTime>();

            var uniqueDates = data.Select(d => d.Date).Distinct().OrderBy(d => d).ToList();
            for (var i = 0; i < data.Count - 1; i++)
            {
                if (uniqueDates[i+1] - uniqueDates[i] < _dateAssociationTreshold)
                {
                    dateMappings.Add(uniqueDates[i], uniqueDates[i+1]);
                }
            }

            for (var i = 0; i < data.Count; i++)
            {
                var measurement = data[i];
                if (dateMappings.ContainsKey(measurement.Date))
                {
                    measurement.Date = dateMappings[measurement.Date];
                    i--;
                }
            }
        }

        public IEnumerable<MeasurementAverage> CalculateAverages(List<Measurement> data)
        {
            var uniqueDates = data.Select(d => d.Date).Distinct().OrderBy(d => d).ToList();
            foreach (var uniqueDate in uniqueDates)
            {
                var dataSubset = data.Where(d => d.Date == uniqueDate).ToList();
                var measurementAverage = new MeasurementAverage
                {
                    Date = uniqueDate,
                    Weight = dataSubset.Select(d => d.Weight).Average(),
                    MuscleMass = dataSubset.Select(d => d.MuscleMass).Average(),
                    BMI = dataSubset.Select(d => d.BMI).Average(),
                    VisceralFatIndex = dataSubset.Select(d => d.VisceralFatIndex).Average(),
                    FatMassPercentage = dataSubset.Select(d => d.FatMassPercentage).Average(),
                    MeasurementsCount = dataSubset.Count
                };
                yield return measurementAverage;
            }
        }
    }
}
