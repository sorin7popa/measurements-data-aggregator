using MeasurementsDataAggregator.Model;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace MeasurementsDataAggregator.DataProcessing
{
    class SpreadsheetParser
    {
        public IEnumerable<Measurement> Parse(IList<IList<object>> sheetData, string fileName)
        {
            for (var i = 0; i < sheetData.Count; i++)
            {
                var row = sheetData[i];
                if (row == null || row.Count < 8)
                {
                    break;
                }

                var dateOk = DateTime.TryParseExact(row[0].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out var date);
                if (!dateOk || date == DateTime.MinValue)
                {
                    dateOk = DateTime.TryParseExact(row[0].ToString(), "yyyy-MM-dd", CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out date);
                }
                var weightOk = float.TryParse(row[2].ToString().Replace(',', '.'), out var weight);
                var muscleMassOk = float.TryParse(row[4].ToString().Replace(',', '.'), out var muscleMass);
                var bmiOk = float.TryParse(row[5].ToString().Replace(',', '.'), out var bmi);
                var visceralFatIndexOk = float.TryParse(row[6].ToString().Replace(',', '.'), out var visceralFatIndex);
                var fatMassPercentageOk = float.TryParse(row[7].ToString().Replace(',', '.'), out var fatMassPercentage);

                if (dateOk && DateTime.MinValue != date
                           && weightOk && default(float) != weight
                           && muscleMassOk && default(float) != muscleMass
                           && bmiOk && default(float) != bmi
                           && visceralFatIndexOk && default(float) != visceralFatIndex
                           && fatMassPercentageOk && default(float) != fatMassPercentage)
                {
                    yield return new Measurement
                    {
                        Date = date,
                        Weight = weight,
                        MuscleMass = muscleMass,
                        BMI = bmi,
                        VisceralFatIndex = visceralFatIndex,
                        FatMassPercentage = fatMassPercentage
                    };
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Unable to parse line {8+i} from file '{fileName}'");
                    Console.ForegroundColor = ConsoleColor.Gray;
                }
            }
        }
    }
}
