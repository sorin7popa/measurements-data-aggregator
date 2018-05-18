using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using MeasurementsDataAggregator.Model;

namespace MeasurementsDataAggregator.DataProcessing
{
    class SpreadsheetReader
    {
        public IEnumerable<Measurement> Parse(SheetsService sheetsService, File file)
        {
            var content = Read(sheetsService, file.Id);
            for (var i = 0; i < content.Count; i++)
            {
                var row = content[i];
                if (row == null || row.Count == 0)
                {
                    break;
                }

                var dateOk = DateTime.TryParseExact(row[0].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out var date);
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
                    Console.WriteLine($"Unable to parse line {8+i} from file '{file.Name}'");
                    Console.ForegroundColor = ConsoleColor.White;
                }
            }
        }

        private IList<IList<Object>> Read(SheetsService sheetsService, string fileId)
        {
            String range = "Sheet1!A8:H";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                sheetsService.Spreadsheets.Values.Get(fileId, range);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            return values;
        }
    }
}
