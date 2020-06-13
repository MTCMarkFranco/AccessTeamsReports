using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Graph;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessTeamsReports.Utilities
{
    public static class Helper
    {
        static Helper()
        {
            
        }

        public static Report DeserializeCSVReportObject<T>(string inputString)
        {
            T[] records = null;

            // Remove the API stamp at the end of the file
            string csvResponse = Helper.SanitizeCSV(inputString);

            using (TextReader reader = new StringReader(csvResponse))

            {
                using (var csv = new CsvReader(reader, System.Globalization.CultureInfo.GetCultureInfo("en-CA")))
                {
                    csv.Configuration.HasHeaderRecord = true;
                    csv.Configuration.MissingFieldFound = null;
                    csv.Configuration.TrimOptions = TrimOptions.Trim;
                    csv.Configuration.BadDataFound = null;

                    // Step 1. Deserialize CSV to POCO (TeamsUserActivityUserDetail)
                    records = csv.GetRecords<T>().ToArray();

                    // Step 2. Serialize POCO to JSON String and fill the MemoryStream to pass back to the pipeline
                    var jsonString = JsonConvert.SerializeObject(records, Formatting.Indented);
                    var jsonMemorryStream = new MemoryStream(Encoding.Default.GetBytes(jsonString));

                    // Step 3. Create the new Return value (Report Microsoft.Graph.Class)
                    Report report = new Report();
                    report.Content = jsonMemorryStream;

                    return report;

                }
            }
        }


        /// <summary>
        /// Removes the API Appendage of the HTTP call Known Bug.
        /// </summary>
        /// <param name="inputString"></param>
        /// <returns></returns>
        private static string SanitizeCSV(string inputString)
        {
            StringBuilder sb = new StringBuilder();

            using (TextReader reader = new StringReader(inputString))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (!line.Contains("responseHeaders"))
                    {
                        sb.AppendLine(line);
                    }
                }
            }

            return sb.ToString();
        }

    }

}
