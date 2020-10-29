using ExcelDataReader;
using System;
using System.IO;
using CsvHelper;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;

namespace ConvertExcelToText
{
    class Program
    {
        static void Main(string[] args)
        {
            IList<RowSheet> list = new List<RowSheet>();

            IConfiguration configuration = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                .Build();

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(configuration["read"], FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read()) //Each ROW
                    {


                        if (AreAllColumnsEmpty(reader))
                            continue;

                        var row = new RowSheet
                        {
                            Column1 = GetValueColumn(reader, 0),
                            Column2 = GetValueColumn(reader, 1),
                            Column3 = GetValueColumn(reader, 2),
                            Column4 = GetValueColumn(reader, 3),
                            Column5 = GetValueColumn(reader, 4),
                            Column6 = GetValueColumn(reader, 5),
                            Column7 = GetValueColumn(reader, 6),
                            Column8 = GetValueColumn(reader, 7),
                            Column9 = GetValueColumn(reader, 8),
                            Column10 = GetValueColumn(reader, 9),
                            Column11 = GetValueColumn(reader, 10),
                            Column12 = GetValueColumn(reader, 11),
                            Column13 = GetValueColumn(reader, 12),
                            Column14 = GetValueColumn(reader, 13),
                            Column15 = GetValueColumn(reader, 14),
                            Column16 = GetValueColumn(reader, 15),
                            Column17 = GetValueColumn(reader, 16),
                            Column18 = GetValueColumn(reader, 17),
                            Column19 = GetValueColumn(reader, 18),
                            Column20 = GetValueColumn(reader, 19),
                            Column21 = GetValueColumn(reader, 20)

                        };

                        list.Add(row);
                    }
                }
            }

            CsvHelper.Configuration.CsvConfiguration configurationCsv = new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = "|",
                HasHeaderRecord = false
            };

            using (var writer = new StreamWriter(configuration["write"]))
            using (var csv = new CsvWriter(writer, configurationCsv))
            {
                csv.WriteRecords(list);
            }
        }

        private static bool AreAllColumnsEmpty(IExcelDataReader reader)
        {
            bool result = true;

            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (!reader.IsDBNull(i))
                {
                    result = false;
                    break;
                }
            }

            return result;
        }

        private static string GetValueColumn(IExcelDataReader reader, int idx)
        {
            var value = reader.GetValue(idx);

            return value != null ? value.ToString() : string.Empty;
        }
    }
}
