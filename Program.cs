using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;
using Mono.Options;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace ExcelDataExporter
{
    class Program
    {
        static int Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string excelDir = AppDomain.CurrentDomain.BaseDirectory;
            string jsonDir = AppDomain.CurrentDomain.BaseDirectory;
            bool showHelp = false;

            var p = new OptionSet()
            {
                { "e|excelDir=", "game data excel files input directory", v => excelDir = v },
                { "j|jsonDir=", "json files output directory", v => excelDir = v },
                { "h|help=", "show this message and exit", v => excelDir = v },
            };

            try
            {
                p.Parse(args);
            }
            catch (OptionException)
            {
                Console.WriteLine("Try '--help' for more information");
                return -1;
            }

            if (showHelp)
            {
                p.WriteOptionDescriptions(Console.Out);
                return 0;
            }

            try
            {
                foreach (var file in Directory.GetFiles(excelDir, "*.xlsx"))
                {
                    string filename = Path.GetFileName(file).Split(".").First();
                    string excelPath = Path.Combine(excelDir, file);

                    using var excelPackage = new ExcelPackage(excelPath);
                    foreach (var sheet in excelPackage.Workbook.Worksheets)
                    {
                        if (!sheet.Name.Equals(filename)) continue;

                        string outputPath = Path.Combine(excelDir, $"GameData{filename}.json");

                        using var stream = new FileStream(outputPath, FileMode.Create);
                        using var writer = new Utf8JsonWriter(stream, new JsonWriterOptions() { Indented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping });

                        writer.WriteStartObject();

                        var start = sheet.Dimension.Start;
                        var end = sheet.Dimension.End;

                        var nameHeaderRow = start.Row;
                        int typeHeaderRow = start.Row + 1;

                        for (int r = start.Row; r <= end.Row; r++)
                        {
                            if (r == nameHeaderRow || r == typeHeaderRow) continue;

                            for (int c = start.Column; c <= end.Column; c++)
                            {
                                string name = sheet.GetValue<string>(nameHeaderRow, c);
                                if (name.Equals("Key"))
                                {
                                    var key = sheet.GetValue<string>(r, c);
                                    if (string.IsNullOrEmpty(key))
                                    {
                                        break;
                                    }

                                    writer.WriteStartObject(key);
                                }

                                string type = sheet.GetValue<string>(typeHeaderRow, c);
                                switch (type.ToLower())
                                {
                                    case "bool":
                                        writer.WriteBoolean(name, sheet.GetValue<bool>(r, c));
                                        break;
                                    case "float":
                                        writer.WriteNumber(name, sheet.GetValue<float>(r, c));
                                        break;
                                    case "int":
                                        writer.WriteNumber(name, sheet.GetValue<int>(r, c));
                                        break;
                                    case "string":
                                        writer.WriteString(name, sheet.GetValue<string>(r, c) ?? "");
                                        break;
                                }
                            }

                            writer.WriteEndObject();
                        }

                        writer.WriteEndObject();
                        writer.Flush();
                    }
                }
            }
            catch
            {
                return -1;
            }

            return 0;
        }
    }
}
