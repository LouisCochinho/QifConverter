using ExcelDataReader;
using ShellProgressBar;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace QifConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Title = "QifConverter";
            Console.WriteLine(@"
   ____ _  __   ___                          _            
  /___ (_)/ _| / __\___  _ ____   _____ _ __| |_ ___ _ __ 
 //  / / | |_ / /  / _ \| '_ \ \ / / _ \ '__| __/ _ \ '__|
/ \_/ /| |  _/ /__| (_) | | | \ V /  __/ |  | ||  __/ |   
\___,_\|_|_| \____/\___/|_| |_|\_/ \___|_|   \__\___|_|   
                                                          
");
            Console.WriteLine("Press any key to start conversion...");
            Console.ReadLine();

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            CheckArgument(args);

            var repertoryPath = args[0];

            Console.WriteLine($"Repertory path is {repertoryPath}");

            var fileNames = GetExcelFileNamesFromDirectoryPath(repertoryPath);

            var rows = ConvertFilesContentToRows(fileNames);

            var qifContent = ConvertRowsToQif(rows);
            var qifFilePath = $"{repertoryPath}/result.qif";

            WriteInFile(qifContent, qifFilePath);


            Console.WriteLine("Excel > QIF ---> OK");

            Exit($"Conversion succeeded. Qif file is {qifFilePath}", 0);
        }

        private static void CheckArgument(string[] args)
        {
            if (!args.Any())
            {
                Exit("The program expects an argument : the path to the repertory is not specified", 1);
            }

            if (!Directory.Exists(args[0]))
            {
                Exit($"The specified repertory doesn't exist. Please check the repertory path. Specified repertory path is : {args[0]}", 1);
            }
        }

        private static string[] GetExcelFileNamesFromDirectoryPath(string directoryPath)
        {
            return Directory.GetFiles(directoryPath, "*.xlsx");
        }

        private static List<Row> ConvertFilesContentToRows(string[] fileNames)
        {
            var rows = new List<Row>();
            var excelDataSetConf = new ExcelDataSetConfiguration()
            {
                UseColumnDataType = false,
                FilterSheet = (tableReader, sheetIndex) => true,
                ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            };

            using (var progressBar = new ProgressBar(fileNames.Length, "Processing excel files..."))
            {
                foreach (var fileName in fileNames)
                {
                    using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet(excelDataSetConf);

                            for (int i = 0; i < result.Tables.Count; i++)
                            {
                                for (int j = result.Tables[i].Rows.Count - 1; j >= 0; j--)
                                {
                                    try
                                    {
                                        var rowString = result.Tables[i].Rows[j];

                                        rows.Add(new Row
                                        {
                                            Date = DateTime.Parse(rowString.ItemArray[0].ToString()),
                                            Label = ProcessLabel(rowString.ItemArray[1].ToString()),
                                            Amount = rowString.ItemArray[2].ToString()
                                        });
                                    }
                                    catch (Exception e)
                                    {
                                        Exit($"Unknown error during conversion : {e}", 1);
                                    }
                                    finally
                                    {
                                        progressBar.Tick();
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return rows.OrderBy(x => x.Date).ToList();
        }

        private static string ConvertRowsToQif(List<Row> rows)
        {
            var qif = "!Type:Bank\n";
            using (var progressBar = new ProgressBar(rows.Count, "Converting to Qif file..."))
            {
                try
                {
                    foreach (var row in rows)
                    {
                        qif += $"D{row.Date.ToShortDateString()}\n" +
                            $"T{row.Amount}\n" +
                            $"M{row.Label}\n^\n";
                    }
                }
                catch (Exception e)
                {
                    Exit($"Unknown error during conversion : {e}", 1);
                }
                finally
                {
                    progressBar.Tick();
                }
            }

            return qif;
        }

        private static void WriteInFile(string content, string path)
        {
            File.WriteAllText(path, content, Encoding.UTF8);
        }

        private static void Exit(string message, int exitCode)
        {
            Console.WriteLine($"{message}");
            Console.WriteLine("\nPress any key to exit...");
            Console.ReadLine();
            Environment.Exit(exitCode);
        }

        private static string ProcessLabel(string label)
        {
            return label.Replace("&amp;", "&");
        }
    }
}