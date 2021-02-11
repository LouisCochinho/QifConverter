using ExcelDataReader;
using ShellProgressBar;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace QifConverter
{
    public class Program
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

            var initialAmount = AskInitialAmount();
            var onlyTransactions = AskIfOnlyTransactions();
            Console.WriteLine("Press any key to start conversion...");
            Console.ReadLine();

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            CheckArgument(args);

            var repertoryPath = args[0];

            Console.WriteLine($"Repertory path is {repertoryPath}");

            var fileNames = GetExcelFileNamesFromDirectoryPath(repertoryPath);

            var rows = ConvertFilesContentToRows(fileNames);

            var qifContent = ConvertRowsToQif(rows, initialAmount, onlyTransactions);
            var qifFilePath = $"{repertoryPath}/result.qif";

            WriteInFile(qifContent, qifFilePath);


            Console.WriteLine("Excel > QIF ---> OK");

            Exit($"Conversion succeeded. Qif file is {qifFilePath}", 0);
        }

        private static bool AskIfOnlyTransactions()
        {
            string response;
            do
            {
                Console.Write("Only Transactions ? (y or n) : ");
                response = Console.ReadLine();
            } while (response != "y" && response != "n");

            return response.Equals("y");
        }

        private static float AskInitialAmount()
        {
            float initialAmount;

            do
            {
                Console.Write("Initial amount ? (0 if no initial amount) : ");
            }
            while (!float.TryParse(Console.ReadLine(), out initialAmount));

            return initialAmount;
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

        private static string ConvertRowsToQif(List<Row> rows, float initialAmount, bool onlyTransactions)
        {
            var qif = string.Empty;
            using (var progressBar = new ProgressBar(rows.Count, "Converting to Qif file..."))
            {
                try
                {
                    qif = RowsToQif(rows, initialAmount, onlyTransactions);
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

        public static string RowsToQif(List<Row> rows, float initialAmount, bool onlyTransactions)
        {
            var qif = "!Type:Bank\n";

            if (initialAmount != 0)
            {
                qif += $"D{rows[0].Date.ToShortDateString()}\n" +
                       $"T{initialAmount}\n" +
                       $"MInitial Amount\n^\n";
            }

            if (onlyTransactions)
            {
                foreach (var row in rows)
                {
                    qif += $"D{row.Date.ToShortDateString()}\n" +
                           $"T{row.Amount}\n" +
                           $"M{row.Label}\n^\n";
                }
            }
            else
            {
                qif += $"D{rows[0].Date.ToShortDateString()}\n" +
                       $"T{float.Parse(rows[0].Amount) - initialAmount}\n" +
                       $"M{rows[0].Label}\n^\n";

                for (int i = 1; i < rows.Count; i++)
                {
                    qif += $"D{rows[i].Date.ToShortDateString()}\n" +
                           $"T{float.Parse(rows[i].Amount) - float.Parse(rows[i - 1].Amount)}\n" +
                           $"M{rows[i].Label}\n^\n";
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