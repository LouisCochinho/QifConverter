using ExcelDataReader;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QifConverter.Api.Services
{
    public class QifConverterService : IQifConverterService
    {
        private readonly ILogger<QifConverterService> _logger;

        public QifConverterService(ILogger<QifConverterService> logger)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            _logger = logger;
        }

        public string[] GetExcelFileNamesFromDirectoryPath(string directoryPath)
        {
            return Directory.GetFiles(directoryPath, "*.xlsx");
        }

        public List<Row> ConvertFilesContentToRows(string[] fileNames)
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
                                    //Exit($"Unknown error during conversion : {e}", 1);
                                }
                            }
                        }
                    }
                }
            }

            return rows.OrderBy(x => x.Date).ToList();
        }

        public string ConvertRowsToQif(List<Row> rows, float initialAmount, bool onlyTransactions)
        {
            var qif = string.Empty;
            try
            {
                qif = RowsToQif(rows, initialAmount, onlyTransactions);
            }
            catch (Exception e)
            {
                //Exit($"Unknown error during conversion : {e}", 1);
            }


            return qif;
        }

        public async Task WriteInFile(string content, string path)
        {
            await File.WriteAllTextAsync(path, content, Encoding.UTF8);

        }

        #region private methods
        private string RowsToQif(List<Row> rows, float initialAmount, bool onlyTransactions)
        {
            var qif = "!Type:Bank\n";

            if (initialAmount != 0)
            {
                qif += RowToQif(rows[0].Date.ToShortDateString(), initialAmount.ToString(), "Initial Amount");
            }

            if (onlyTransactions)
            {
                foreach (var row in rows)
                {
                    qif += RowToQif(row.Date.ToShortDateString(), row.Amount.ToString(), row.Label);
                }
            }
            else
            {
                qif += RowToQif(rows[0].Date.ToShortDateString(), (float.Parse(rows[0].Amount) - initialAmount).ToString(), rows[0].Label);

                for (int i = 1; i < rows.Count; i++)
                {
                    qif += RowToQif(rows[i].Date.ToShortDateString(), (float.Parse(rows[i].Amount) - float.Parse(rows[i - 1].Amount)).ToString(), rows[i].Label);
                }
            }

            return qif;
        }

        private string RowToQif(string date, string amount, string label)
        {
            label = string.IsNullOrWhiteSpace(label) ? "Transaction" : label;
            return $"D{date}\n" +
                           $"T{amount}\n" +
                           $"M{label}\n^\n";
        }

        private string ProcessLabel(string label)
        {
            return label.Replace("&amp;", "&");
        }


        #endregion
    }
}
