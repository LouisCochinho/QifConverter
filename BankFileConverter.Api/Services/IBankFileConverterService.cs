using BankFileConverter.Api.Models;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace BankFileConverter.Api.Services
{
    public interface IBankFileConverterService
    {
        string[] GetExcelFileNamesFromDirectoryPath(string directoryPath);

        List<Row> ConvertFilesContentToRows(string[] fileNames, InputFileType inputFileType);

        string ConvertRowsToBankFile(List<Row> rows, float initialAmount, bool onlyTransactions, OutputFileType outputFilePath);

        Task WriteInFile(string content, string path);
    }
}
