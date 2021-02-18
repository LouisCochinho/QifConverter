using System.Collections.Generic;
using System.Threading.Tasks;

namespace QifConverter.Api.Services
{
    public interface IQifConverterService
    {
        string[] GetExcelFileNamesFromDirectoryPath(string directoryPath);

        List<Row> ConvertFilesContentToRows(string[] fileNames);

        string ConvertRowsToQif(List<Row> rows, float initialAmount, bool onlyTransactions);

        Task WriteInFile(string content, string path);
    }
}
