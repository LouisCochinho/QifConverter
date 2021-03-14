using BankFileConverter.Api.Models;
using BankFileConverter.Api.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Threading.Tasks;

namespace BankFileConverter.Api.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class BankFileConverterController : ControllerBase
    {
        private readonly ILogger<BankFileConverterController> _logger;
        private readonly IBankFileConverterService _bankFileConverterService;

        public BankFileConverterController(ILogger<BankFileConverterController> logger, IBankFileConverterService qifConverterService)
        {
            _logger = logger;
            _bankFileConverterService = qifConverterService;
        }

        /// <summary>
        /// Convert excel file(s) content to the specified output bank file format.
        /// </summary>
        /// <param name="inputPath">The path to the repertory containing the input excel files.</param>
        /// <param name="outputPath">The path to the repertory containing the output bank files.</param>
        /// <param name="initialAmount">An initial amount on which starting transactions in QIF file.</param>
        /// <param name="onlyTransactions">If true, the excel file is only containing transactions amount. If false, it is only containing amounts.</param>
        /// <param name="inputFileType"> The input file type.</param>
        /// <param name="outputFileType"> The output bank file format.</param>
        /// <response code="200">OK.</response>
        /// <response code="400">BadRequest : invalid url provided or empty excel file(s).</response>
        /// <response code="404">No excel file(s) found.</response>
        /// <response code="500">Internal server error.</response>
        /// <response code="501">Not implemented.</response>
        /// <returns>QIF or OFX file.</returns>
        [HttpPost("file")]
        [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        [ProducesResponseType(StatusCodes.Status501NotImplemented)]
        [Produces("application/json")]
        public async Task<IActionResult> ConvertToBankFile(
            [Required] string inputPath,
            string outputPath,
            InputFileType inputFileType = InputFileType.ANY,
            OutputFileType outputFileType = OutputFileType.QIF,
            bool onlyTransactions = true,
            float initialAmount = 0)
        {
            if (!inputFileType.Equals(InputFileType.ANY) || !outputFileType.Equals(OutputFileType.QIF))
            {
                _logger.LogError($"Not implemented input File type.");
                return new StatusCodeResult(501);
            }

            if (!Directory.Exists(inputPath) || (outputPath != null && !Directory.Exists(outputPath)))
            {
                _logger.LogWarning("Malformed url");
                return new BadRequestObjectResult("Malformed url");
            }

            _logger.LogDebug($"Gettings excel files from {inputPath}...");

            var fileNames = _bankFileConverterService.GetExcelFileNamesFromDirectoryPath(inputPath);

            if (fileNames == null || fileNames.Length == 0)
            {
                _logger.LogWarning("No file(s) to scan.");
                return new NotFoundObjectResult("No excel files to scan.");
            }

            _logger.LogDebug($"Converting excel files from {inputPath} into bank file...");

            var rows = _bankFileConverterService.ConvertFilesContentToRows(fileNames, inputFileType);

            if (rows == null || rows.Count == 0)
            {
                _logger.LogWarning("Empty file(s).");
                return new BadRequestObjectResult("Empty file(s).");
            }

            var qif = _bankFileConverterService.ConvertRowsToBankFile(rows, initialAmount, onlyTransactions, outputFileType);

            if (string.IsNullOrEmpty(qif))
            {
                _logger.LogWarning("Empty qif.");
                return new BadRequestObjectResult("Empty qif.");
            }

            var fileName = "result.qif";
            var qifFileUri = string.IsNullOrEmpty(outputPath) ? $"{inputPath}/{fileName}" : $"{outputPath}/{fileName}";

            await _bankFileConverterService.WriteInFile(qif, qifFileUri);

            if (!System.IO.File.Exists(qifFileUri))
            {
                _logger.LogError($"Error when creating qif file.");
                return new BadRequestObjectResult("Error when creating qif file.");
            }

            return File(await System.IO.File.ReadAllBytesAsync(qifFileUri), "application/octet-stream", fileName);
        }
    }
}
