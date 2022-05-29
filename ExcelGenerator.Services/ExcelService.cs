using Microsoft.Extensions.Logging;
using OfficeOpenXml;

namespace ExcelGenerator.Services
{
    public class ExcelService : IExcelService
    {
        private readonly ILogger<ExcelService> _logger;

        public ExcelService(ILogger<ExcelService> logger)
        {
            _logger = logger;
        }

        public async Task<byte[]> GenerateAndReturnExcel<T>(List<T> data)
        {
            var fileName = @"C:\TempFiles\Temp.xlsx";
            await GenerateExcel(data, fileName);

            return await File.ReadAllBytesAsync(fileName);
        }

        public async Task GenerateExcel<T>(List<T> data, string fileName)
        {
            if (data.Count == 0)
                throw new ArgumentException("List data is empty");
            if (fileName == null)
                throw new ArgumentNullException(nameof(fileName));

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(fileName);

            DeleteIfFileExists(file);

            using var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Main");
            var range = ws.Cells["A1"].LoadFromCollection(data, true);
            range.AutoFitColumns();

            ws.Row(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Row(1).Style.Font.Bold = true;

            await package.SaveAsync();

            _logger.LogInformation($"Excel created Successfully at location {fileName}");
        }

        private void DeleteIfFileExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
                _logger.LogInformation($"Previous file was deleted successfully");
            }
        }
    }
}