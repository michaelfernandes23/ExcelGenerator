using Microsoft.Extensions.Logging;

namespace ExcelGenerator.Services
{
    public class ExcelService : IExcelService
    {
        private readonly ILogger<ExcelService> _logger;

        public ExcelService(ILogger<ExcelService> logger)
        {
            _logger = logger;
        }

        public async Task<string> GenerateExcel()
        {
            return string.Empty;
        }
    }
}