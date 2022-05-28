
namespace ExcelGenerator.Services
{
    public interface IExcelService
    {
        Task<string> GenerateExcel();
    }
}