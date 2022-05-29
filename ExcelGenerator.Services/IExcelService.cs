
namespace ExcelGenerator.Services
{
    public interface IExcelService
    {
        Task<byte[]> GenerateAndReturnExcel<T>(List<T> data);
        Task GenerateExcel<T>(List<T> data, string fileName);
    }
}