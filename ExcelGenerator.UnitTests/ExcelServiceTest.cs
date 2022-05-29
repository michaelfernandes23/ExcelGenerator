using ExcelGenerator.Services;
using Microsoft.Extensions.Logging;
using Moq;
using Xunit;

namespace ExcelGenerator.UnitTests
{
    public class ExcelServiceTest
    {
        private readonly Mock<ILogger<ExcelService>> logger;
        private readonly ExcelService excelService;

        public ExcelServiceTest()
        {
            logger = new Mock<ILogger<ExcelService>>();

            excelService = new ExcelService(logger.Object);
        }

        [Fact]
        public async Task GenerateAndReturnExcelTest()
        {
            List<PersonModel> personModels = new List<PersonModel>() { new PersonModel() { Name = "Michael" }, new PersonModel() { Name = "Fernandes" } };
            var result = await excelService.GenerateAndReturnExcel(personModels);

            Assert.NotNull(result);
        }

        [Fact]
        public async Task GenerateExcelTest()
        {
            var fileName = @"C:\TempFiles\Temp.xlsx";
            List<PersonModel> personModels = new List<PersonModel>() { new PersonModel() { Name = "Michael" }, new PersonModel() { Name = "Fernandes" } };
            await excelService.GenerateExcel(personModels, fileName);

            Assert.True(true);
        }

        [Fact]
        public async Task GenerateExcelTestIfDataIsNull()
        {
            var fileName = @"C:\TempFiles\Temp.xlsx";
            List<PersonModel> personModels = new List<PersonModel>();
            await Assert.ThrowsAsync<ArgumentException>(() => excelService.GenerateExcel(personModels, fileName));
        }
    }
}