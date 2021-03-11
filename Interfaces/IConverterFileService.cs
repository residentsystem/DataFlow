using OfficeOpenXml;

namespace DataFlow.Interfaces
{
    public interface IConverterFileService
    {
        void ConvertToCsv(ExcelWorksheet worksheet, int rowcount, int colcount);
    }
}