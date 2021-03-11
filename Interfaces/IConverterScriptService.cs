using OfficeOpenXml;

namespace DataFlow.Interfaces
{
    public interface IConverterScriptService
    {
        void ConvertToScript(ExcelWorksheet worksheet, int rowcount, int colcount);

        void ConvertToMultipleScript(ExcelWorksheet worksheet, int rowcount, int colcount); 

        void ConvertToServerPool(ExcelWorksheet worksheet, int rowcount, int colcount);
    }
}