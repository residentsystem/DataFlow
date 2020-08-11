using OfficeOpenXml;
using DataFlow.Interfaces;

namespace DataFlow.Services
{
    class ConverterService 
    {
        private IConverterService _converterservice;

        public ConverterService(IConverterService converterservice)
        {
            _converterservice = converterservice;
        }

        public void CreateScript(ExcelWorksheet worksheet, int rowcount, int colcount)
        {
            _converterservice.ConvertToCsv(worksheet, rowcount, colcount);
            _converterservice.ConvertToScript(worksheet, rowcount, colcount);
            _converterservice.ConvertToMultipleScript(worksheet, rowcount, colcount);
            _converterservice.ConvertToServerPool(worksheet, rowcount, colcount);
        }
    }
}