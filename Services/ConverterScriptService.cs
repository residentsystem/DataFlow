using OfficeOpenXml;
using DataFlow.Interfaces;

namespace DataFlow.Services
{
    class ConverterScriptService 
    {
        IConverterScriptService _converterscriptservice;

        public ConverterScriptService(IConverterScriptService converterscriptservice)
        {
            _converterscriptservice = converterscriptservice;
        }

        public void CreateScript(ExcelWorksheet worksheet, int rowcount, int colcount)
        {
            _converterscriptservice.ConvertToScript(worksheet, rowcount, colcount);
            _converterscriptservice.ConvertToMultipleScript(worksheet, rowcount, colcount);
            _converterscriptservice.ConvertToServerPool(worksheet, rowcount, colcount);
        }
    }
}