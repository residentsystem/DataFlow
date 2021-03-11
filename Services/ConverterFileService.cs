using OfficeOpenXml;
using DataFlow.Interfaces;

namespace DataFlow.Services
{
    class ConverterFileService 
    {
        private IConverterFileService _converterfileservice;

        public ConverterFileService(IConverterFileService converterfileservice)
        {
            _converterfileservice = converterfileservice;
        }

        public void CreateFile(ExcelWorksheet worksheet, int rowcount, int colcount)
        {
            _converterfileservice.ConvertToCsv(worksheet, rowcount, colcount);
        }
    }
}