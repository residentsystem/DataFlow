using DataFlow.Interfaces;
using Microsoft.Extensions.DependencyInjection;

namespace DataFlow.Converters
{
    static class DependencyInjectionService
    {
        public static void AddServiceCollection()
        {
            // Setup serviceCollection for Dependency Injection
            var serviceProvider = new ServiceCollection()
            .AddSingleton<IConverterScriptService, WindowsConverter>()
            .AddSingleton<IConverterScriptService, LinuxConverter>()
            .AddSingleton<IConverterFileService, CsvConverter>()
            .BuildServiceProvider();
        }
    }
}