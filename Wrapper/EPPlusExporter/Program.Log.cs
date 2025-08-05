using System.Reflection;
using Serilog;
using Serilog.Exceptions;

namespace OfficeOpenXml;

public partial class Exporter{
    
    private static readonly string _logFile = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, GetLibraryName() + ".log");
    
    private static readonly Logger _log = LoggerFactory.GetLogger(
        MethodBase.GetCurrentMethod());

    private static readonly string template =
        "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} " +
        "{Level:u4} {ProcessId}-{ThreadId} {Message:lj}{NewLine}{Exception}";

    private static void InitGlobalLog(){
        
        // init log
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .Enrich.FromLogContext()
            .Enrich.WithThreadId()
            .Enrich.WithProcessId()
            .Enrich.WithExceptionDetails()
            .WriteTo.Console(outputTemplate: template)
            .WriteTo.PersistentFile(
                _logFile,
                outputTemplate: template,
                shared: true,
                rollOnFileSizeLimit: true,
                fileSizeLimitBytes: 20 * 1024 * 1024,
                retainedFileCountLimit: 2,
                flushToDiskInterval: TimeSpan.FromSeconds(3),
                preserveLogFilename: true)
            .CreateLogger();
        
        _log.Info($"Log file path: {_logFile}");
    }
}