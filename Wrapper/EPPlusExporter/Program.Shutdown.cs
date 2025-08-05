using Serilog;

namespace OfficeOpenXml;

public partial class Exporter{
    private static void InitGracefulShutdown(){
        AppDomain.CurrentDomain.ProcessExit += (sender, args) => {
            try{
                Log.CloseAndFlush();
            }
            catch{
                // ignored
            }
        };
    }
}