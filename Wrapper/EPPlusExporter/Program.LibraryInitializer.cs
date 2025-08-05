using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Serilog;
using Serilog.Exceptions;

namespace OfficeOpenXml;

public partial class Exporter{
    private static readonly object lockObject = new();

    private static bool IsInit;

    [UnmanagedCallersOnly(EntryPoint = nameof(InitLibrary))]
    public static void InitLibrary(IntPtr exceptionCallbackIntPtr){
        lock (lockObject){
            //Modified the global variable IsInit and locked it

            //Do not call any reflection related functions here, as exceptions may occur, including printing methods
            EnterExportMethod(exceptionCallbackIntPtr);

            if (IsInit){
                _log.Info("Already inited.");
                return;
            }

            InitGlobalLog();

            InitGracefulShutdown();

            ExporterExceptionHandler.Instance.SetExceptionCallback(exceptionCallbackIntPtr);

            InitGlobalMemory();

            IsInit = true;

            _log.Info($"Load library {Assembly.GetExecutingAssembly().GetName()} ok.");
        }
    }

    public static void EnterExportMethod(IntPtr intPtr, [CallerMemberName] string caller = ""){
        _log.Info($"{caller}() intPtr=0x{intPtr:X} ");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static string GetLibraryName(){
        // EPPlus, Version=4.5.3.3, Culture=neutral, PublicKeyToken=null
        var name = Assembly.GetExecutingAssembly().GetName().ToString();
        var splits = name.Split(new string[]{ ",", "=" }, StringSplitOptions.RemoveEmptyEntries);
        return splits[0];
    }
}