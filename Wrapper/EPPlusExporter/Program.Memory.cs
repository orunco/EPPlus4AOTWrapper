using System.Runtime.InteropServices;

namespace OfficeOpenXml;

public partial class Exporter{
    private static void InitGlobalMemory(){
        try{
            // may be exception
            // GCHeapHardLimitPercent is not supported, so please convert it. It is the default and does not need to be set
            ulong totalMemory = PhysicalMemory.GetTotalBytes();
            _log.Info($"            Total Physical Memory = " +
                      $"{MemoryHelper.HumanReadableBytes((long)totalMemory)}");

            // 50% free ，diff and generate to 25%，Byte
            AppContext.SetData("GCHeapHardLimit", (ulong)(totalMemory * 0.25));

            AppContext.SetSwitch("RetainVM", false);
            GC.RefreshMemoryLimit();
        }
        catch (Exception e){
            _log.Info(e);
        }

        _log.Info($"Total library GC Available Memory = " +
                  $"{MemoryHelper.HumanReadableBytes(GC.GetGCMemoryInfo().TotalAvailableMemoryBytes)}");

        _log.Info($"             Total Process Memory = {MemoryHelper.GetProcessReadableMemoryMB()}");
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(LibraryMemoryRelease))]
    public static void LibraryMemoryRelease(){
        MemoryHelper.GC_Collect();
        _log.Info($"After Memory Release, {MemoryHelper.GetProcessReadableMemoryMB()}");
    }
}