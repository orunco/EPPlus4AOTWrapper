using System.Runtime.InteropServices;

namespace OfficeOpenXml;

// https://stackoverflow.com/questions/54076659/pinning-class-instance-with-gchandle-alloc-in-c-sharp
// because GCHandle will have current addres of your object even if GC would move it ...
// and you would pass GCHandle pointer there not address of your object ...
public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelPackage_Ctor))]
    public static IntPtr ExcelPackage_Ctor(){
#if AccessViolationException
        EnterExportMethod(IntPtr.Zero);
#endif

        try{
            var excelPackage = new ExcelPackage();
            
            var result = GCHandle.ToIntPtr(GCHandle.Alloc(excelPackage));
            
#if AccessViolationException
            _log.Info($"result=0x{result:X} ");
#endif

            return result;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(IntPtr.Zero, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelPackage_GetWorkbook))]
    public static IntPtr ExcelPackage_GetWorkbook(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = ToPackage(thisIntPtr);
            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Workbook));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelPackage_SaveAs))]
    public static void ExcelPackage_SaveAs(IntPtr thisIntPtr, IntPtr nameIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = ToPackage(thisIntPtr);

            var name = Marshal.PtrToStringUni(nameIntPtr);
            ArgumentException.ThrowIfNullOrEmpty(name);

            // EPPLUSE5.3.3 logic need bugfix
            var fileInfo = new FileInfo(name);

            if (File.Exists(fileInfo.FullName)){
                throw new Exception($"File exist {fileInfo.FullName}");
            }

            _this.SaveAs(new FileInfo(name));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelPackage_Dispose))]
    public static void ExcelPackage_Dispose(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var gcHandle = GCHandle.FromIntPtr(thisIntPtr);
            var _this = gcHandle.Target as ExcelPackage;
 
            _this?.Dispose();
            _log.Info($"After dispose, {MemoryHelper.GetProcessReadableMemoryMB()}");
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }

    private static ExcelPackage ToPackage(IntPtr excelPackageIntPtr){
        var _this = GCHandle.FromIntPtr(excelPackageIntPtr).Target as ExcelPackage;
        ArgumentNullException.ThrowIfNull(_this);
        return _this;
    }
}