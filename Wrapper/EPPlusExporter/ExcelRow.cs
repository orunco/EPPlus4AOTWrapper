using System.Runtime.InteropServices;

// ReSharper disable CheckNamespace
namespace OfficeOpenXml;

public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRow_GetRow))]
    public static int ExcelRow_GetRow(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelRow;
            ArgumentNullException.ThrowIfNull(_this);

            return _this.Row;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return -1;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRow_GetStyle))]
    public static IntPtr ExcelRow_GetStyle(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelRow;
            ArgumentNullException.ThrowIfNull(_this);

            var result = GCHandle.ToIntPtr(GCHandle.Alloc(_this.Style));
            
            return result;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }
}