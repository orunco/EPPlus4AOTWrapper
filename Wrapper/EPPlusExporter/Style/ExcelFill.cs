using System.Runtime.InteropServices;
using OfficeOpenXml.Style;

// ReSharper disable CheckNamespace
namespace OfficeOpenXml;

public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelFill_SetPatternType))]
    public static void ExcelFill_SetPatternType(IntPtr thisIntPtr, ExcelFillStyle patternType){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelFill;
            ArgumentNullException.ThrowIfNull(_this);

            _this.PatternType = patternType;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelFill_GetBackgroundColor))]
    public static IntPtr ExcelFill_GetBackgroundColor(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelFill;
            ArgumentNullException.ThrowIfNull(_this);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.BackgroundColor));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }
}