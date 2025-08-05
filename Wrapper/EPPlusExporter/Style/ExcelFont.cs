using System.Runtime.InteropServices;
using OfficeOpenXml.Style;

// ReSharper disable CheckNamespace
namespace OfficeOpenXml;

public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelFont_SetBold))]
    public static void ExcelFont_SetBold(IntPtr thisIntPtr, bool isBold){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelFont;
            ArgumentNullException.ThrowIfNull(_this);

            _this.Bold = isBold;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelFont_GetColor))]
    public static IntPtr ExcelFont_GetColor(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelFont;
            ArgumentNullException.ThrowIfNull(_this);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Color));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }
}