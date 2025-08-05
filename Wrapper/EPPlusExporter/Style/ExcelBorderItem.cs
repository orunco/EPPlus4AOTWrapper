using System.Runtime.InteropServices;
using OfficeOpenXml.Style;

// ReSharper disable CheckNamespace
namespace OfficeOpenXml;

public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelBorderItem_SetStyle))]
    public static void ExcelBorderItem_SetStyle(IntPtr thisIntPtr, ExcelBorderStyle style){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelBorderItem;
            ArgumentNullException.ThrowIfNull(_this);

            _this.Style = style;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelBorderItem_GetColor))]
    public static IntPtr ExcelBorderItem_GetColor(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelBorderItem;
            ArgumentNullException.ThrowIfNull(_this);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Color));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }
}