using System.Runtime.InteropServices;
using OfficeOpenXml.Style;

// ReSharper disable CheckNamespace
namespace OfficeOpenXml;

public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelStyle_GetFont))]
    public static IntPtr ExcelStyle_GetFont(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelStyle;
            ArgumentNullException.ThrowIfNull(_this);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Font));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelStyle_GetFill))]
    public static IntPtr ExcelStyle_GetFill(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelStyle;
            ArgumentNullException.ThrowIfNull(_this);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Fill));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelStyle_GetBorder))]
    public static IntPtr ExcelStyle_GetBorder(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelStyle;
            ArgumentNullException.ThrowIfNull(_this);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Border));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelStyle_SetWrapText))]
    public static void ExcelStyle_SetWrapText(IntPtr thisIntPtr, bool isWrapText){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelStyle;
            ArgumentNullException.ThrowIfNull(_this);

            _this.WrapText = isWrapText;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelStyle_SetHorizontalAlignment))]
    public static void ExcelStyle_SetHorizontalAlignment(IntPtr thisIntPtr,
        ExcelHorizontalAlignment horizontalAlignment){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelStyle;
            ArgumentNullException.ThrowIfNull(_this);

            _this.HorizontalAlignment = horizontalAlignment;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelStyle_SetVerticalAlignment))]
    public static void ExcelStyle_SetVerticalAlignment(IntPtr thisIntPtr,
        ExcelVerticalAlignment verticalAlignment){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelStyle;
            ArgumentNullException.ThrowIfNull(_this);

            _this.VerticalAlignment = verticalAlignment;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }
}