using System.Drawing;
using System.Runtime.InteropServices;
using OfficeOpenXml.Style;

// ReSharper disable CheckNamespace
namespace OfficeOpenXml;

public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRichText_SetItalic))]
    public static void ExcelRichText_SetItalic(IntPtr thisIntPtr, bool isItalic){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelRichText;
            ArgumentNullException.ThrowIfNull(_this);

            _this.Italic = isItalic;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRichText_SetColor))]
    public static void ExcelRichText_SetColor(IntPtr thisIntPtr, Int32 colorArgb){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelRichText;
            ArgumentNullException.ThrowIfNull(_this);

            _this.Color = Color.FromArgb(colorArgb);
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }
}