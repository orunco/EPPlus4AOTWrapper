using System.Drawing;
using System.Runtime.InteropServices;
using OfficeOpenXml.Style;

// ReSharper disable CheckNamespace
namespace OfficeOpenXml;

public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelColor_SetColor))]
    public static void ExcelColor_SetColor(IntPtr thisIntPtr, int colorArgb){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelColor;
            ArgumentNullException.ThrowIfNull(_this);

            _this.SetColor(Color.FromArgb(colorArgb));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }
}