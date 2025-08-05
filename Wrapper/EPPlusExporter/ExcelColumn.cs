using System.Runtime.InteropServices;

// ReSharper disable CheckNamespace
namespace OfficeOpenXml;

public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelColumn_SetWidth))]
    public static void ExcelColumn_SetWidth(IntPtr thisIntPtr, double width){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelColumn;
            ArgumentNullException.ThrowIfNull(_this);

            _this.Width = width;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }
}