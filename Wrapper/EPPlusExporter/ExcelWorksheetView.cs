using System.Runtime.InteropServices;

// ReSharper disable CheckNamespace
namespace OfficeOpenXml;

public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelWorksheetView_FreezePanes))]
    public static void ExcelWorksheetView_FreezePanes(IntPtr thisIntPtr,
        int Row, int Column){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelWorksheetView;
            ArgumentNullException.ThrowIfNull(_this);

            _this.FreezePanes(Row, Column);
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }
}