using System.Runtime.InteropServices;

// ReSharper disable CheckNamespace
namespace OfficeOpenXml;

public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelWorkbookView_SetActiveTab))]
    public static void ExcelWorkbookView_SetActiveTab(IntPtr thisIntPtr, int activeTab){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelWorkbookView;
            ArgumentNullException.ThrowIfNull(_this);

            _this.ActiveTab = activeTab;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }
}