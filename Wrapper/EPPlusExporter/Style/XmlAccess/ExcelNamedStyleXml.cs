using System.Runtime.InteropServices;
using OfficeOpenXml.Style.XmlAccess;

// ReSharper disable CheckNamespace
namespace OfficeOpenXml;

public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelNamedStyleXml_GetStyle))]
    public static IntPtr ExcelNamedStyleXml_GetStyle(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelNamedStyleXml;
            ArgumentNullException.ThrowIfNull(_this);
            var result = _this.Style;
            return GCHandle.ToIntPtr(GCHandle.Alloc(result));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }
}