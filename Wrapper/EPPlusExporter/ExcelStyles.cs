using System.Runtime.InteropServices;

// ReSharper disable CheckNamespace
namespace OfficeOpenXml;

public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelStyles_CreateNamedStyle))]
    public static IntPtr ExcelStyles_CreateNamedStyle(IntPtr thisIntPtr, IntPtr nameIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelStyles;
            ArgumentNullException.ThrowIfNull(_this);

            var name = Marshal.PtrToStringUni(nameIntPtr);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.CreateNamedStyle(name)));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }
}