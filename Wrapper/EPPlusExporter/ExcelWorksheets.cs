using System.Runtime.InteropServices;

namespace OfficeOpenXml;

public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelWorksheets_Add))]
    public static IntPtr ExcelWorksheets_Add(IntPtr thisIntPtr, IntPtr nameIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelWorksheets;
            ArgumentNullException.ThrowIfNull(_this);

            var name = Marshal.PtrToStringUni(nameIntPtr);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Add(name)));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }
}