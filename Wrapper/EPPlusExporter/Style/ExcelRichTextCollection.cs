using System.Runtime.InteropServices;
using OfficeOpenXml.Style;

// ReSharper disable CheckNamespace
namespace OfficeOpenXml;

public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRichTextCollection_Add))]
    public static IntPtr ExcelRichTextCollection_Add(IntPtr thisIntPtr, IntPtr textIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelRichTextCollection;
            ArgumentNullException.ThrowIfNull(_this);

            var text = Marshal.PtrToStringUni(textIntPtr);

            var xx = _this.Add(text);
            return GCHandle.ToIntPtr(GCHandle.Alloc(xx));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }
}