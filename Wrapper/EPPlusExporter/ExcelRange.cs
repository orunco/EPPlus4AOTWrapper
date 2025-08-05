using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeOpenXml;

public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRange_Index_Address))]
    public static IntPtr ExcelRange_Index_Address(IntPtr thisIntPtr, IntPtr addressStrIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = ToRange(thisIntPtr);

            var address = Marshal.PtrToStringUni(addressStrIntPtr);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this[address]));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRange_Index_Row_Col))]
    public static unsafe bool ExcelRange_Index_Row_Col(IntPtr thisIntPtr, 
        int Row, int Col, IntPtr* resultIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif
        *resultIntPtr = IntPtr.Zero;
        
        try{
            var _this = ToRange(thisIntPtr);

            *resultIntPtr= GCHandle.ToIntPtr(GCHandle.Alloc(_this[Row, Col]));
            return true;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
            return false;
        }
 
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRange_Index_FRow_FCol_TRow_TCol))]
    public static IntPtr ExcelRange_Index_FRow_FCol_TRow_TCol(IntPtr thisIntPtr,
        int FromRow, int FromCol, int ToRow, int ToCol){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = ToRange(thisIntPtr);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this[FromRow, FromCol, ToRow, ToCol]));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRange_SetAutoFilter))]
    public static void ExcelRange_SetAutoFilter(IntPtr thisIntPtr, bool isAutoFilter){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = ToRange(thisIntPtr);

            _this.AutoFilter = isAutoFilter;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRange_SetMerge))]
    public static void ExcelRange_SetMerge(IntPtr thisIntPtr, bool isMerge){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = ToRange(thisIntPtr);

            _this.Merge = isMerge;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }


    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRange_SetExcelHyperLink))]
    public static void ExcelRange_SetExcelHyperLink(IntPtr thisIntPtr,
        IntPtr referenceAddressIntPtr,
        IntPtr displayIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = ToRange(thisIntPtr);

            var referenceAddress = Marshal.PtrToStringUni(referenceAddressIntPtr);
            var display = Marshal.PtrToStringUni(displayIntPtr);

            _this.Hyperlink = new ExcelHyperLink(referenceAddress, display);
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRange_SetStyleName))]
    public static void ExcelRange_SetStyleName(IntPtr thisIntPtr, IntPtr nameIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = ToRange(thisIntPtr);

            var name = Marshal.PtrToStringUni(nameIntPtr);

            _this.StyleName = name;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRange_GetRichText))]
    public static IntPtr ExcelRange_GetRichText(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = ToRange(thisIntPtr);
            var result = _this.RichText;
            return GCHandle.ToIntPtr(GCHandle.Alloc(result));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRange_GetStyle))]
    public static IntPtr ExcelRange_GetStyle(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = ToRange(thisIntPtr);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Style));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }
    
    // https://learn.microsoft.com/en-us/dotnet/framework/interop/copying-and-pinning 
    // Copying and Pinning
    // No matter what, copying is necessary, so ref is not used
    // [SuppressGCTransition] performance ignored
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRange_SetValue_String))]
    public static void ExcelRange_SetValue_String(IntPtr thisIntPtr, IntPtr valueStrIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = ToRange(thisIntPtr);
            
            // In theory, it can be used directly without the need for copying.
            // Actually, because the Value object of epplus must be an internal object
            // CAN NOT EXECUTE : 
            // var s = new Span<char>((void*)valueStrIntPtr, 64);
            // _this.Value = s;
            
            // return Marshal.IsNullOrWin32Atom(ptr) ? (string) null : new string((char*) ptr);
            var valueStr = Marshal.PtrToStringUni(valueStrIntPtr);
       
            _this.Value = valueStr;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelRange_SetValue_Int))]
    public static void ExcelRange_SetValue_Int(IntPtr thisIntPtr, int value){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = ToRange(thisIntPtr);

            _this.Value = value;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }

    // performance
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static ExcelRange ToRange(IntPtr rangeIntPtr){
        var _this = GCHandle.FromIntPtr(rangeIntPtr).Target as ExcelRange;
        ArgumentNullException.ThrowIfNull(_this);
        return _this;
    }
}