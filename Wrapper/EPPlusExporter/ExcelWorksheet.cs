using System.Runtime.InteropServices;

namespace OfficeOpenXml;

public partial class Exporter{
    // Here returns a string, and the other end will actively make a call to release it again
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelWorkSheet_GetName))]
    public static IntPtr ExcelWorkSheet_GetName(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelWorksheet;
            ArgumentNullException.ThrowIfNull(_this);

            if (_this.Name.Contains('\0')){
                throw new Exception("this.Name.Contains \0");
            }

            // dotnet-netcore-runtime-8.0.4\src\libraries\System.Runtime.InteropServices\tests\System.Runtime.InteropServices.UnitTests\
            // System\Runtime\InteropServices\Marshal\StringMarshalingTests.cs
            //The string return is not gchandle or Marshal PtrToScriptUni appears in pairs
            //If the string is returned directly, the process will exit directly
            return Marshal.StringToHGlobalUni(_this.Name);
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelWorkSheet_GetView))]
    public static IntPtr ExcelWorkSheet_GetView(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelWorksheet;
            ArgumentNullException.ThrowIfNull(_this);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.View));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelWorkSheet_GetDrawings))]
    public static IntPtr ExcelWorkSheet_GetDrawings(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelWorksheet;
            ArgumentNullException.ThrowIfNull(_this);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Drawings));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelWorkSheet_Row))]
    public static IntPtr ExcelWorkSheet_Row(IntPtr thisIntPtr, int row){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelWorksheet;
            ArgumentNullException.ThrowIfNull(_this);

            var result = GCHandle.ToIntPtr(GCHandle.Alloc(_this.Row(row)));
            
            return result;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelWorkSheet_Column))]
    public static IntPtr ExcelWorkSheet_Column(IntPtr thisIntPtr, int col){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelWorksheet;
            ArgumentNullException.ThrowIfNull(_this);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Column(col)));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelWorkSheet_GetCells))]
    public static IntPtr ExcelWorkSheet_GetCells(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as ExcelWorksheet;
            ArgumentNullException.ThrowIfNull(_this);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Cells));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }
}