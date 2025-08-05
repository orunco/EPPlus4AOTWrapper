using System.Runtime.InteropServices;
using OfficeOpenXml.Style;

// ReSharper disable CheckNamespace
namespace OfficeOpenXml;

// name is not ExcelBorder
public partial class Exporter{
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelBorder_GetLeft))]
    public static IntPtr ExcelBorder_GetLeft(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as Border;
            ArgumentNullException.ThrowIfNull(_this);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Left));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelBorder_GetRight))]
    public static IntPtr ExcelBorder_GetRight(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as Border;
            ArgumentNullException.ThrowIfNull(_this);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Right));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelBorder_GetTop))]
    public static IntPtr ExcelBorder_GetTop(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as Border;
            ArgumentNullException.ThrowIfNull(_this);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Top));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelBorder_GetBottom))]
    public static IntPtr ExcelBorder_GetBottom(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as Border;
            ArgumentNullException.ThrowIfNull(_this);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Bottom));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelBorder_GetDiagonal))]
    public static IntPtr ExcelBorder_GetDiagonal(IntPtr thisIntPtr){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as Border;
            ArgumentNullException.ThrowIfNull(_this);

            return GCHandle.ToIntPtr(GCHandle.Alloc(_this.Diagonal));
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }

        return IntPtr.Zero;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelBorder_SetDiagonalDown))]
    public static void ExcelBorder_SetDiagonalDown(IntPtr thisIntPtr, bool isDiagonalDown){
#if AccessViolationException
        EnterExportMethod(thisIntPtr);
#endif

        try{
            var _this = GCHandle.FromIntPtr(thisIntPtr).Target as Border;
            ArgumentNullException.ThrowIfNull(_this);

            _this.DiagonalDown = isDiagonalDown;
        }
        catch (Exception e){
            ExporterExceptionHandler.Instance.OnException(thisIntPtr, e);
        }
    }
}