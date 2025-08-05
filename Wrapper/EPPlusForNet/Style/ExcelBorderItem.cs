using System.Runtime.InteropServices;

namespace OfficeOpenXml.Style;

public class ExcelBorderItem(IntPtr _this) : ExcelBase(_this){
    public ExcelBorderStyle Style{
        set{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern void ExcelBorderItem_SetStyle(IntPtr thisHandle,
                ExcelBorderStyle style);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                ExcelBorderItem_SetStyle(ThisSafeHandle.DangerousGetHandle(), value);
                ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
            }
            finally{
                if (addedRef){
                    ThisSafeHandle.DangerousRelease();
                }
            }
        }
    }

    public ExcelColor Color{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelBorderItem_GetColor(IntPtr thisHandle);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                var result = new ExcelColor(
                    ExcelBorderItem_GetColor(ThisSafeHandle.DangerousGetHandle()));
                ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();

                return result;
            }
            finally{
                if (addedRef){
                    ThisSafeHandle.DangerousRelease();
                }
            }
        }
    }
}