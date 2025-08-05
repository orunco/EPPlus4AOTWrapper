using System.Runtime.InteropServices;
using OfficeOpenXml.Style;

namespace OfficeOpenXml;

public class ExcelRow(IntPtr _this) : ExcelBase(_this){
    public int Row{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern int ExcelRow_GetRow(IntPtr thisHandle);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);
                ExceptionForNet.I.ClearExceptionCache();
                var result = ExcelRow_GetRow(ThisSafeHandle.DangerousGetHandle());
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

    public ExcelStyle Style{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelRow_GetStyle(IntPtr thisHandle);

#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);
                ExceptionForNet.I.ClearExceptionCache();
                var result = new ExcelStyle(ExcelRow_GetStyle(ThisSafeHandle.DangerousGetHandle()));
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