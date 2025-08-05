using System.Drawing;
using System.Runtime.InteropServices;


namespace OfficeOpenXml.Style;

public class ExcelColor(IntPtr _this) : ExcelBase(_this){
    public void SetColor(Color color){
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern void ExcelColor_SetColor(IntPtr thisHandle,
            int colorArgb);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
        ThrowIfDisposed();

        var addedRef = false;
        try{
            ThisSafeHandle.DangerousAddRef(ref addedRef);

            ExceptionForNet.I.ClearExceptionCache();
            // 转为rgb
            ExcelColor_SetColor(ThisSafeHandle.DangerousGetHandle(), color.ToArgb());
            ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
        }
        finally{
            if (addedRef){
                ThisSafeHandle.DangerousRelease();
            }
        }
    }
}