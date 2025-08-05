using System.Runtime.InteropServices;

namespace OfficeOpenXml;

public class ExcelWorksheetView(IntPtr _this) : ExcelBase(_this){
    public void FreezePanes(int Row, int Column){
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern void ExcelWorksheetView_FreezePanes(IntPtr thisHandle,
            int Row, int Column);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
        ThrowIfDisposed();

        var addedRef = false;
        try{
            ThisSafeHandle.DangerousAddRef(ref addedRef);

            ExceptionForNet.I.ClearExceptionCache();
            ExcelWorksheetView_FreezePanes(ThisSafeHandle.DangerousGetHandle(), Row, Column);
            ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
        }
        finally{
            if (addedRef){
                ThisSafeHandle.DangerousRelease();
            }
        }
    }
}