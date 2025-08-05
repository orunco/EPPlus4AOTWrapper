using System.Runtime.InteropServices;

namespace OfficeOpenXml;

public class ExcelWorksheets(IntPtr _this) : ExcelBase(_this){
    public ExcelWorksheet Add(string name){
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern IntPtr ExcelWorksheets_Add(IntPtr thisHandle,
            [MarshalAs(UnmanagedType.LPWStr)] string name);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
        ThrowIfDisposed();

        var addedRef = false;
        try{
            ThisSafeHandle.DangerousAddRef(ref addedRef);

            ExceptionForNet.I.ClearExceptionCache();
            var result = new ExcelWorksheet(
                ExcelWorksheets_Add(ThisSafeHandle.DangerousGetHandle(), name));
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