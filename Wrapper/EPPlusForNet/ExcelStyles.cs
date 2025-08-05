using System.Runtime.InteropServices;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml;

public class ExcelStyles(IntPtr _this) : ExcelBase(_this){
    public ExcelNamedStyleXml CreateNamedStyle(string name){
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern IntPtr ExcelStyles_CreateNamedStyle(IntPtr thisHandle,
            [MarshalAs(UnmanagedType.LPWStr)] string name);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
        ThrowIfDisposed();

        var addedRef = false;
        try{
            ThisSafeHandle.DangerousAddRef(ref addedRef);

            ExceptionForNet.I.ClearExceptionCache();
            var result = new ExcelNamedStyleXml(ExcelStyles_CreateNamedStyle(
                ThisSafeHandle.DangerousGetHandle(), name));
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