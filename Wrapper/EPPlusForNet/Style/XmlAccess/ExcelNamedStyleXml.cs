using System.Runtime.InteropServices;

namespace OfficeOpenXml.Style.XmlAccess;

public class ExcelNamedStyleXml(IntPtr _this) : ExcelBase(_this){
    public ExcelStyle Style{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelNamedStyleXml_GetStyle(IntPtr thisHandle);

#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                var result = new ExcelStyle(
                    ExcelNamedStyleXml_GetStyle(
                        ThisSafeHandle.DangerousGetHandle()));
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