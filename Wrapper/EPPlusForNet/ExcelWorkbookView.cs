using System.Runtime.InteropServices;

namespace OfficeOpenXml;

public class ExcelWorkbookView(IntPtr _this) : ExcelBase(_this){
    public int ActiveTab{
        set{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern void ExcelWorkbookView_SetActiveTab(IntPtr thisHandle,
                int activeTab);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                ExcelWorkbookView_SetActiveTab(ThisSafeHandle.DangerousGetHandle(), value);
                ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
            }
            finally{
                if (addedRef){
                    ThisSafeHandle.DangerousRelease();
                }
            }
        }
    }
}