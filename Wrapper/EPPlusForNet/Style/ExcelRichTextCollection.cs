using System.Runtime.InteropServices;

namespace OfficeOpenXml.Style;

public class ExcelRichTextCollection(IntPtr _this) : ExcelBase(_this){
    public ExcelRichText Add(string text){
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern IntPtr ExcelRichTextCollection_Add(IntPtr thisHandle,
            [MarshalAs(UnmanagedType.LPWStr)] string text);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
        ThrowIfDisposed();

        var addedRef = false;
        try{
            ThisSafeHandle.DangerousAddRef(ref addedRef);

            ExceptionForNet.I.ClearExceptionCache();
            var result = new ExcelRichText(
                ExcelRichTextCollection_Add(ThisSafeHandle.DangerousGetHandle(), text));
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