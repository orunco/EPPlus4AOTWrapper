using System.Drawing;
using System.Runtime.InteropServices;

namespace OfficeOpenXml.Style;

public class ExcelRichText(IntPtr _this) : ExcelBase(_this){
    public bool Italic{
        set{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern void ExcelRichText_SetItalic(IntPtr thisHandle,
                [MarshalAs(UnmanagedType.Bool)] bool isItalic);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();

                ExcelRichText_SetItalic(ThisSafeHandle.DangerousGetHandle(), value);
                ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
            }
            finally{
                if (addedRef){
                    ThisSafeHandle.DangerousRelease();
                }
            }
        }
    }

    public Color Color{
        set{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern void ExcelRichText_SetColor(IntPtr thisHandle,
                Int32 colorArgb);

#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();

                // 转为rgb
                ExcelRichText_SetColor(ThisSafeHandle.DangerousGetHandle(), value.ToArgb());
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