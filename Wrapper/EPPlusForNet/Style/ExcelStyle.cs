using System.Runtime.InteropServices;

namespace OfficeOpenXml.Style;

public class ExcelStyle(IntPtr _this) : ExcelBase(_this){
    public ExcelFont Font{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelStyle_GetFont(IntPtr thisHandle);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                var result = new ExcelFont(ExcelStyle_GetFont(ThisSafeHandle.DangerousGetHandle()));
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

    public ExcelFill Fill{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelStyle_GetFill(IntPtr thisHandle);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                var result = new ExcelFill(
                    ExcelStyle_GetFill(ThisSafeHandle.DangerousGetHandle()));
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

    public ExcelBorder Border{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelStyle_GetBorder(IntPtr thisHandle);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                var result = new ExcelBorder(
                    ExcelStyle_GetBorder(ThisSafeHandle.DangerousGetHandle()));
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

    public bool WrapText{
        set{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelStyle_SetWrapText(IntPtr thisHandle,
                [MarshalAs(UnmanagedType.Bool)] bool isWrapText);

#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();

                ExcelStyle_SetWrapText(ThisSafeHandle.DangerousGetHandle(), value);
                ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
            }
            finally{
                if (addedRef){
                    ThisSafeHandle.DangerousRelease();
                }
            }
        }
    }

    public ExcelHorizontalAlignment HorizontalAlignment{
        set{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelStyle_SetHorizontalAlignment(IntPtr thisHandle,
                ExcelHorizontalAlignment horizontalAlignment);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                ExcelStyle_SetHorizontalAlignment(ThisSafeHandle.DangerousGetHandle(), value);
                ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
            }
            finally{
                if (addedRef){
                    ThisSafeHandle.DangerousRelease();
                }
            }
        }
    }

    public ExcelVerticalAlignment VerticalAlignment{
        set{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelStyle_SetVerticalAlignment(IntPtr thisHandle,
                ExcelVerticalAlignment verticalAlignment);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();

                ExcelStyle_SetVerticalAlignment(ThisSafeHandle.DangerousGetHandle(), value);
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