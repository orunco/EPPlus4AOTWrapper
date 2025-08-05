using System.Runtime.InteropServices;
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml;

public class ExcelWorksheet(IntPtr _this) : ExcelBase(_this){
    public string Name{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelWorkSheet_GetName(IntPtr thisHandle);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);
                ExceptionForNet.I.ClearExceptionCache();
                var hGlobalIntPtr = ExcelWorkSheet_GetName(ThisSafeHandle.DangerousGetHandle());
                ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();

                ArgumentNullException.ThrowIfNull(hGlobalIntPtr);

                string? str = Marshal.PtrToStringUni(hGlobalIntPtr);
                if (str != null){
                    ExcelSafeHandle.ReleaseHGlobal(hGlobalIntPtr);
                }

                ArgumentNullException.ThrowIfNull(str);
                return str;
            }
            finally{
                if (addedRef){
                    ThisSafeHandle.DangerousRelease();
                }
            }
        }
    }

    public ExcelWorksheetView View{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelWorkSheet_GetView(IntPtr thisHandle);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                var result = new ExcelWorksheetView(
                    ExcelWorkSheet_GetView(ThisSafeHandle.DangerousGetHandle()));
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

    public ExcelDrawings Drawings{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelWorkSheet_GetDrawings(IntPtr thisHandle);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                var result = new ExcelDrawings(
                    ExcelWorkSheet_GetDrawings(ThisSafeHandle.DangerousGetHandle()));
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


    public ExcelRow Row(int row){
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern IntPtr ExcelWorkSheet_Row(IntPtr thisHandle, int row);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
        ThrowIfDisposed();

        var addedRef = false;
        try{
            ThisSafeHandle.DangerousAddRef(ref addedRef);

            ExceptionForNet.I.ClearExceptionCache();
            var result = new ExcelRow(
                ExcelWorkSheet_Row(ThisSafeHandle.DangerousGetHandle(), row));
            ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();

            return result;
        }
        finally{
            if (addedRef){
                ThisSafeHandle.DangerousRelease();
            }
        }
    }

    public ExcelColumn Column(int col){
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern IntPtr ExcelWorkSheet_Column(IntPtr thisHandle, int col);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
        ThrowIfDisposed();

        var addedRef = false;
        try{
            ThisSafeHandle.DangerousAddRef(ref addedRef);

            ExceptionForNet.I.ClearExceptionCache();
            var result = new ExcelColumn(
                ExcelWorkSheet_Column(ThisSafeHandle.DangerousGetHandle(), col));
            ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();

            return result;
        }
        finally{
            if (addedRef){
                ThisSafeHandle.DangerousRelease();
            }
        }
    }

    public ExcelRange Cells{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelWorkSheet_GetCells(IntPtr thisHandle);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                var result = new ExcelRange(
                    ExcelWorkSheet_GetCells(ThisSafeHandle.DangerousGetHandle()));
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