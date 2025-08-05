using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace OfficeOpenXml;

public class ExcelPackage : IDisposable{
    private readonly ExcelSafeHandle ThisSafeHandle;

    public const int MaxColumns = 16384;
    public const int MaxRows = 1048576;

    public ExcelPackage(){
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern IntPtr ExcelPackage_Ctor();

#if AccessViolationException
        ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
        // 测试发现因为顺序的问题, 如果仅仅是static ExcelException()是不行的
        // 后一步触发，所以还是需要主动设置，这个类是入口类，就在这里进行设置
        ExceptionForNet.I.Init();

        ExceptionForNet.I.ClearExceptionCache();

        ThisSafeHandle = new ExcelSafeHandle(ExcelPackage_Ctor());

        ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
    }

    public ExcelWorkbook Workbook{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelPackage_GetWorkbook(IntPtr thisHandle);

#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                var result = new ExcelWorkbook(
                    ExcelPackage_GetWorkbook(ThisSafeHandle.DangerousGetHandle()));
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

    public void SaveAs(string name){
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern void ExcelPackage_SaveAs(IntPtr thisHandle,
            [MarshalAs(UnmanagedType.LPWStr)] string name);


#if AccessViolationException
        ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
        ThrowIfDisposed();

        var addedRef = false;
        try{
            ThisSafeHandle.DangerousAddRef(ref addedRef);

            ExceptionForNet.I.ClearExceptionCache();
            ExcelPackage_SaveAs(ThisSafeHandle.DangerousGetHandle(), name);
            ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
        }
        finally{
            if (addedRef){
                ThisSafeHandle.DangerousRelease();
            }
        }

        // 外层会根据业务逻辑主动调用dispose以及内存GC，saveAs不要混合这个操作
    }

    public void MemoryRelease(){
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern void LibraryMemoryRelease();
        
        LibraryMemoryRelease();
    }

    #region dispose

    // ------------------------------------------------------------------------
    private bool _disposed;

    ~ExcelPackage() => Dispose(false);

    public void Dispose(){
        // 创建的顺序是先父再子，倒过来就是先子对象再父对象
        // Cascade dispose calls
        // _bar.Dispose();

        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing){
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern void ExcelPackage_Dispose(IntPtr thisHandle);
#if AccessViolationException
        ExceptionForNet.CallLibraryMethod(null);
#endif
        if (!_disposed){
            if (disposing){
                // dispose managed state (managed objects)

                // entry类，特殊处理
                // 1 先处理业务对象
                var addedRef = false;
                try{
                    ThisSafeHandle.DangerousAddRef(ref addedRef);

                    ExceptionForNet.I.ClearExceptionCache();
                    ExcelPackage_Dispose(ThisSafeHandle.DangerousGetHandle());
                    ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
                }
                finally{
                    if (addedRef){
                        ThisSafeHandle.DangerousRelease();
                    }
                }

                // 2 再处理safehandle
                ThisSafeHandle.Dispose();
            }

            // free unmanaged resources (unmanaged objects) and override finalizer

            // set large fields to null

            _disposed = true;
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void ThrowIfDisposed(){
        ObjectDisposedException.ThrowIf(_disposed, this);
    }

    #endregion
}