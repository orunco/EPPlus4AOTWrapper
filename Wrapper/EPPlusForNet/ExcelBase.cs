using System.Runtime.CompilerServices;

namespace OfficeOpenXml;

// 20240822 定位了半天，这个测试用例会卡主，仔细看了代码，也许是因为指针是本地的
// 但是ExcelSafeHandle释放资源时会自动发到dll，对端没有这个指针但是释放了
// 内存结构微乱 所以把ExcelBase改成abstract
public abstract class ExcelBase(IntPtr _this) : IDisposable{
    protected readonly ExcelSafeHandle ThisSafeHandle = new(_this);

    // ------------------------------------------------------------------------
    private bool _disposed;

    ~ExcelBase() => Dispose(false);

    public void Dispose(){
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing){
        if (!_disposed){
            if (disposing){
                // dispose managed state (managed objects)
                ThisSafeHandle.Dispose();
            }

            // free unmanaged resources (unmanaged objects) and override finalizer

            // set large fields to null

            _disposed = true;
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    protected void ThrowIfDisposed(){
        ObjectDisposedException.ThrowIf(_disposed, this);
    }
}