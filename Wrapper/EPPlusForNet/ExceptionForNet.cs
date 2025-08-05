using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Diagnostics.Tracing;
using System.Globalization;
using System.Reflection;
using System.Resources;
using System.Runtime;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeOpenXml.Corestone;
using InvalidOperationException = System.InvalidOperationException;
using Logger = OfficeOpenXml.Corestone.Logger;

namespace OfficeOpenXml;

// https://learn.microsoft.com/en-us/cpp/dotnet/how-to-marshal-callbacks-and-delegates-by-using-cpp-interop?view=msvc-170
// azure-functions-dotnet-worker NativeWorkerClient

public partial class ExceptionForNet : IDisposable{
    private static readonly object lockObject = new();

    private static readonly Logger log = LoggerFactory.GetLogger(MethodBase.GetCurrentMethod());

    public static ExceptionForNet I{ get; } = new();

    private ExceptionForNet(){
    }

    private delegate void ExceptionCallback(string typeName, string serializedData);

    static readonly ExceptionCallback _exceptionCallback = (OnException);

    private ThreadLocal<(string?, string?)> ExceptionCache = new();

    private GCHandle _gcHandle = default;

    static ExceptionForNet(){
        // Process terminated. A callback was made on a garbage collected delegate of type 
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern void InitLibrary(IntPtr exceptionHandlerIntPtr);

        lock (lockObject){
            // 修改了全局变量_gcHandle、ExceptionInfo， 锁
#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(IntPtr.Zero);
#endif
            LibraryResolver.Init();

            // 防止GC移动
            I._gcHandle = GCHandle.Alloc(_exceptionCallback);

            InitLibrary(Marshal.GetFunctionPointerForDelegate(_exceptionCallback));
        }
    }

    // 必须要主动设置
    public void Init(){
    }

    // 多线程调用没有问题，是ThreadLocal变量
    public void ClearExceptionCache(){
        lock (lockObject){
            // 修改了全局变量 ExceptionCache， 锁
            ExceptionCache.Value = (null, null);
        }
    }

    // C# 的 rethrow机制，如果是throw e;默认会破坏原始堆栈，dll内部的堆栈就看不到了，
    // 所以这里要阻止原始堆栈被破坏 没想到这个函数沉睡了这样久，在这里用上了
    private static void PreserveStackTrace(Exception exception){
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning disable CS8602 // Dereference of a possibly null reference.

        MethodInfo preserveStackTrace = typeof(Exception).GetMethod(
            "InternalPreserveStackTrace",
            BindingFlags.Instance | BindingFlags.NonPublic);
        preserveStackTrace.Invoke(exception, null);

#pragma warning restore CS8602 // Dereference of a possibly null reference.
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
    }

    // 忠实还原，力求保持高度一致，
    // 这个函数是在dll中执行的，采用函数指针的方法函数签名必须加上UnmanagedCallersOnly
    // 因为这个函数确实是dll中执行的，很有道理啊
    static void OnException(string typeName, string serializedData){
        lock (lockObject){
            // 修改了全局变量 ExceptionCache， 锁
            I.ExceptionCache.Value = (typeName, serializedData);
        }
    }

    [ExcludeFromCodeCoverage]
    public void ThrowExceptionWhenCacheHasValue(){
        lock (lockObject){
            // 修改了全局变量 ExceptionCache， 锁
            if (ExceptionCache.Value.Item1 == null || ExceptionCache.Value.Item2 == null){
                return; //无异常发生
            }

            var typeName = new string(ExceptionCache.Value.Item1);
            var serializedData = new string(ExceptionCache.Value.Item2);

            // 上面已经获取到了异常，为了防止多线程或者多个函数的相互干扰，比如测试用例同时测试
            // 这个函数是为了配合ClearExceptionCache使用的，为了最大化防止代码没有clear
            ClearExceptionCache();

            var type = Type.GetType(typeName);

            if (type == null){
                throw new Exception(typeName + Environment.NewLine + serializedData);
            }

            // 一次解包
            var layout = JsonSerializer.Deserialize(serializedData,
                JsonContext.Default.ExceptionLayout);

            if (layout == null){
                throw new Exception(typeName + Environment.NewLine + serializedData);
            }

            // 似乎json反序列不太成功，message和StackTrace都不行
            // 二次模拟，最多100次可以找到
            // System.Private.CoreLib, Version=8.0.0.0, Culture=neutral, PublicKeyToken=7cec85d7bea7798e
            if (type.ToString() == typeof(BadImageFormatException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.BadImageFormatException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(TypeLoadException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.TypeLoadException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(AccessViolationException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.AccessViolationException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(AggregateException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.AggregateException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(AppDomainUnloadedException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.AppDomainUnloadedException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(ApplicationException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.ApplicationException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(ArgumentException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.ArgumentException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(ArgumentNullException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.ArgumentNullException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(ArgumentOutOfRangeException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.ArgumentOutOfRangeException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(ArithmeticException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.ArithmeticException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(ArrayTypeMismatchException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.ArrayTypeMismatchException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(CannotUnloadAppDomainException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.CannotUnloadAppDomainException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(ContextMarshalException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.ContextMarshalException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(DataMisalignedException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.DataMisalignedException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(DivideByZeroException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.DivideByZeroException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(DllNotFoundException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.DllNotFoundException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(DuplicateWaitObjectException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.DuplicateWaitObjectException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(EntryPointNotFoundException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.EntryPointNotFoundException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(FieldAccessException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.FieldAccessException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(FormatException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.FormatException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(IndexOutOfRangeException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.IndexOutOfRangeException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(InsufficientExecutionStackException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.InsufficientExecutionStackException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(InsufficientMemoryException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.InsufficientMemoryException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(InvalidCastException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.InvalidCastException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(InvalidOperationException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.InvalidOperationException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(InvalidProgramException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.InvalidProgramException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(InvalidTimeZoneException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.InvalidTimeZoneException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(MemberAccessException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.MemberAccessException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(MethodAccessException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.MethodAccessException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(MissingFieldException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.MissingFieldException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(MissingMemberException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.MissingMemberException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(MissingMethodException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.MissingMethodException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(MulticastNotSupportedException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.MulticastNotSupportedException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(NotFiniteNumberException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.NotFiniteNumberException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(NotImplementedException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.NotImplementedException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(NotSupportedException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.NotSupportedException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(NullReferenceException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.NullReferenceException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(ObjectDisposedException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.ObjectDisposedException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(OperationCanceledException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.OperationCanceledException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(OutOfMemoryException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.OutOfMemoryException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(OverflowException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.OverflowException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(PlatformNotSupportedException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.PlatformNotSupportedException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(RankException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.RankException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(StackOverflowException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.StackOverflowException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(SystemException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.SystemException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(TimeoutException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.TimeoutException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(TimeZoneNotFoundException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.TimeZoneNotFoundException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(TypeAccessException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.TypeAccessException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(TypeInitializationException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.TypeInitializationException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(TypeUnloadedException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.TypeUnloadedException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(UnauthorizedAccessException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.UnauthorizedAccessException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(CultureNotFoundException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.CultureNotFoundException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(Win32Exception).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.Win32Exception);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(AbandonedMutexException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.AbandonedMutexException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(LockRecursionException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.LockRecursionException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(SemaphoreFullException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.SemaphoreFullException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(SynchronizationLockException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.SynchronizationLockException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(ThreadAbortException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.ThreadAbortException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(ThreadInterruptedException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.ThreadInterruptedException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(ThreadStartException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.ThreadStartException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(ThreadStateException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.ThreadStateException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(WaitHandleCannotBeOpenedException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.WaitHandleCannotBeOpenedException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(TaskCanceledException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.TaskCanceledException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(TaskSchedulerException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.TaskSchedulerException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(DecoderFallbackException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.DecoderFallbackException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(EncoderFallbackException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.EncoderFallbackException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(SecurityException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.SecurityException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(VerificationException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.VerificationException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(CryptographicException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.CryptographicException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(AmbiguousImplementationException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.AmbiguousImplementationException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(SerializationException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.SerializationException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(COMException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.COMException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(ExternalException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.ExternalException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(InvalidComObjectException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.InvalidComObjectException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(InvalidOleVariantTypeException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.InvalidOleVariantTypeException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(MarshalDirectiveException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.MarshalDirectiveException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(SafeArrayRankMismatchException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.SafeArrayRankMismatchException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(SafeArrayTypeMismatchException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.SafeArrayTypeMismatchException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(SEHException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.SEHException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(RuntimeWrappedException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.RuntimeWrappedException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(SwitchExpressionException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.SwitchExpressionException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(MissingManifestResourceException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.MissingManifestResourceException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(MissingSatelliteAssemblyException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.MissingSatelliteAssemblyException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(AmbiguousMatchException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.AmbiguousMatchException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(CustomAttributeFormatException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.CustomAttributeFormatException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(InvalidFilterCriteriaException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.InvalidFilterCriteriaException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(ReflectionTypeLoadException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.ReflectionTypeLoadException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(TargetException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.TargetException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(TargetInvocationException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.TargetInvocationException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(TargetParameterCountException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.TargetParameterCountException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(FileLoadException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.FileLoadException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(FileNotFoundException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.FileNotFoundException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(DirectoryNotFoundException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.DirectoryNotFoundException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(EndOfStreamException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.EndOfStreamException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(InvalidDataException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.InvalidDataException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(IOException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.IOException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(PathTooLongException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.PathTooLongException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(UnreachableException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.UnreachableException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(EventSourceException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.EventSourceException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }

            if (type.ToString() == typeof(KeyNotFoundException).ToString()){
                var e = JsonSerializer.Deserialize(
                    serializedData, JsonContext.Default.KeyNotFoundException);
                if (e == null)
                    throw new Exception(typeName + Environment.NewLine + serializedData);
                SetMessageStackTrace(e, layout.Message, layout.StackTrace);
                PreserveStackTrace(e);
                throw e;
            }
// EPPlus, Version=4.5.3.3, Culture=neutral, PublicKeyToken=null
// System.Runtime, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a
// System.Console, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a

            // 到这里还没有结果，fallback
            // if (type.ToString() == typeof(Exception).ToString()){
            var exception = JsonSerializer.Deserialize(
                serializedData, JsonContext.Default.Exception);
            if (exception == null)
                throw new Exception(typeName + Environment.NewLine + serializedData);
            SetMessageStackTrace(exception, layout.Message, layout.StackTrace);
            throw exception;
        }
    }


    private static void SetMessageStackTrace(
        Exception exception,
        string message,
        string stackTrace){
        // private string _stackTraceString;
        var _message = typeof(Exception).GetField("_message",
            BindingFlags.NonPublic | BindingFlags.Instance);
        if (_message == null){
            return;
        }

        _message.SetValue(exception, message);


        // private string _stackTraceString;
        var _stackTraceString = typeof(Exception).GetField("_stackTraceString",
            BindingFlags.NonPublic | BindingFlags.Instance);
        if (_stackTraceString == null){
            return;
        }

        _stackTraceString.SetValue(exception, stackTrace);
    }

    // EXCEPTION的序列化都OK，反序列化有点问题：只能设置message，其他的都没有办法设置...
    // 因为setter不允许，这...... 包括innerException
    // 不能改变顺序
    [StructLayout(LayoutKind.Sequential)]
    public class ExceptionLayout{
        public Dictionary<string, string> Data{ get; set; }
        public string? HelpLink{ get; set; }
        public int HResult{ get; set; }
        public ExceptionLayout? InnerException;
        public string Message{ get; set; }
        public string Source{ get; set; }
        public string StackTrace{ get; set; }
        public MethodBase? TargetSite;
    }


    // 解包
    [JsonSerializable(typeof(ExceptionLayout), GenerationMode = JsonSourceGenerationMode.Metadata)]

    // System.Private.CoreLib, Version=8.0.0.0, Culture=neutral, PublicKeyToken=7cec85d7bea7798e
    [JsonSerializable(typeof(BadImageFormatException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(Exception), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(TypeLoadException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(AccessViolationException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(AggregateException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(AppDomainUnloadedException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(ApplicationException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(ArgumentException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(ArgumentNullException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(ArgumentOutOfRangeException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(ArithmeticException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(ArrayTypeMismatchException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(CannotUnloadAppDomainException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(ContextMarshalException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(DataMisalignedException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(DivideByZeroException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(DllNotFoundException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(DuplicateWaitObjectException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(EntryPointNotFoundException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    //[JsonSerializable(typeof(System.ExecutionEngineException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(FieldAccessException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(FormatException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(IndexOutOfRangeException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(InsufficientExecutionStackException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(InsufficientMemoryException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(InvalidCastException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(InvalidOperationException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(InvalidProgramException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(InvalidTimeZoneException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(MemberAccessException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(MethodAccessException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(MissingFieldException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(MissingMemberException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(MissingMethodException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(MulticastNotSupportedException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(NotFiniteNumberException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(NotImplementedException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(NotSupportedException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(NullReferenceException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(ObjectDisposedException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(OperationCanceledException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(OutOfMemoryException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(OverflowException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(PlatformNotSupportedException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(RankException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(StackOverflowException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(SystemException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(TimeoutException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(TimeZoneNotFoundException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(TypeAccessException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(TypeInitializationException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(TypeUnloadedException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(UnauthorizedAccessException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(CultureNotFoundException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(Win32Exception), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(AbandonedMutexException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(LockRecursionException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(SemaphoreFullException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(SynchronizationLockException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(ThreadAbortException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(ThreadInterruptedException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(ThreadStartException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(ThreadStateException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(WaitHandleCannotBeOpenedException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(TaskCanceledException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(TaskSchedulerException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(DecoderFallbackException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(EncoderFallbackException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(SecurityException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(VerificationException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(CryptographicException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(AmbiguousImplementationException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(SerializationException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(COMException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(ExternalException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(InvalidComObjectException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(InvalidOleVariantTypeException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(MarshalDirectiveException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(SafeArrayRankMismatchException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(SafeArrayTypeMismatchException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(SEHException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(RuntimeWrappedException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(SwitchExpressionException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(MissingManifestResourceException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(MissingSatelliteAssemblyException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    //[JsonSerializable(typeof(System.Reflection.MetadataException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(AmbiguousMatchException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(CustomAttributeFormatException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(InvalidFilterCriteriaException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(ReflectionTypeLoadException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(TargetException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(TargetInvocationException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(TargetParameterCountException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(FileLoadException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(FileNotFoundException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(DirectoryNotFoundException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(EndOfStreamException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(InvalidDataException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(IOException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(PathTooLongException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    // [JsonSerializable(typeof(System.Diagnostics.DebugProvider DebugAssertException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(UnreachableException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    // [JsonSerializable(typeof(System.Diagnostics.Contracts.ContractException), GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(EventSourceException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(KeyNotFoundException),
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    // EPPlus, Version=4.5.3.3, Culture=neutral, PublicKeyToken=null
    // System.Runtime, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a
    // System.Console, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a
    internal partial class JsonContext : JsonSerializerContext{
    }

    // ------------------------------------------------------------------------
    private bool _disposed;

    ~ExceptionForNet() => Dispose(false);

    public void Dispose(){
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing){
        if (!_disposed){
            if (disposing){
                // dispose managed state (managed objects)
                if (_gcHandle != default){
                    _gcHandle.Free();
                }
            }

            // free unmanaged resources (unmanaged objects) and override finalizer

            // set large fields to null

            _disposed = true;
        }
    }

    //-------------------------------------------------------------------------
    public int LoadLib_ForTest(){
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern int ExceptionHandler_LoadLib_ForTest();

        return ExceptionHandler_LoadLib_ForTest();
    }

    public void ThrowException_ForTest(string message){
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern void ExceptionHandler_ThrowException_ForTest(
            [MarshalAs(UnmanagedType.LPWStr)] string message);

        I.ClearExceptionCache();
        ExceptionHandler_ThrowException_ForTest(message);
        I.ThrowExceptionWhenCacheHasValue();
    }
}