using System.ComponentModel;
using System.Diagnostics;
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
using Serilog;
using Serilog.Exceptions;

namespace OfficeOpenXml;

//Any problem can be encountered on August 2, 2024
//  https://github.com/dotnet/runtimelab/issues/961 Unable to capture exceptions
//Encountering PalRaiseFailed FastException, upgrading the VNet version or upgrading the business code version
public partial class ExporterExceptionHandler{
    private static readonly object lockObject = new();

    private static readonly Logger _log = LoggerFactory.GetLogger(MethodBase.GetCurrentMethod());

    public static ExporterExceptionHandler Instance{ get; } = new();

    private ExporterExceptionHandler(){
    }

    private delegate void ExceptionCallback(string typeName, string serializedData);

    private static ExceptionCallback? _exceptionCallback;

    public void SetExceptionCallback(IntPtr exceptionHandlerIntPtr){
        lock (lockObject){
            //Modified the global variable 'execeptioncallback' and locked it

            _log.Info($"handler=0x{exceptionHandlerIntPtr:X} ");

            //Only one registration is allowed globally, otherwise the process will exit directly the second time without even a stack
            if (_exceptionCallback == null){
                _log.Info(
                    $"handler=0x{exceptionHandlerIntPtr:X} _exceptionCallback is null, start reg.");

                _exceptionCallback = Marshal.GetDelegateForFunctionPointer
                    <ExceptionCallback>(exceptionHandlerIntPtr);

                //Other abnormalities also need to be captured. Although sparrows are small, they have all five organs
                AppDomain.CurrentDomain.UnhandledException += DomainUnhandledException;
                TaskScheduler.UnobservedTaskException += TaskSchedulerUnobservedTaskException;

                _log.Info($"handler=0x{exceptionHandlerIntPtr:X} reg ok.");
            }

            _log.Info($"handler=0x{exceptionHandlerIntPtr:X} finish");
        }
    }

    public void OnException(
        IntPtr thisIntPtr,
        Exception exception,
        [CallerMemberName] string caller = ""){
        _log.Info($"{caller}() handler=0x{thisIntPtr:X} throw exception");
        _log.Info(exception);

        //Prevent AccessViolationException, even without a stack; If the caller forgets to register
        if (_exceptionCallback == null){
            return;
        }

        _log.Info($"{caller}() handler=0x{thisIntPtr:X} _exceptionCallback not null, call it.");

        //https://stackoverflow.com/questions/8315592/implications-of-throwing-exception-in-delegate-of-unmanaged-callback
        //If there is an exception in this callback function, the native code (i.e. this DLL) will crash due to an unprocessed exception and the program will terminate
        //Finally, PalRaiseFailed FastException was triggered, causing the process to exit
        //So, function exceptions can only be handled through return values, similar to traditional return values or FHIR astException
        try{
            _exceptionCallback(
                exception.GetType().ToString(),
                Serialize(exception));
        }
        catch (Exception e){
            _log.Info($"{caller}() handler=0x{thisIntPtr:X} " +
                      $"throw inner exception: {Environment.NewLine}" +
                      $"{e}");
        }
    }

    private static string Serialize(Exception exception){
        var typeInfo = JsonContext.Default.GetTypeInfo(exception.GetType());

        return typeInfo == null
            ? JsonSerializer.Serialize(exception, JsonContext.Default.Exception)
            : JsonSerializer.Serialize(exception, typeInfo); //Falling back to the most basic type
    }

    //20240727 I don't know why, that's all for now. Bin serialization+hot won't work anymore, JSON serialization+hot won+source code generator can,
    //But it requires brute force enumeration of all types and arrangement
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

    //The main class must be partial for the following class to be correct Source code generator, there is no prompt for this
    internal partial class JsonContext : JsonSerializerContext{
    }

    //For unit testing, throwing exceptions without registration will not result in errors
    [UnmanagedCallersOnly(EntryPoint = nameof(ExceptionHandler_LoadLib_ForTest))]
    public static int ExceptionHandler_LoadLib_ForTest(){
#if AccessViolationException
        Exporter.EnterExportMethod(IntPtr.Zero);
#endif

        try{
            return 1;
        }
        catch (Exception e){
            Instance.OnException(IntPtr.Zero, e);
        }

        return 0;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExceptionHandler_ThrowException_ForTest))]
    public static void ExceptionHandler_ThrowException_ForTest(IntPtr messageIntPtr){
#if AccessViolationException
        Exporter.EnterExportMethod(messageIntPtr);
#endif

        try{
            var message = Marshal.PtrToStringUni(messageIntPtr);
            throw new Exception(message);
        }
        catch (Exception e){
            Instance.OnException(IntPtr.Zero, e);
        }
    }
}