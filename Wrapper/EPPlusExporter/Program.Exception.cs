namespace OfficeOpenXml;

public partial class ExporterExceptionHandler{
    
    private void DomainUnhandledException(object sender, UnhandledExceptionEventArgs args){
        var exception = (Exception)args.ExceptionObject;
        _log.Error("Domain Unhandled Exception.");
        _log.Error(exception);
        _log.Error("Runtime terminating:" + args.IsTerminating);
        OnException(IntPtr.Zero, exception);
    }

    private void TaskSchedulerUnobservedTaskException(object? sender,
        UnobservedTaskExceptionEventArgs args){
        try{
            _log.Error($"TaskScheduler Unobserved Task Exception");
            foreach (var innerException in args.Exception.InnerExceptions){
                _log.Error($"innerException: {innerException.GetType().Name}");
                _log.Error(innerException);
                OnException(IntPtr.Zero, innerException);
            }
        }
        catch (Exception e){
            _log.Error(e);
            OnException(IntPtr.Zero, e);
        }

        // Prevent abnormal bubbling
        args.SetObserved();
    }
}