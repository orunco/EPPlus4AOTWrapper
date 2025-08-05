using System.Runtime.CompilerServices;
using System.Text;
using Serilog;

namespace OfficeOpenXml.Corestone;

//log4j style
public class Logger(string className){
    // CallerFilePath cannot be used for two reasons:
    // one is that it leaks the path information of the development machine
    // (recorded during compilation), and the
    // second is that the file path is not equal to the type path
    // It only needs to support debug, info, warn, and error to simplify the mental burden
    public void Debug(object messageObject,
        [CallerLineNumber] int number = 0,
        [CallerMemberName] string method = ""){
        var messageTemplate = GetMessageTemplate(messageObject);

        Log.Debug(ConcatSource(number, method) + messageTemplate);
    }

    public void Info(object messageObject,
        [CallerLineNumber] int number = 0,
        [CallerMemberName] string method = ""){
        var messageTemplate = GetMessageTemplate(messageObject);

        Log.Information(ConcatSource(number, method) + messageTemplate);
    }

    public void Warn(object messageObject,
        [CallerLineNumber] int number = 0,
        [CallerMemberName] string method = ""){
        var messageTemplate = GetMessageTemplate(messageObject);

        Log.Warning(ConcatSource(number, method) + messageTemplate);
    }

    public void Error(object messageObject,
        [CallerLineNumber] int number = 0,
        [CallerMemberName] string method = ""){
        var messageTemplate = GetMessageTemplate(messageObject);

        Log.Error(ConcatSource(number, method) + messageTemplate);
    }

    private string GetMessageTemplate(object? messageObject){
        string messageTemplate = string.Empty;

        if (messageObject is string messageStr){
            messageTemplate = messageStr;
        }
        else if (messageObject is Exception messageException){
            messageTemplate = GetExceptionMsg(messageException);
        }
        else if (messageObject != null){
            string? messageToString = messageObject.ToString();
            if (messageToString != null){
                messageTemplate = messageToString;
            }
        }

        return messageTemplate;
    }

    private string ConcatSource(int number, string method){
        return className + ".cs" + ":" + number + " " + method + "() | ";
    }

    public static string GetExceptionMsg(Exception ex){
        StringBuilder sb = new StringBuilder();

        if (ex.InnerException != null){
            sb.AppendLine(ex.InnerException.GetType().Name);
            sb.AppendLine(ex.InnerException.Message);
            sb.AppendLine(ex.InnerException.StackTrace);
        }
        else{
            sb.AppendLine(ex.GetType().Name);
            sb.AppendLine(ex.Message);
            sb.AppendLine(ex.StackTrace);
        }

        return sb.ToString();
    }
}