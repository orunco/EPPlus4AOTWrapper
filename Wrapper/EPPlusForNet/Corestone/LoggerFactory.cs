using System.Reflection;

namespace OfficeOpenXml.Corestone;

//log4j style
public class LoggerFactory{
    public static Logger GetLogger(MethodBase? m){
        // if (m == null){
        //     throw new Exception("LoggerFactory.GetLogger(): MethodBase.GetCurrentMethod() is null");
        // }
        //
        // if (m.DeclaringType == null){
        //     throw new Exception("LoggerFactory.GetLogger(): DeclaringType is null");
        // }
        //
        // if (m.DeclaringType.FullName == null){
        //     throw new Exception("LoggerFactory.GetLogger(): DeclaringType.FullName is null");
        // }

        if (m != null && m.DeclaringType != null && m.DeclaringType.FullName != null){
            return new Logger(m.DeclaringType.FullName);
        }

        return new Logger("__EMPTY_CLASS__");
    }
}