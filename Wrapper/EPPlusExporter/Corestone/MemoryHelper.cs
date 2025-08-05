using System.Diagnostics;
using System.Globalization;

namespace OfficeOpenXml;


public class MemoryHelper{
    public static void GC_Collect(){
        /*
        https://learn.microsoft.com/en-us/dotnet/core/runtime-config/garbage-collector

"System.GC.RetainVM": false               The memory is returned
    System.GC.ConserveMemory          Values between 1 and 9 inclusive, are also valid. The higher the value, 
                                      the more the garbage collector tries to save 
                                      memory, thus keeping the heap small.
"System.GC.HeapHardLimitPercent" 10
"System.GC.HeapHardLimit": 4294967296
It can be used at the same time, whichever is smaller
4G      = 4294967296/1024/1024/1024
2G      = 2147483648
1G      = 1073741824
512MB   =  536870912

         */
        GC.GetTotalMemory(true);  
        
        // GC. SuppressFinalize: Requests the CLR not to call the finalizer
        // of the specified object'. That is, when you manually call the Dispose
        // or Close method to release the unmanaged resource,
        // This method forcibly tells the CLR not to trigger my destructor anymore,
        // otherwise executing the destructor again is equivalent to cleaning up
        // the unmanaged resources again, causing unknown risks.
    }

    public static string GetProcessReadableMemoryMB(){
        var proc = Process.GetCurrentProcess();
        long b = proc.PrivateMemorySize64;
        for (int i = 0; i < 2; i++){
            b /= 1024;
        }
        
        string aotStr = AotHelper.IsAOT() ? "AOT mode" : "JIT mode";

        string bit = !Environment.Is64BitProcess ? "32 bit" : "64 bit";
         
#if DEBUG
        return $"Process {b} MB [{DateTime.Now}   Debug mode | {aotStr} | {bit}] {Environment.OSVersion} ";
#else
        return $"Process {b} MB [{DateTime.Now} Release mode | {aotStr} | {bit}] {Environment.OSVersion} ";
#endif
    }
    
    public static long GetProcessMemoryMB(){
        var proc = Process.GetCurrentProcess();
        long b = proc.PrivateMemorySize64;
        for (int i = 0; i < 2; i++){
            b /= 1024;
        }

        return b;
    }
    
    public static string HumanReadableBytes(long value)
    {
        string suffix;
        double readable;
        switch (Math.Abs(value))
        {
            case >= 0x1000000000000000:
                suffix = "EB";
                readable = value >> 50;
                break;
            case >= 0x4000000000000:
                suffix = "PB";
                readable = value >> 40;
                break;
            case >= 0x10000000000:
                suffix = "TB";
                readable = value >> 30;
                break;
            case >= 0x40000000:
                suffix = "GB";
                readable = value >> 20;
                break;
            case >= 0x100000:
                suffix = "MB";
                readable = value >> 10;
                break;
            case >= 0x400:
                suffix = "KB";
                readable = value;
                break;
            default:
                return value.ToString("0 B");
        }

        return (readable / 1024).ToString("0.## ", CultureInfo.InvariantCulture) + suffix;
    }
}