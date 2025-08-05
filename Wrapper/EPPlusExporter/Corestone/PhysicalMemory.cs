using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace OfficeOpenXml;

[ExcludeFromCodeCoverage]
public class PhysicalMemory{
    public static ulong GetTotalBytes(){
        if (IsUnix()){
            return GetUnixTotal();
        }

        return GetWindowsTotal();
    }

    private static bool IsUnix(){
        var isUnix = RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ||
                     RuntimeInformation.IsOSPlatform(OSPlatform.Linux);

        return isUnix;
    }

    private static ulong GetWindowsTotal(){
        var output = "";

        var info = new ProcessStartInfo{
            FileName = "wmic",
            Arguments = "OS get TotalVisibleMemorySize /Value",
            RedirectStandardOutput = true,
            CreateNoWindow = true,
            UseShellExecute = false
        };

        using (var process = Process.Start(info)){
            output = process.StandardOutput.ReadToEnd();
        }

        var lines = output.Trim().Split("\n");
        var totalMemoryParts = lines[0].Split("=", StringSplitOptions.RemoveEmptyEntries);

        // KB => Byte
        return ulong.Parse(totalMemoryParts[1]) * 1024;
    }

    private static ulong GetUnixTotal(){
        var output = "";

        var info = new ProcessStartInfo("free -m"){
            FileName = "/bin/bash",
            Arguments = "-c \"free -b\"",
            RedirectStandardOutput = true,
            CreateNoWindow = true,
            UseShellExecute = false
        };

        using (var process = Process.Start(info)){
            output = process.StandardOutput.ReadToEnd();
        }

        var lines = output.Split("\n");
        var memory = lines[1].Split(" ", StringSplitOptions.RemoveEmptyEntries);

        return ulong.Parse(memory[1]);
    }
}