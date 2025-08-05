using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Runtime.InteropServices;

namespace OfficeOpenXml;

[ExcludeFromCodeCoverage]
public class LibraryResolver{
    public const string Name = "EPPlus4";

    // 保证全局只加载一次
    static LibraryResolver(){
        NativeLibrary.SetDllImportResolver(
            typeof(LibraryResolver).Assembly,
            ResolveNativeLibrary);
    }

    // libraryName是NativeMethod中DllImport中的名称
    static nint ResolveNativeLibrary(
        string libraryName,
        Assembly assembly,
        DllImportSearchPath? searchPath){
        var extension = GetExtension();

        var libraryFullPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            $"{libraryName}.Native-{RuntimeInformation.RuntimeIdentifier}.{extension}");

        if (!File.Exists(libraryFullPath))
            throw new FileNotFoundException($"{libraryFullPath} library not found");

        return NativeLibrary.Load(libraryFullPath, assembly, searchPath);
    }

    // 事实上，其他平台虽然后缀为so/dylib等，但很多后缀还没有做好准备，而且csproj也是这个
    protected static readonly string LibraryPostfixDefaultDll = "dll";

    private static string GetExtension(){
        return LibraryPostfixDefaultDll;
    }

    public static void Init(){
    }
}