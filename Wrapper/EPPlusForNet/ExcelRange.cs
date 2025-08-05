using System.Runtime.InteropServices;
using OfficeOpenXml.Style;

namespace OfficeOpenXml;

public class ExcelRange(IntPtr _this) : ExcelBase(_this){
    public ExcelRange this[string Address]{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelRange_Index_Address(IntPtr thisHandle,
                [MarshalAs(UnmanagedType.LPWStr)] string address);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                var result = new ExcelRange(
                    ExcelRange_Index_Address(ThisSafeHandle.DangerousGetHandle(), Address));
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

    // 这个函数调用次数太多了，导致CheckException次数多，所以修改函数导出接口降低检查次数
    // out IntPtr的对端是IntPtr*, unsafe代码还是少写为妙，不易全面铺开
    // 全局唯一一个out出去的写法 ref 
    public ExcelRange this[int Row, int Col]{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            [return: MarshalAs(UnmanagedType.Bool)]
            static extern bool ExcelRange_Index_Row_Col(IntPtr thisHandle,
                int Row, int Col, out IntPtr resultIntPtr);

#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                if (ExcelRange_Index_Row_Col(ThisSafeHandle.DangerousGetHandle(),
                        Row, Col, out IntPtr resultIntPtr)){
                    return new ExcelRange(resultIntPtr);
                }
                else{
                    ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
                }
            }
            finally{
                if (addedRef){
                    ThisSafeHandle.DangerousRelease();
                }
            }

            throw new Exception("code error"); 
        }
    }

    public ExcelRange this[int FromRow, int FromCol, int ToRow, int ToCol]{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelRange_Index_FRow_FCol_TRow_TCol(IntPtr thisHandle,
                int FromRow, int FromCol, int ToRow, int ToCol);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();


            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                var result = new ExcelRange(ExcelRange_Index_FRow_FCol_TRow_TCol(
                    ThisSafeHandle.DangerousGetHandle(),
                    FromRow, FromCol, ToRow, ToCol));
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

    public bool AutoFilter{
        set{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern void ExcelRange_SetAutoFilter(IntPtr thisHandle,
                [MarshalAs(UnmanagedType.Bool)] bool isAutoFilter);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();


            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                ExcelRange_SetAutoFilter(ThisSafeHandle.DangerousGetHandle(), value);
                ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
            }
            finally{
                if (addedRef){
                    ThisSafeHandle.DangerousRelease();
                }
            }
        }
    }

    public bool Merge{
        set{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern void ExcelRange_SetMerge(IntPtr thisHandle,
                [MarshalAs(UnmanagedType.Bool)] bool isMerge);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();


            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();

                ExcelRange_SetMerge(ThisSafeHandle.DangerousGetHandle(), value);
                ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
            }
            finally{
                if (addedRef){
                    ThisSafeHandle.DangerousRelease();
                }
            }
        }
    }

    public ExcelHyperLink Hyperlink{
        set{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelRange_SetExcelHyperLink(IntPtr thisHandle,
                [MarshalAs(UnmanagedType.LPWStr)] string referenceAddress,
                [MarshalAs(UnmanagedType.LPWStr)] string display);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();


            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();

                ExcelRange_SetExcelHyperLink(ThisSafeHandle.DangerousGetHandle(),
                    value.ReferenceAddress, value.Display);
                ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
            }
            finally{
                if (addedRef){
                    ThisSafeHandle.DangerousRelease();
                }
            }
        }
    }

    public string StyleName{
        set{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern void ExcelRange_SetStyleName(IntPtr thisHandle,
                [MarshalAs(UnmanagedType.LPWStr)] string name);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();


            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                ExcelRange_SetStyleName(ThisSafeHandle.DangerousGetHandle(), value);
                ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
            }
            finally{
                if (addedRef){
                    ThisSafeHandle.DangerousRelease();
                }
            }
        }
    }

    public ExcelRichTextCollection RichText{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelRange_GetRichText(IntPtr thisHandle);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();


            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                var result = new ExcelRichTextCollection(
                    ExcelRange_GetRichText(ThisSafeHandle.DangerousGetHandle()));
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

    public ExcelStyle Style{
        get{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern IntPtr ExcelRange_GetStyle(IntPtr thisHandle);


#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();


            var addedRef = false;
            try{
                ThisSafeHandle.DangerousAddRef(ref addedRef);

                ExceptionForNet.I.ClearExceptionCache();
                var result = new ExcelStyle(ExcelRange_GetStyle(ThisSafeHandle.DangerousGetHandle()));
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

    /*
    这个设置是性能的关键所在
    https://learn.microsoft.com/en-us/dotnet/framework/interop/copying-and-pinning 
    https://learn.microsoft.com/en-us/dotnet/framework/interop/default-marshalling-for-strings
    When the CharSet is Unicode or a string argument is explicitly marked as 
    [MarshalAs(UnmanagedType.LPWSTR)] and the string is passed by value (not ref or out), 
    the string is pinned [自动被托管GC] and used directly by native code.
    
    Otherwise, platform invoke copies string arguments, converting from the 
    .NET Framework format (Unicode) to the platform unmanaged format. 
    我们不存在这种情况，这种情况性能很差，因为复制了
    
    Strings are immutable and are not copied back from unmanaged memory
     to managed memory when the call returns.
     
    
    Native code is only responsible for releasing the memory when the 
    string is passed by reference and it assigns a new value. 
    我们不存在这种情况
    
    Otherwise, the .NET runtime owns the memory and will release it after the call.
    也就是宿主GC pin后会自动release
    
    https://learn.microsoft.com/en-us/dotnet/standard/native-interop/best-practices
    String parameters
    A string is pinned and used directly by native code (rather than copied) 
    when passed by value (not ref or out) and any one of the following:
    

     */
    public object Value{
        set{
            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern void ExcelRange_SetValue_String(IntPtr thisHandle,
                [MarshalAs(UnmanagedType.LPWStr)] string value);

            [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
            static extern void ExcelRange_SetValue_Int(IntPtr thisHandle,
                int value);

            // 值有可能为空
            // ReSharper disable ConditionIsAlwaysTrueOrFalseAccordingToNullableAPIContract
            if (value == null){
                return;
            }
            // ReSharper restore ConditionIsAlwaysTrueOrFalseAccordingToNullableAPIContract

#if AccessViolationException
            ExceptionForNet.CallLibraryMethod(ThisSafeHandle);
#endif
            ThrowIfDisposed();

            if (value is string valueStr){
                var addedRef = false;
                try{
                    ThisSafeHandle.DangerousAddRef(ref addedRef);

                    ExceptionForNet.I.ClearExceptionCache();

                    // GC.AddMemoryPressure(); 意义不大
                    ExcelRange_SetValue_String(ThisSafeHandle.DangerousGetHandle(), valueStr);
                    ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
                }
                finally{
                    if (addedRef){
                        ThisSafeHandle.DangerousRelease();
                    }
                }
            }

            else if (value is int intStr){
                var addedRef = false;
                try{
                    ThisSafeHandle.DangerousAddRef(ref addedRef);
                    ExceptionForNet.I.ClearExceptionCache();

                    ExcelRange_SetValue_Int(ThisSafeHandle.DangerousGetHandle(), intStr);
                    ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
                }
                finally{
                    if (addedRef){
                        ThisSafeHandle.DangerousRelease();
                    }
                }
            }
            else{
                // 一律转为string
                var addedRef = false;
                try{
                    ThisSafeHandle.DangerousAddRef(ref addedRef);

                    ExceptionForNet.I.ClearExceptionCache();

#pragma warning disable CS8604 // Possible null reference argument.
                    ExcelRange_SetValue_String(ThisSafeHandle.DangerousGetHandle(), value.ToString());
#pragma warning restore CS8604 // Possible null reference argument.
                    ExceptionForNet.I.ThrowExceptionWhenCacheHasValue();
                }
                finally{
                    if (addedRef){
                        ThisSafeHandle.DangerousRelease();
                    }
                }
                //throw new NotSupportedException(value.GetType().ToString());
            }
        }
    }
}