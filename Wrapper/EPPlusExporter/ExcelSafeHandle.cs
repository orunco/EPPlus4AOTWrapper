using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace OfficeOpenXml;

public partial class Exporter{
    //Generating too many objects at once, free is also very laborious, and the performance is poor,
    //But it cannot be released in batches because the timing is unknown and cannot be accumulated
    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelSafeHandle_FreeHandleInDll))]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public static void ExcelSafeHandle_FreeHandleInDll(IntPtr objIntPtr){
        lock (lockObject){
            //Prevent multiple threads from releasing the same object multiple times

#if AccessViolationException
            EnterExportMethod(objIntPtr);
#endif

            try{
                var gcHandle = GCHandle.FromIntPtr(objIntPtr);

                // this fun call internal static extern void InternalFree(IntPtr handle);
                // A bizarre test case: When the other end sends its own objIntPtr, the memory
                // structure becomes chaotic after free, causing the test case to become stuck and exit
                // Access Violation: Attempted to read or write protected memory
                // so need prepare check
                if (gcHandle.IsAllocated){
                    // so many log
                    // log.Error("gcHandle is not allocated.");
                    // Here, there will be no sword from the previous dynasty killing the officials of this dynasty
                    gcHandle.Free();
                }

                // if (gcHandle.Target == null){
                //     // so many log
                //     // log.Error("gcHandle.Target == null");
                //     return;
                // }
            }
            catch (Exception e){
                // try best
                // An exception has occurred here, indicating a bug in the code that needs to be addressed
                _log.Error(e);
            }
        }
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(ExcelSafeHandle_FreeHGlobalInDll))]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public static void ExcelSafeHandle_FreeHGlobalInDll(IntPtr hGlobalIntPtr){
        lock (lockObject){
            // Prevent multiple threads from releasing the same object multiple times
#if AccessViolationException
            EnterExportMethod(hGlobalIntPtr);
#endif

            try{
                //A bizarre test case: The other end sends its own objIntPtr, but after free, the memory structure becomes disordered and the test case freezes
                //Unable to determine the validity of the object, once operated, it will result in
                // Access Violation: Attempted to read or write protected memory.
                // Process finished with exit code -1,073,740,940.
                //HGlobalIntPtr has no way to detect whether an application has been made, only test the code and do not write such test cases
                //And code review to prevent such incidents from happening
                //Internal judgment will be made on if (Marshal. IsFullOrWin32Atom (hglobal))
                Marshal.FreeHGlobal(hGlobalIntPtr);
            }
            catch (Exception e){
                // try best 
                _log.Error(e);
            }
        }
    }
}