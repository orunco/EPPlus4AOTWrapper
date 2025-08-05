using System.Reflection;
using System.Runtime.InteropServices;
using OfficeOpenXml.Corestone;
using Logger = OfficeOpenXml.Corestone.Logger;

namespace OfficeOpenXml;

/*
20240730 经过2天的研究，对SafeHandle有新的理解：
1、SafeHandle和CriticalHandle的选择，正如有个文档写的：SafeHandle多了引用计数特性
   因为我们会大量取SafeHandle内部的IntPtr进行操作，最好引用计数
   另外，多线程在获取内部handle时，很容易出现bug，只有引用计数可以拯救

2、SafeHandleZeroOrMinusOneIsInvalid一开始就觉得名字怪怪的，而且是WIN32的，打开看了，
   原来就多了定义
   IsInvalid => this.handle == IntPtr.Zero || this.handle == new IntPtr(-1);
   这个-1是WIN32函数特有的：如果该函数发生错误,其返回值为-1
   所以继承自这个类是错误的
3、一开始每一个类都配置了一个特定的Handle，因为需要对端dll特定的FreeHandle操作函数
   仔细想想，对端dll只需要配置一个特定的dll FreeHandle导出函数就可以了，因为dll内置的
   gc来者不拒：GCHandle.FromIntPtr(intptr).Free(); 并不区分类型
   当然，这个也是想了半天的结果；最开始还有另外一种奇葩的想法：对端dll都不需要导出FreeHandle
   直接在托管代码中也可以GCHandle.FromIntPtr(intptr).Free()......

   下面代码是托管环境的实测代码
   var intPtr = dll对端过来的指针

   "1 " + ((Entry)GCHandle.FromIntPtr(intPtr).Target).Print();
   直接报System.AccessViolationException: Attempted to read or write protected memory.
   因为托管环境中没法转换这个对象

   GCHandle.FromIntPtr(intPtr).Free();
   但是这一句不会报错，一路查询到gc源代码handletable.cpp的void HndDestroyHandle函数
        // fetch the handle table pointer
        HandleTable *pTable = Table(hTable);

        // return the handle to the table's cache
        TableFreeSingleHandleToCache(pTable, uType, handle);
    托管环境中的GC持有了HandleTable

    "2 " + ((Entry)GCHandle.FromIntPtr(intPtr).Target).Print();
    这句报空指针，和上面的有差异。

    所以，这里有个关键点：dll中的coreCLR和ForNet中的coreCLR是2套环境。
    谁申请、谁释放；谁建设、谁负责，这不是很明显的事么？

4、IntPtr free后本身的值不会变成zero，需要特殊设置

5、关于dll中GCHandle的理解：有一篇文档讲解的相当好，引用类的对象可能会被GC移动，所以框柱这个对象
   是必须的，有3种方法：
        fixed：之前用过了
        Marshal： 直接复制数据了
        GCHandle.alloc(pinned): 直接pin住
   但是看很多的代码，只有数组等少量的代码需要pinned，其他的都不需要，为什么？ 仔细思考，才发现
   这些代码是拿到了数组指针后，直接对地址上的值直接操作的，所以不pin是不行的。而我们是对象操作，
   GCHandle.alloc本身返回的指针由dll保证了指针修复，保证了跨越dll到外层环境的一致性。

 */
public class ExcelSafeHandle : SafeHandle{
    
    private static readonly Logger log = LoggerFactory.GetLogger(MethodBase.GetCurrentMethod());

    public ExcelSafeHandle(IntPtr initIntPtr) : base(initIntPtr, true){
        // 不需要判断initIntPtr是否为zero，没必要
    }

    public override bool IsInvalid => handle == IntPtr.Zero;

    // GC保证这个函数只执行一次
    protected override bool ReleaseHandle(){
        // 谁申请，谁释放
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern bool ExcelSafeHandle_FreeHandleInDll(IntPtr thisHandle);

#if AccessViolationException
        ExceptionForNet.CallLibraryMethod(this);
#endif
        
        if (IsInvalid){
            return IsInvalid;
        }

        try{
            // localHandle返回的是handle的原值，这里是多线程环境下设置handle为zero
            var localHandle = Interlocked.Exchange(ref handle, IntPtr.Zero);

            if (localHandle != IntPtr.Zero){
#if AccessViolationException
        log.Info($"localHandle={localHandle:X}, try free.");
#endif
                try{
                    // 虽然对端承诺 不会返回异常 但是测试发现，如果这个localHandle是本端产生的
                    // 发到对端，虽然这样不正确，测试用例会卡主
                    // 尽最大努力不发生异常
                    ExcelSafeHandle_FreeHandleInDll(localHandle);
                }
                catch (Exception e){
                    // ignored
                }
            }
        }
        catch (Exception e){
#if DEBUG
            log.Info(e);
#endif
        }
        finally{
            SetHandleAsInvalid();
        }

        return true;
    }

    // 注意：hGlobalIntPtr必须是对端有效指针，否则释放会引发进程异常
    public static void ReleaseHGlobal(IntPtr hGlobalIntPtr){
        // 谁申请，谁释放
        [DllImport(LibraryResolver.Name, CharSet = CharSet.Unicode)]
        static extern void ExcelSafeHandle_FreeHGlobalInDll(IntPtr hGlobalIntPtr);
#if AccessViolationException
        ExceptionForNet.CallLibraryMethod(hGlobalIntPtr);
#endif
        // 不会返回异常
        ExcelSafeHandle_FreeHGlobalInDll(hGlobalIntPtr);
    }
}