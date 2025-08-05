using System.Reflection.Emit;

namespace OfficeOpenXml;

public class AotHelper{
    public static bool IsAOT(){
        
        try{
            new DynamicMethod("IsAOT", null, Type.EmptyTypes)
                .GetILGenerator()
                .Emit(OpCodes.Nop);
        }
        catch (Exception){
            return true;
        }

        return false;
    }
}