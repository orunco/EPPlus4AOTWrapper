namespace OfficeOpenXml;

public class ExceptionCodeTool{
    public static void Main(string[] args){
        GetSubClassNames(typeof(Exception));
    }

    // Retrieve all exception types in the assembly at once
    private static void GetSubClassNames(Type parentType){
        foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies()){
            Console.WriteLine($"// {assembly}");
            foreach (var itemType in assembly.GetTypes()) 
            {
                if (parentType.IsAssignableFrom(itemType))
                {
                    var template =
                        """
                        if (type.ToString() == typeof(__PLACE__).ToString()){
                            var e = JsonSerializer.Deserialize(
                                serializedData, JsonContext.Default.__PLACE__);
                            if (e == null)
                                throw new Exception(typeName + Environment.NewLine + serializedData);
                            SetStackTrace(e,layout.StackTrace);
                            throw e;
                        }
                        """;
                    var str = itemType.ToString();
                    int pos = str.LastIndexOf('.');
                    Console.WriteLine(template.Replace(
                        "__PLACE__",
                        str.Substring(pos + 1)));
                    // if (type.ToString() == typeof(ArgumentException).ToString()){
                    //     var e = JsonSerializer.Deserialize(
                    //         serializedData, JsonContext.Default.ArgumentException);
                    //     if (e == null)
                    //         throw new Exception(typeName + Environment.NewLine + serializedData);
                    //     SetStackTrace(e,layout.StackTrace);
                    //     throw e;
                    // }

                    //[JsonSerializable(typeof(Exception), GenerationMode = JsonSourceGenerationMode.Metadata)]
                    //log.Info($"[JsonSerializable(typeof({itemType}), GenerationMode = JsonSourceGenerationMode.Metadata)]");
                }
            }
        }
    }
}