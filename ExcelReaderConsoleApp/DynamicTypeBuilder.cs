using System.ComponentModel.DataAnnotations;
using System.Reflection;
using System.Reflection.Emit;

namespace ExcelReaderConsoleApp;

public class DynamicTypeBuilder
{
    public Type CreateDynamicType(string typeName, List<string> propertyNames, List<Type> propertyTypes)
    {
        var assemblyName = new AssemblyName("DynamicTypes");
        var assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
        var moduleBuilder = assemblyBuilder.DefineDynamicModule("MainModule");
        var typeBuilder = moduleBuilder.DefineType(typeName, TypeAttributes.Public);

        for (int i = 0; i < propertyNames.Count; i++)
        {
            var fieldBuilder = typeBuilder.DefineField("_" + propertyNames[i], propertyTypes[i], FieldAttributes.Private);
            var propertyBuilder = typeBuilder.DefineProperty(propertyNames[i], PropertyAttributes.HasDefault, propertyTypes[i], null);

            var getterMethod = typeBuilder.DefineMethod("get_" + propertyNames[i], MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig, propertyTypes[i], Type.EmptyTypes);
            var getterIL = getterMethod.GetILGenerator();
            getterIL.Emit(OpCodes.Ldarg_0);
            getterIL.Emit(OpCodes.Ldfld, fieldBuilder);
            getterIL.Emit(OpCodes.Ret);

            var setterMethod = typeBuilder.DefineMethod("set_" + propertyNames[i], MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig, null, [propertyTypes[i]]);
            var setterIL = setterMethod.GetILGenerator();
            setterIL.Emit(OpCodes.Ldarg_0);
            setterIL.Emit(OpCodes.Ldarg_1);
            setterIL.Emit(OpCodes.Stfld, fieldBuilder);
            setterIL.Emit(OpCodes.Ret);

            propertyBuilder.SetGetMethod(getterMethod);
            propertyBuilder.SetSetMethod(setterMethod);

            // Add [Key] attribute to the PrimaryKey property
            if (propertyNames[i] == "PrimaryKey")
            {
                var keyAttributeConstructor = typeof(KeyAttribute).GetConstructor(Type.EmptyTypes);
                var keyAttributeBuilder = new CustomAttributeBuilder(keyAttributeConstructor!, []);
                propertyBuilder.SetCustomAttribute(keyAttributeBuilder);
            }
        }

        return typeBuilder.CreateType();
    }

    public Type CreateInheritedType(string typeName, Type baseType, Type[] constructorArgs)
    {
        var assemblyName = new AssemblyName("DynamicTypes");
        var assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
        var moduleBuilder = assemblyBuilder.DefineDynamicModule("MainModule");
        var typeBuilder = moduleBuilder.DefineType(typeName, TypeAttributes.Public, baseType);

        // Define a constructor with parameters
        var constructorBuilder = typeBuilder.DefineConstructor(MethodAttributes.Public, CallingConventions.Standard, constructorArgs);

        var ilGenerator = constructorBuilder.GetILGenerator();
        ilGenerator.Emit(OpCodes.Ldarg_0); // Load "this"
        for (int i = 0; i < constructorArgs.Length; i++)
        {
            ilGenerator.Emit(OpCodes.Ldarg, i + 1); // Load each argument
        }

        ilGenerator.Emit(OpCodes.Call, baseType.GetConstructor(constructorArgs)); // Call base constructor
        ilGenerator.Emit(OpCodes.Ret); // Return

        return typeBuilder.CreateType();
    }


}