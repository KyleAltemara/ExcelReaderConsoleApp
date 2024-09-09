using System.ComponentModel.DataAnnotations;
using System.Reflection;
using System.Reflection.Emit;

namespace ExcelReaderConsoleApp;

/// <summary>
/// A class that provides methods for dynamically creating types.
/// </summary>
public class DynamicTypeBuilder
{
    /// <summary>
    /// Creates a dynamic type with the specified name, property names, and property types.
    /// </summary>
    /// <param name="typeName">The name of the dynamic type.</param>
    /// <param name="propertyNames">The names of the properties.</param>
    /// <param name="propertyTypes">The types of the properties.</param>
    /// <returns>The created dynamic type.</returns>
    public Type CreateDynamicType(string typeName, List<string> propertyNames, List<Type> propertyTypes)
    {
        // Define the assembly, module, and type
        var assemblyName = new AssemblyName("DynamicTypes");
        var assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
        var moduleBuilder = assemblyBuilder.DefineDynamicModule("MainModule");
        var typeBuilder = moduleBuilder.DefineType(typeName, TypeAttributes.Public);

        for (int i = 0; i < propertyNames.Count; i++)
        {
            // Define a private field for each property
            var fieldBuilder = typeBuilder.DefineField("_" + propertyNames[i], propertyTypes[i], FieldAttributes.Private);

            // Define a property for each field
            var propertyBuilder = typeBuilder.DefineProperty(propertyNames[i], PropertyAttributes.HasDefault, propertyTypes[i], null);

            // Define a getter method for the property
            var getterMethod = typeBuilder.DefineMethod("get_" + propertyNames[i], MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig, propertyTypes[i], Type.EmptyTypes);
            var getterIL = getterMethod.GetILGenerator();
            getterIL.Emit(OpCodes.Ldarg_0); // Load "this"
            getterIL.Emit(OpCodes.Ldfld, fieldBuilder); // Load the field value
            getterIL.Emit(OpCodes.Ret); // Return the field value

            // Define a setter method for the property
            var setterMethod = typeBuilder.DefineMethod("set_" + propertyNames[i], MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig, null, [propertyTypes[i]]);
            var setterIL = setterMethod.GetILGenerator();
            setterIL.Emit(OpCodes.Ldarg_0); // Load "this"
            setterIL.Emit(OpCodes.Ldarg_1); // Load the value to set
            setterIL.Emit(OpCodes.Stfld, fieldBuilder); // Set the field value
            setterIL.Emit(OpCodes.Ret); // Return

            // Set the getter and setter methods for the property
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

        // Create the dynamic type
        return typeBuilder.CreateType();
    }

    /// <summary>
    /// Creates a dynamic type that inherits from the specified base type and has a constructor with the specified arguments.
    /// </summary>
    /// <param name="typeName">The name of the dynamic type.</param>
    /// <param name="baseType">The base type to inherit from.</param>
    /// <param name="constructorArgs">The types of the constructor arguments.</param>
    /// <returns>The created dynamic type.</returns>
    public Type CreateInheritedType(string typeName, Type baseType, Type[] constructorArgs)
    {
        var assemblyName = new AssemblyName("DynamicTypes");
        var assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
        var moduleBuilder = assemblyBuilder.DefineDynamicModule("MainModule");
        var typeBuilder = moduleBuilder.DefineType(typeName, TypeAttributes.Public, baseType);

        // Define a constructor with parameters
        var constructorBuilder = typeBuilder.DefineConstructor(MethodAttributes.Public, CallingConventions.Standard, constructorArgs);

        // Generate IL code for the constructor, IL code is a stack-based language that is used to define the behavior of methods
        // ilGenerator.Emit is used to add instructions to the IL code
        var ilGenerator = constructorBuilder.GetILGenerator();
        ilGenerator.Emit(OpCodes.Ldarg_0); // Load "this"
        for (int i = 0; i < constructorArgs.Length; i++)
        {
            ilGenerator.Emit(OpCodes.Ldarg, i + 1); // Load each argument
        }

        ilGenerator.Emit(OpCodes.Call, baseType.GetConstructor(constructorArgs)!); // Call base constructor
        ilGenerator.Emit(OpCodes.Ret); // Return

        // Create the dynamic type
        return typeBuilder.CreateType();
    }
}
