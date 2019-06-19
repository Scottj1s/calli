using System;
using System.Diagnostics;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;

namespace ConsoleApp
{
    unsafe delegate int delegateQueryInterface([In] IntPtr pThis, [In] Guid iid, IntPtr* pObject);
    delegate uint delegateAddRefRelease([In] IntPtr pThis);
    delegate void delegateVoid([In] IntPtr pThis);

    delegate uint delegateAddRefReleaseCallI(IntPtr pThis, IntPtr pFunction);

    
    class Program
    {
        [DllImport("dummy32.dll", EntryPoint = "CreateDummyUnknown", PreserveSig = true)]
        static extern unsafe IntPtr CreateDummyUnknown32();

        [DllImport("dummy64.dll", EntryPoint = "CreateDummyUnknown", PreserveSig = true)]
        static extern unsafe IntPtr CreateDummyUnknown64();

        //private static byte[] ilbytes;

        //public static byte[] GetILBytes(DynamicMethod dynamicMethod)
        //{
        //    var resolver = typeof(DynamicMethod).GetField("m_resolver", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(dynamicMethod);
        //    if (resolver == null) throw new ArgumentException("The dynamic method's IL has not been finalized.");
        //    return (byte[])resolver.GetType().GetField("m_code", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(resolver);
        //}

        public static Type MakeDefinedCalli()
        {
            var domain = AppDomain.CurrentDomain;
            var asmName = new AssemblyName("MyDynamicAsm");
            var asmBuilder = domain.DefineDynamicAssembly(asmName, AssemblyBuilderAccess.RunAndSave);
            var module = asmBuilder.DefineDynamicModule("MyDynamicAsm", "MyDynamicAsm.dll");
            var typeBuilder = module.DefineType("MyDynamicType", 
                TypeAttributes.Public | TypeAttributes.Sealed | TypeAttributes.Abstract);

            var t_uint = typeof(uint);
            var t_intptr = typeof(IntPtr);
            var method = typeBuilder.DefineMethod("DefinedCalli",
                MethodAttributes.Public | MethodAttributes.Static, CallingConventions.Standard,
                t_uint, new Type[] { t_intptr, t_intptr });
            //method.CreateMethodBody(ilbytes, ilbytes.Length);
            var il = method.GetILGenerator();
            il.Emit(OpCodes.Ldarg_0);
            il.Emit(OpCodes.Ldarg_1);
            il.EmitCalli(OpCodes.Calli, CallingConvention.StdCall, t_uint, new Type[] { t_intptr });
            il.Emit(OpCodes.Ret);

            var type = typeBuilder.CreateType();
            //asmBuilder.Save("MyDynamicAsm.dll");
            return type;
        }

        unsafe static private delegateAddRefReleaseCallI MakeDynamicCalli()
        {
            var t_uint = typeof(uint);
            var t_intptr = typeof(IntPtr);
            var method = new DynamicMethod("DynamicCalli", 
                MethodAttributes.Static | MethodAttributes.Public, CallingConventions.Standard, 
                t_uint, new Type[] { t_intptr, t_intptr }, t_uint, true);
            var il = method.GetILGenerator();
            il.Emit(OpCodes.Ldarg_0);
            il.Emit(OpCodes.Ldarg_1);
//            il.Emit(OpCodes.Conv_I);
            il.EmitCalli(OpCodes.Calli, CallingConvention.StdCall, t_uint, new Type[] { t_intptr });
            il.Emit(OpCodes.Ret);

            var del = (delegateAddRefReleaseCallI)method.CreateDelegate(typeof(delegateAddRefReleaseCallI));
            //ilbytes = GetILBytes(method);

            return del;
        }

        unsafe static void Main(string[] args)
        {
            //IntPtr object_ptr = GetActivationFactory("Windows.Foundation.Uri");
            IntPtr object_ptr = (IntPtr.Size == 4) ? CreateDummyUnknown32() : CreateDummyUnknown64();

            void** object_vfptr = *(void***)object_ptr.ToPointer();
            void* AddRefPtr = object_vfptr[1];
            void* ReleasePtr = object_vfptr[2];
            void* VoidPtr = object_vfptr[3];

            // Generate AddRef and Release calli methods via DynamicMethod
            var AddRefCalli = MakeDynamicCalli();
            var ReleaseCalli = MakeDynamicCalli();

            // Call LCG AddRef calli methods - this works
            AddRefCalli(object_ptr, (IntPtr)AddRefPtr);

            // "Static" IL call causes NullReferenceException on return 
            call.IUnknown.AddRefRelease(object_ptr, (IntPtr)AddRefPtr);
            //call.IUnknown.Void(object_ptr, (IntPtr)VoidPtr);

            // DefineMethod also fails with NullReferenceException
            var calli_type = MakeDefinedCalli();
            calli_type.InvokeMember("DefinedCalli",
                BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static,
                null, null, new object[] { object_ptr, (IntPtr)AddRefPtr });

            // Delegate around defined method also causes NullReferenceException
            var calli_method = calli_type.GetMethod("DefinedCalli");
            var bytes = calli_method.GetMethodBody();
            var calli_delegate = (delegateAddRefReleaseCallI)calli_method.CreateDelegate(typeof(delegateAddRefReleaseCallI));
            var result = calli_delegate(object_ptr, (IntPtr)AddRefPtr);

            Console.ReadKey();
        }
    }
}
