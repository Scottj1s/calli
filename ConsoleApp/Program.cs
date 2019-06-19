#define INTEROP

using System;
using System.Diagnostics;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;

#if INTEROP

using Windows;
using Windows.Foundation;

namespace Windows.Foundation
{

    public sealed class HString : IDisposable
    {
        public static class Interop
        {
            [DllImport("api-ms-win-core-winrt-string-l1-1-0.dll", CallingConvention = CallingConvention.StdCall)]
            static extern unsafe int WindowsCreateString([MarshalAs(UnmanagedType.LPWStr)] string source, int length, [Out] IntPtr* hstring);

            [DllImport("api-ms-win-core-winrt-string-l1-1-0.dll", CallingConvention = CallingConvention.StdCall)]
            static extern int WindowsDeleteString(IntPtr hstring);

            [DllImport("api-ms-win-core-winrt-string-l1-1-0.dll", CallingConvention = CallingConvention.StdCall)]
            static extern unsafe char* WindowsGetStringRawBuffer(IntPtr hstring, [Out] uint* length);

            public static IntPtr Create(string source)
            {
                if (string.IsNullOrEmpty(source))
                {
                    return IntPtr.Zero;
                }

                IntPtr hstring;
                unsafe
                {
                    Marshal.ThrowExceptionForHR(WindowsCreateString(source, source.Length, &hstring));
                }
                return hstring;
            }

            public static void Delete(IntPtr hstring)
            {
                Marshal.ThrowExceptionForHR(WindowsDeleteString(hstring));
            }

            public static string ToString(IntPtr hstring)
            {
                if (hstring == IntPtr.Zero)
                {
                    return string.Empty;
                }

                unsafe
                {
                    uint length;
                    var buffer = WindowsGetStringRawBuffer(hstring, &length);
                    return new string(buffer, 0, checked((int)length));
                }
            }
        }

        private IntPtr _handle;
        private bool _disposed = false;

        public HString(IntPtr handle)
        {
            _handle = handle;
        }

        public HString() : this(null)
        {
        }

        public HString(string value)
        {
            _handle = Interop.Create(value);
        }

        public IntPtr Handle => _disposed ? throw new ObjectDisposedException("Windows.HString") : _handle;

        public override string ToString()
        {
            if (_disposed) throw new ObjectDisposedException("Windows.HString");
            return Interop.ToString(_handle);
        }

        void _Dispose()
        {
            if (!_disposed)
            {
                Interop.Delete(_handle);
                _handle = IntPtr.Zero;
                _disposed = true;
            }
        }

        ~HString()
        {
            _Dispose();
        }

        public void Dispose()
        {
            _Dispose();
            GC.SuppressFinalize(this);
        }
    }

    public sealed class ComPtr : IDisposable
    {
        private IntPtr _value;
        private bool _disposed = false;

        public ComPtr(IntPtr value)
        {
            _value = value;
        }

        public IntPtr Value => _disposed ? throw new ObjectDisposedException("Windows.ComPtr") : _value;

        public ComPtr As(Guid iid)
        {
            if (_disposed) throw new ObjectDisposedException("Windows.ComPtr");
            return Windows.Foundation.IUnknown.As(_value, iid);
        }

        void _Dispose()
        {
            if (!_disposed)
            {
                Windows.Foundation.IUnknown.Release(_value);
                _value = IntPtr.Zero;

                _disposed = true;
            }
        }

        ~ComPtr()
        {
            _Dispose();
        }

        public void Dispose()
        {
            _Dispose();
            System.GC.SuppressFinalize(this);
        }
    }

    //#pragma warning disable 0649
    public static class IUnknown
    {
        public static readonly Guid IID = new Guid("00000000-0000-0000-C000-000000000046");

        unsafe delegate int delegateQueryInterface(IntPtr @this, [In] ref Guid iid, IntPtr* @object);

        unsafe public static int invokeQueryInterface(IntPtr @this, ref Guid iid, IntPtr* @object)
        {
            void* __slot = (*(void***)@this.ToPointer())[0];
            var __delegate = Marshal.GetDelegateForFunctionPointer<delegateQueryInterface>(new IntPtr(__slot));
            return __delegate(@this, ref iid, @object);
        }

        public static IntPtr QueryInterface(IntPtr @this, Guid iid)
        {
            IntPtr instance = IntPtr.Zero;
            unsafe
            {
                Marshal.GetExceptionForHR(invokeQueryInterface(@this, ref iid, &instance));
            }
            return instance;
        }

        public static ComPtr As(IntPtr @this, Guid iid)
        {
            return new ComPtr(QueryInterface(@this, iid));
        }

        delegate uint delegateAddRef(IntPtr @this);

        unsafe public static uint invokeAddRef(IntPtr @this)
        {
            void* __slot = (*(void***)@this.ToPointer())[1];
            var __delegate = Marshal.GetDelegateForFunctionPointer<delegateAddRef>(new IntPtr(__slot));
            return __delegate(@this);
        }

        public static uint AddRef(IntPtr @this)
        {
            return invokeAddRef(@this);
        }

        delegate uint delegateRelease(IntPtr @this);

        unsafe public static uint invokeRelease(IntPtr @this)
        {
            void* __slot = (*(void***)@this.ToPointer())[2];
            var __delegate = Marshal.GetDelegateForFunctionPointer<delegateRelease>(new IntPtr(__slot));
            return __delegate(@this);
        }

        public static uint Release(IntPtr @this)
        {
            return invokeRelease(@this);
        }
    }
}

#endif

namespace ConsoleApp
{
    unsafe delegate int delegateQueryInterface([In] IntPtr pThis, [In] Guid iid, IntPtr* pObject);
    delegate uint delegateAddRefRelease([In] IntPtr pThis);
    delegate void delegateVoid([In] IntPtr pThis);

    delegate uint delegateAddRefReleaseCallI(IntPtr pThis, IntPtr pFunction);

#if INTEROP

    struct IUnknownVftbl
    {
#pragma warning disable 0649  
        public delegateQueryInterface QueryInterface;
        public delegateAddRefRelease AddRef;
        public delegateAddRefRelease Release;
        public delegateVoid Void;
#pragma warning restore 0649
    }

    struct IUnkownObject
    {
#pragma warning disable 0649
        public IntPtr Vftbl;
#pragma warning restore 0649
    }
#endif


    class Program
    {
        [DllImport("dummy32.dll", EntryPoint = "CreateDummyUnknown", PreserveSig = true)]
        static extern unsafe IntPtr CreateDummyUnknown32();

        [DllImport("dummy64.dll", EntryPoint = "CreateDummyUnknown", PreserveSig = true)]
        static extern unsafe IntPtr CreateDummyUnknown64();

#if INTEROP

        [DllImport("api-ms-win-core-winrt-l1-1-0.dll", PreserveSig = true)]
        static extern unsafe int RoGetActivationFactory([In] IntPtr activatableClassId, [In] ref Guid iid, [Out] IntPtr* factory);

        public static readonly Guid IActivationFactory_IID = new Guid("00000035-0000-0000-C000-000000000046");

        public static IntPtr GetActivationFactory(string className)
        {
            using (var classNameHstring = new HString(className))
            {
                IntPtr factory = IntPtr.Zero;
                unsafe
                {
                    Guid iid = IActivationFactory_IID;
                    Marshal.ThrowExceptionForHR(RoGetActivationFactory(classNameHstring.Handle, ref iid, &factory));
                }
                return factory;
            }
        }
#endif

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

        unsafe static private DynamicMethod MakeInvokeCalli(IntPtr function_ptr)
        {
            var t_uint = typeof(uint);
            var t_intptr = typeof(IntPtr);
            var method = new DynamicMethod("DynamicCalli",
                MethodAttributes.Static | MethodAttributes.Public, CallingConventions.Standard,
                t_uint, new Type[] { t_intptr }, t_uint, true);
            var il = method.GetILGenerator();
            il.Emit(OpCodes.Ldarg_0);
            il.Emit(OpCodes.Ldc_I8, (long)function_ptr);
            //il.Emit(OpCodes.Conv_I);
            il.EmitCalli(OpCodes.Calli, CallingConvention.StdCall, t_uint, new Type[] { t_intptr });
            il.Emit(OpCodes.Ret);
            return method;
        }

        const long COUNT = 20;
        const long MILLION = 1000000;

        unsafe static void Main(string[] args)
        {
            //IntPtr object_ptr = GetActivationFactory("Windows.Foundation.Uri");
            IntPtr object_ptr = (IntPtr.Size == 4) ? CreateDummyUnknown32() : CreateDummyUnknown64();

            void** object_vfptr = *(void***)object_ptr.ToPointer();
            void* AddRefPtr = object_vfptr[1];
            void* ReleasePtr = object_vfptr[2];

            // Generate AddRef and Release calli methods via DynamicMethod
            var AddRefCalli = MakeDynamicCalli();
            var ReleaseCalli = MakeDynamicCalli();

            if(false)
            {
            // Call LCG AddRef calli methods - this works
            AddRefCalli(object_ptr, (IntPtr)AddRefPtr);

            // "Static" IL call causes NullReferenceException on return 
            call.IUnknown.AddRefRelease(object_ptr, (IntPtr)AddRefPtr);
            //void* VoidPtr = object_vfptr[3];
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
            }

#if INTEROP

            var sw = Stopwatch.StartNew();
            // Test IL calli
            {
                Trace.Assert(call.IUnknown.AddRefRelease(object_ptr, (IntPtr)AddRefPtr) == (uint)2);
                Trace.Assert(call.IUnknown.AddRefRelease(object_ptr, (IntPtr)ReleasePtr) == (uint)1);

                sw.Restart();
                for (long i = 0; i < (COUNT * MILLION); i++)
                {
                    call.IUnknown.AddRefRelease(object_ptr, (IntPtr)AddRefPtr);
                    call.IUnknown.AddRefRelease(object_ptr, (IntPtr)ReleasePtr);
                }
                Console.WriteLine("IL calli: {0:D}nsec", sw.ElapsedMilliseconds / COUNT);
            }

            // Test emitted calli
            {
                //AddRefCalli += (IntPtr) => { return 0; };

                sw.Restart();
                for (long i = 0; i < (COUNT * MILLION); i++)
                {
                    AddRefCalli(object_ptr, (IntPtr)AddRefPtr);
                    ReleaseCalli(object_ptr, (IntPtr)ReleasePtr);
                }
                Console.WriteLine("Emitted calli Delegate: {0:D}nsec", sw.ElapsedMilliseconds / COUNT);

                //// For some reason, this is super slow...
                //sw.Restart();
                //var AddRefInvoke = MakeInvokeCalli((IntPtr)AddRefPtr);
                //var ReleaseInvoke = MakeInvokeCalli((IntPtr)ReleasePtr);
                //var invoke_args = new object[] { object_ptr };
                //for (long i = 0; i < (COUNT * MILLION); i++)
                //{
                //    AddRefInvoke.Invoke(null, invoke_args);
                //    ReleaseInvoke.Invoke(null, invoke_args);
                //}
                //Console.WriteLine("Emitted calli Invoke: {0:D}nsec", sw.ElapsedMilliseconds / COUNT);
            }

            // Test Marshal.GetDelegateForFunctionPointer
            {
                var AddRef = Marshal.GetDelegateForFunctionPointer<delegateAddRefRelease>(new IntPtr(AddRefPtr));
                var Release = Marshal.GetDelegateForFunctionPointer<delegateAddRefRelease>(new IntPtr(ReleasePtr));
                //AddRef += (IntPtr) => { return 0; };

                sw.Restart();
                for (long i = 0; i < (COUNT * MILLION); i++)
                {
                    AddRef(object_ptr);
                    Release(object_ptr);
                }
                Console.WriteLine("Marshal.GetDelegateForFunctionPointer: {0:D}nsec", sw.ElapsedMilliseconds / COUNT);
            }

            // Test Marshal.PtrToStructure
            {
                var factory_obj = Marshal.PtrToStructure<IUnkownObject>(object_ptr);
                var factory_unk = Marshal.PtrToStructure<IUnknownVftbl>(factory_obj.Vftbl);

                sw.Restart();
                for (long i = 0; i < (COUNT * MILLION); i++)
                {
                    factory_unk.AddRef(object_ptr);
                    factory_unk.Release(object_ptr);
                }
                Console.WriteLine("Marshal.PtrToStructure: {0:D}nsec", sw.ElapsedMilliseconds / COUNT);
            }
            IUnknown.Release(object_ptr);
#endif

//            Console.WriteLine("<press any key>");
            Console.ReadKey();
        }
    }
}
