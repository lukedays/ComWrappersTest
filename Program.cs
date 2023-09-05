using System.Collections;
using System.Diagnostics;
using System.Runtime.InteropServices;
using static System.Runtime.InteropServices.ComWrappers;

namespace ComWrappersTest
{
    internal class Program
    {
        public static void Main()
        {
            // The CLSID for WScript.Shell
            var clsid = new Guid("{F935DC22-1CF0-11D0-ADB9-00C04FD58A0B}");

            // The IID for IWshShell
            var iid = IWshShell.IID_IRunType;

            var cw = new MyComWrappers();

            var shell = ActivateClass<IWshShell>(cw, clsid, iid);

            shell.Run("notepad.exe");
        }

        [DllImport("Ole32")]
        static extern int CoCreateInstance(
            ref Guid rclsid,
            IntPtr pUnkOuter,
            uint dwClsContext,
            ref Guid riid,
            out IntPtr ppv
        );

        public static I ActivateClass<I>(ComWrappers cw, Guid clsid, Guid iid)
        {
            int hr = CoCreateInstance(
                ref clsid,
                IntPtr.Zero, /*CLSCTX_INPROC_SERVER*/
                1,
                ref iid,
                out IntPtr obj
            );
            if (hr < 0)
            {
                Marshal.ThrowExceptionForHR(hr);
            }
            return (I)cw.GetOrCreateObjectForComInstance(obj, CreateObjectFlags.UniqueInstance);
        }
    }

    interface IWshShell
    {
        public static Guid IID_IRunType = new("F935DC21-1CF0-11D0-ADB9-00C04FD58A0B");

        int Run(string Command);
    }

    sealed unsafe class MyComWrappers : ComWrappers
    {
        //static readonly IntPtr s_IRunVTable;
        //static readonly ComInterfaceEntry* s_DemoImplDefinition;
        //static readonly int s_DemoImplDefinitionLen;

        //static MyComWrappers()
        //{
        //    //// Get system provided IUnknown implementation.
        //    GetIUnknownImpl(out IntPtr fpQueryInterface, out IntPtr fpAddRef, out IntPtr fpRelease);

        //    //////
        //    ////// Construct VTables for supported interfaces
        //    //////
        //    {
        //        int tableCount = 10;
        //        int idx = 0;
        //        var vtable = (IntPtr*)
        //            RuntimeHelpers.AllocateTypeAssociatedMemory(
        //                typeof(MyComWrappers),
        //                IntPtr.Size * tableCount
        //            );
        //        vtable[idx++] = fpQueryInterface;
        //        vtable[idx++] = fpAddRef;
        //        vtable[idx++] = fpRelease;
        //        vtable[idx++] = IntPtr.Zero;
        //        vtable[idx++] = IntPtr.Zero;
        //        vtable[idx++] = IntPtr.Zero;
        //        vtable[idx++] = IntPtr.Zero;
        //        vtable[idx++] = IntPtr.Zero;
        //        vtable[idx++] = IntPtr.Zero;
        //        vtable[idx++] = (IntPtr)
        //            (delegate* unmanaged<IntPtr, IntPtr, int>)&ABI.IRunTypeManagedWrapper.Run;
        //        Debug.Assert(tableCount == idx);
        //        s_IRunVTable = (IntPtr)vtable;
        //    }

        //    //////
        //    ////// Construct entries for supported managed types
        //    //////
        //    {
        //        s_DemoImplDefinitionLen = 1;
        //        int idx = 0;
        //        var entries = (ComInterfaceEntry*)
        //            RuntimeHelpers.AllocateTypeAssociatedMemory(
        //                typeof(MyComWrappers),
        //                sizeof(ComInterfaceEntry) * s_DemoImplDefinitionLen
        //            );
        //        entries[idx].IID = IWshShell.IID_IRunType;
        //        entries[idx++].Vtable = s_IRunVTable;
        //        Debug.Assert(s_DemoImplDefinitionLen == idx);
        //        s_DemoImplDefinition = entries;
        //    }
        //}

        readonly delegate*<IntPtr, object?> _createIfSupported;

        public MyComWrappers(bool useDynamicNativeWrapper = false)
        {
            // Determine which wrapper create function to use.
            _createIfSupported = useDynamicNativeWrapper
                ? &ABI.NativeDynamicWrapper.CreateIfSupported
                : &ABI.NativeStaticWrapper.CreateIfSupported;
        }

        protected override unsafe ComInterfaceEntry* ComputeVtables(
            object obj,
            CreateComInterfaceFlags flags,
            out int count
        )
        {
            throw new NotImplementedException();
        }

        //protected override unsafe ComInterfaceEntry* ComputeVtables(
        //    object obj,
        //    CreateComInterfaceFlags flags,
        //    out int count
        //)
        //{
        //    Debug.Assert(flags is CreateComInterfaceFlags.None);

        //    if (obj is WshShellImpl)
        //    {
        //        count = s_DemoImplDefinitionLen;
        //        return s_DemoImplDefinition;
        //    }

        //    // Note: this implementation does not handle cases where the passed in object implements
        //    // one or both of the supported interfaces but is not the expected .NET class.
        //    count = 0;
        //    return null;
        //}

        protected override object CreateObject(IntPtr externalComObject, CreateObjectFlags flags)
        {
            Debug.Assert(flags.HasFlag(CreateObjectFlags.UniqueInstance));

            return _createIfSupported(externalComObject) ?? throw new NotSupportedException();
        }

        protected override void ReleaseObjects(IEnumerable objects)
        {
            throw new NotImplementedException();
        }
    }

    namespace ABI
    {
        internal enum HRESULT : int
        {
            S_OK = 0
        }

        internal static unsafe class IRunTypeManagedWrapper
        {
            [UnmanagedCallersOnly]
            public static int Run(IntPtr _this, IntPtr Command)
            {
                // Marshal to .NET
                string? strLocal = Command == IntPtr.Zero ? null : Marshal.PtrToStringUni(Command);

                try
                {
                    int s = ComInterfaceDispatch
                        .GetInstance<IWshShell>((ComInterfaceDispatch*)_this)
                        .Run(strLocal);
                }
                catch (Exception e)
                {
                    return e.HResult;
                }

                return (int)HRESULT.S_OK;
            }
        }

        internal class NativeStaticWrapper : IWshShell, IDisposable
        {
            bool _isDisposed = false;

            private NativeStaticWrapper() { }

            ~NativeStaticWrapper()
            {
                DisposeInternal();
            }

            public IntPtr IRunTypeInst { get; init; }

            public static NativeStaticWrapper? CreateIfSupported(IntPtr ptr)
            {
                int hr = Marshal.QueryInterface(
                    ptr,
                    ref IWshShell.IID_IRunType,
                    out IntPtr IRunTypeInst
                );
                if (hr != (int)HRESULT.S_OK)
                {
                    return default;
                }

                return new NativeStaticWrapper() { IRunTypeInst = IRunTypeInst };
            }

            public void Dispose()
            {
                DisposeInternal();
                GC.SuppressFinalize(this);
            }

            public int Run(string Command) => IRunTypeNativeWrapper.Run(IRunTypeInst, Command);

            void DisposeInternal()
            {
                if (_isDisposed)
                    return;

                Marshal.Release(IRunTypeInst);

                _isDisposed = true;
            }
        }

        internal class NativeDynamicWrapper : IDynamicInterfaceCastable, IDisposable
        {
            bool _isDisposed = false;

            private NativeDynamicWrapper() { }

            ~NativeDynamicWrapper()
            {
                DisposeInternal();
            }

            public IntPtr IRunTypeInst { get; init; }

            public static NativeDynamicWrapper? CreateIfSupported(IntPtr ptr)
            {
                int hr = Marshal.QueryInterface(
                    ptr,
                    ref IWshShell.IID_IRunType,
                    out IntPtr IDemoGetTypeInst
                );
                if (hr != (int)HRESULT.S_OK)
                {
                    return default;
                }

                return new NativeDynamicWrapper() { IRunTypeInst = IDemoGetTypeInst };
            }

            public RuntimeTypeHandle GetInterfaceImplementation(RuntimeTypeHandle interfaceType)
            {
                if (interfaceType.Equals(typeof(IWshShell).TypeHandle))
                {
                    return typeof(IRunTypeNativeWrapper).TypeHandle;
                }

                return default;
            }

            public bool IsInterfaceImplemented(
                RuntimeTypeHandle interfaceType,
                bool throwIfNotImplemented
            )
            {
                if (interfaceType.Equals(typeof(IWshShell).TypeHandle))
                {
                    return true;
                }

                if (throwIfNotImplemented)
                    throw new InvalidCastException(
                        $"{nameof(NativeDynamicWrapper)} doesn't support {interfaceType}"
                    );

                return false;
            }

            public void Dispose()
            {
                DisposeInternal();
                GC.SuppressFinalize(this);
            }

            void DisposeInternal()
            {
                if (_isDisposed)
                    return;

                Marshal.Release(IRunTypeInst);

                _isDisposed = true;
            }
        }

        [DynamicInterfaceCastableImplementation]
        internal unsafe interface IRunTypeNativeWrapper : IWshShell
        {
            public static int Run(IntPtr inst, string Command)
            {
                //var func = (delegate* unmanaged<IntPtr, IntPtr, int>)(
                //    *(
                //        *(void***)inst + 9 /* IWshShell.Run slot */
                //    )
                //);

                var func = (delegate* unmanaged<IntPtr, IntPtr, int, bool, int>)(
                    *(
                        *(void***)inst + 9 /* IWshShell.Run slot */
                    )
                );

                // ACCESS VIOLATION ERROR HERE
                //var ret = func(inst, Marshal.StringToBSTR(Command));
                var ret = func(inst, Marshal.StringToBSTR(Command), 1, true);

                return 1;
            }

            int IWshShell.Run(string Command)
            {
                var inst = ((NativeDynamicWrapper)this).IRunTypeInst;
                return Run(inst, Command);
            }
        }
    }
}
