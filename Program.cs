using System.Collections;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ComWrappersTest
{
    internal class Program
    {
        public static void Main()
        {
            // The CLSID for WScript.Shell
            var clsid = new Guid("{F935DC22-1CF0-11D0-ADB9-00C04FD58A0B}");

            // The IID for IWshShell
            var iid = IWshShell.IID;

            var shell = ActivateClass<IWshShell>(clsid, iid);

            Console.WriteLine(shell.ExpandEnvironmentStrings("%USERPROFILE%"));
        }

        [DllImport("ole32.dll")]
        static extern int CoCreateInstance(
            ref Guid rclsid,
            IntPtr pUnkOuter,
            int dwClsContext,
            ref Guid riid,
            out IntPtr ppv
        );

        public static I ActivateClass<I>(Guid clsid, Guid iid)
        {
            int hr = CoCreateInstance(
                ref clsid,
                IntPtr.Zero,
                1, /*CLSCTX_INPROC_SERVER*/
                ref iid,
                out IntPtr obj
            );
            if (hr < 0)
            {
                Marshal.ThrowExceptionForHR(hr);
            }
            var cw = new MyComWrappers();
            return (I)cw.GetOrCreateObjectForComInstance(obj, CreateObjectFlags.UniqueInstance);
        }
    }

    interface IWshShell
    {
        public static Guid IID = new("41904400-BE18-11D3-A28B-00104BD35090");

        string ExpandEnvironmentStrings(string Src);
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


        protected override unsafe ComInterfaceEntry* ComputeVtables(
            object obj,
            CreateComInterfaceFlags flags,
            out int count
        )
        {
            throw new NotImplementedException();
        }

        protected override object CreateObject(IntPtr externalComObject, CreateObjectFlags flags)
        {
            Debug.Assert(flags.HasFlag(CreateObjectFlags.UniqueInstance));

            return ABI.NativeStaticWrapper.CreateIfSupported(externalComObject)
                ?? throw new NotSupportedException();
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

        internal class NativeStaticWrapper : IWshShell, IDisposable
        {
            bool _isDisposed = false;

            public IntPtr Inst { get; init; }

            public static NativeStaticWrapper? CreateIfSupported(IntPtr ptr)
            {
                int hr = Marshal.QueryInterface(ptr, ref IWshShell.IID, out IntPtr Inst);
                if (hr != (int)HRESULT.S_OK)
                {
                    return default;
                }

                return new NativeStaticWrapper() { Inst = Inst };
            }

            public string ExpandEnvironmentStrings(string Src) =>
                INativeWrapper.ExpandEnvironmentStrings(Inst, Src);

            ~NativeStaticWrapper()
            {
                DisposeInternal();
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

                Marshal.Release(Inst);

                _isDisposed = true;
            }
        }

        [DynamicInterfaceCastableImplementation]
        internal unsafe interface INativeWrapper : IWshShell
        {
            public static string ExpandEnvironmentStrings(IntPtr inst, string Src)
            {
                var func = (delegate* unmanaged<IntPtr, IntPtr, IntPtr*, void>)(
                    *(
                        *(void***)inst + 12 /* IWshShell.ExpandEnvironmentStrings slot */
                    )
                );

                IntPtr retPtr;
                func(inst, Marshal.StringToBSTR(Src), &retPtr);
                var ret = Marshal.PtrToStringBSTR(retPtr);

                return ret;
            }
        }
    }
}
