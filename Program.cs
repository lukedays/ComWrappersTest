using System.Collections;
using System.Diagnostics;
using System.Runtime.InteropServices;
using static ComWrappersTest.Constants;

namespace ComWrappersTest
{
    public class Program
    {
        public static void Main()
        {
            // This is an adaptation of https://github.com/dotnet/samples/tree/main/core/interop/comwrappers/Tutorial
            // Intro to COM https://learn.microsoft.com/en-us/windows/win32/com/com-technical-overview

            // The CLSID for WScript.Shell (COMView.exe -> CLSID table)
            var clsid = new Guid("{F935DC22-1CF0-11D0-ADB9-00C04FD58A0B}");

            // The IID for IWshShell (COMView.exe -> TypeLib -> IWshShell interface)
            var iid = IWshShell.IID;

            // COMView.exe -> CLSID table -> Type column -> InProcServer32
            var server = CLSCTX.CLSCTX_INPROC_SERVER;

            // Create pointer to COM object with wrapper. Technically, a RCW
            var shell = ActivateClass<IWshShell>(clsid, iid, server);

            // Call COM functions. Since it's very tedious, we can use source generators to have more functions.
            shell.Run("calc.exe", 1, true);

            Console.WriteLine(shell.ExpandEnvironmentStrings("%USERPROFILE%"));
        }

        [DllImport("ole32.dll")]
        static extern int CoCreateInstance(
            ref Guid rclsid,
            IntPtr pUnkOuter,
            uint dwClsContext,
            ref Guid riid,
            out IntPtr ppv
        );

        public static I ActivateClass<I>(Guid clsid, Guid iid, CLSCTX server)
        {
            int hr = CoCreateInstance(
                ref clsid,
                IntPtr.Zero,
                (uint)server,
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
        public static Guid IID = new("F935DC21-1CF0-11D0-ADB9-00C04FD58A0B");

        int Run(string Command, int WindowStyle, bool WaitOnReturn);
        string ExpandEnvironmentStrings(string Src);
    }

    unsafe class MyComWrappers : ComWrappers
    {
        protected override unsafe ComInterfaceEntry* ComputeVtables(
            object obj,
            CreateComInterfaceFlags flags,
            out int count
        )
        {
            // This is only used for CCWs
            throw new NotImplementedException();
        }

        protected override object CreateObject(IntPtr externalComObject, CreateObjectFlags flags)
        {
            Debug.Assert(flags.HasFlag(CreateObjectFlags.UniqueInstance));

            return new ABI.NativeStaticWrapper(externalComObject);
        }

        protected override void ReleaseObjects(IEnumerable objects)
        {
            throw new NotImplementedException();
        }
    }

    namespace ABI
    {
        public unsafe class NativeStaticWrapper : IWshShell, IDisposable
        {
            bool _isDisposed = false;

            IntPtr _interfacePtr;

            public NativeStaticWrapper(IntPtr ptr)
            {
                int hr = Marshal.QueryInterface(ptr, ref IWshShell.IID, out IntPtr interfacePtr);
                if (hr != (int)HRESULT.S_OK)
                {
                    throw new NotSupportedException();
                }

                _interfacePtr = interfacePtr;
            }

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

                Marshal.Release(_interfacePtr);

                _isDisposed = true;
            }

            // The slot and parameters for the functions were found using COMView.exe and the C/C++ signature
            // Starting from 0 on the interface view, we can find the function slot numbers
            public int Run(string Command, int WindowStyle, bool WaitOnReturn)
            {
                var func = (delegate* unmanaged<IntPtr, IntPtr, IntPtr, IntPtr, int*, void>)(
                    *(
                        *(void***)_interfacePtr + 9 /* IWshShell.Run slot */
                    )
                );

                var arg1 = new VARIANT { vt = VARENUM.VT_I4, lVal = WindowStyle };

                var arg2 = new VARIANT
                {
                    vt = VARENUM.VT_BOOL,
                    boolVal = WaitOnReturn ? VARIANT_BOOL.VARIANT_TRUE : VARIANT_BOOL.VARIANT_FALSE
                };

                int retPtr;
                func(
                    _interfacePtr,
                    Marshal.StringToBSTR(Command),
                    VariantToPtr(arg1),
                    VariantToPtr(arg2),
                    &retPtr
                );

                return retPtr;
            }

            static IntPtr VariantToPtr(VARIANT str)
            {
                var ptr = Marshal.AllocHGlobal(Marshal.SizeOf(str));
                Marshal.StructureToPtr(str, ptr, false);
                return ptr;
            }

            public string ExpandEnvironmentStrings(string Src)
            {
                var func = (delegate* unmanaged<IntPtr, IntPtr, IntPtr*, void>)(
                    *(
                        *(void***)_interfacePtr + 12 /* IWshShell.ExpandEnvironmentStrings slot */
                    )
                );

                IntPtr retPtr;
                func(_interfacePtr, Marshal.StringToBSTR(Src), &retPtr);
                var ret = Marshal.PtrToStringBSTR(retPtr);

                return ret;
            }
        }
    }
}
