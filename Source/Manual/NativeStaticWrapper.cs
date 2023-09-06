using System.Runtime.InteropServices;
using static ComWrappersTest.Source.Constants;

namespace ComWrappersTest.Source.Manual;

public unsafe class NativeStaticWrapper : IWshShell, IDisposable
{
    bool _isDisposed = false;

    nint _interfacePtr;

    public NativeStaticWrapper(nint ptr)
    {
        var guid = typeof(IWshShell).GUID;
        int hr = Marshal.QueryInterface(ptr, ref guid, out nint interfacePtr);
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
        var func = (delegate* unmanaged<nint, nint, nint, nint, int*, void>)(
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

    static nint VariantToPtr(VARIANT str)
    {
        var ptr = Marshal.AllocHGlobal(Marshal.SizeOf(str));
        Marshal.StructureToPtr(str, ptr, false);
        return ptr;
    }

    public string ExpandEnvironmentStrings(string Src)
    {
        var func = (delegate* unmanaged<void*, nint, nint*, void>)(
            *(
                *(void***)_interfacePtr + 12 /* IWshShell.ExpandEnvironmentStrings slot */
            )
        );

        nint retPtr;
        func((void*)_interfacePtr, Marshal.StringToBSTR(Src), &retPtr);
        var ret = Marshal.PtrToStringBSTR(retPtr);

        return ret;
    }
}
